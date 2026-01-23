# Excel出力モジュール（win32com + テンプレート方式）
"""
テンプレートExcelにデータを書き込み、グラフ軸を統一設定する
"""

import shutil
import pandas as pd
from pathlib import Path
from datetime import date
from typing import Dict, List
import math

import win32com.client as win32

from .config import get_output_path
from .aggregator import pivot_for_chart


# テンプレート設定
TEMPLATE_FILENAME = "閲覧_FrRrゾーン特化.xlsm"
SHEET_NAME = "ゾーンFrRr流出"

# テーブル名とグラフ名のマッピング
TABLE_CHART_MAP = {
    ("アルヴェル", "Fr"): {"table": "_アルヴェルFr", "chart": "グラフ1"},
    ("アルヴェル", "Rr"): {"table": "_アルヴェルRr", "chart": "グラフ2"},
    ("ノアヴォク", "Fr"): {"table": "_ノアヴォクFr", "chart": "グラフ3"},
    ("ノアヴォク", "Rr"): {"table": "_ノアヴォクRr", "chart": "グラフ4"},
}

# 品番列の値（テーブルごとに異なる）
HINBAN_MAP = {
    "Fr": ["Fr LH", "Fr RH"],
    "Rr": ["Rr LH", "Rr RH"],
}

# ゾーン
ZONES = ["A", "B", "C", "D", "E"]


def get_template_path() -> Path:
    """テンプレートファイルのパスを取得"""
    # 開発時: inbox
    inbox_path = Path(__file__).parent.parent.parent / "inbox" / TEMPLATE_FILENAME
    if inbox_path.exists():
        return inbox_path

    # 配布時: resources/templates
    resources_path = Path(__file__).parent / "resources" / "templates" / TEMPLATE_FILENAME
    if resources_path.exists():
        return resources_path

    raise FileNotFoundError(f"テンプレートが見つかりません: {TEMPLATE_FILENAME}")


def calc_nice_max_value(max_value: float) -> float:
    """
    データの最大値から「良い感じの」軸の最大値を計算
    VBAマクロのCalcNiceMaxValueを移植
    """
    if max_value <= 0:
        return 10

    min_target = max_value * 1.1
    max_target = max_value * 1.2

    magnitude = int(math.log10(max_target))
    base = 10 ** magnitude

    candidates = [1, 1.2, 1.5, 2, 2.5, 3, 4, 5, 6, 7, 8, 9, 10]

    for c in candidates:
        nice_value = c * base
        if nice_value >= min_target:
            return nice_value

    return max_target


def calc_nice_tick_interval(max_value: float) -> float:
    """
    軸の最大値に基づいて適切な目盛り間隔を計算
    VBAマクロのCalcNiceTickIntervalを移植
    """
    if max_value <= 0:
        return 1

    target_ticks = 6
    rough_interval = max_value / target_ticks

    if rough_interval <= 0:
        return 1

    magnitude = int(math.log10(rough_interval))
    base = 10 ** magnitude

    ratio = rough_interval / base
    if ratio <= 1:
        return base
    elif ratio <= 2:
        return 2 * base
    elif ratio <= 5:
        return 5 * base
    else:
        return 10 * base


def write_data_to_table(ws, table_name: str, pivot_df: pd.DataFrame, fr_rr: str):
    """
    テーブルにデータを書き込む（win32com版）

    Parameters
    ----------
    ws : win32com Worksheet
        ワークシート
    table_name : str
        テーブル名
    pivot_df : pd.DataFrame
        ピボット済みデータ（行: ゾーン, 列: LH/RH）
    fr_rr : str
        Fr または Rr
    """
    # テーブルを取得
    try:
        tbl = ws.ListObjects(table_name)
    except Exception as e:
        print(f"警告: テーブルが見つかりません: {table_name} ({e})")
        return

    # テーブルのデータ範囲を取得（ヘッダー除く）
    try:
        data_range = tbl.DataBodyRange
        if data_range is None:
            print(f"警告: テーブルにデータ範囲がありません: {table_name}")
            return
    except Exception as e:
        print(f"警告: データ範囲取得エラー: {table_name} ({e})")
        return

    # 各行を処理
    for row_idx in range(1, data_range.Rows.Count + 1):
        hinban = data_range.Cells(row_idx, 1).Value  # 品番列
        zone = data_range.Cells(row_idx, 2).Value    # ゾーン列

        if hinban is None or zone is None:
            continue

        # LH/RHを判定
        hinban_str = str(hinban)
        if "LH" in hinban_str:
            lh_rh = "LH"
        elif "RH" in hinban_str:
            lh_rh = "RH"
        else:
            continue

        # 対応する数量を取得
        try:
            qty = pivot_df.loc[zone, lh_rh] if zone in pivot_df.index else 0
        except KeyError:
            qty = 0

        # 数量を書き込み
        data_range.Cells(row_idx, 3).Value = int(qty) if pd.notna(qty) else 0


def apply_unified_axis(ws, chart_names: List[str], max_values: List[float]):
    """
    4つのグラフに統一された軸設定を適用（win32com版）

    Parameters
    ----------
    ws : win32com Worksheet
        ワークシート
    chart_names : List[str]
        グラフ名のリスト
    max_values : List[float]
        各グラフのデータ最大値
    """
    # 全体の最大値を決定
    overall_max = max(max_values) if max_values else 0

    # 良い感じの軸設定を計算
    axis_max = calc_nice_max_value(overall_max)
    tick_interval = calc_nice_tick_interval(axis_max)

    print(f"軸設定: 最大値={axis_max}, 目盛り間隔={tick_interval}")

    # xlValue = 2
    xlValue = 2

    # 各グラフに適用
    for chart_name in chart_names:
        try:
            chart_obj = ws.ChartObjects(chart_name)
            chart = chart_obj.Chart
            value_axis = chart.Axes(xlValue)

            value_axis.MaximumScaleIsAuto = False
            value_axis.MaximumScale = axis_max
            value_axis.MinimumScaleIsAuto = False
            value_axis.MinimumScale = 0
            value_axis.MajorUnitIsAuto = False
            value_axis.MajorUnit = tick_interval
            value_axis.MinorUnitIsAuto = False
            value_axis.MinorUnit = tick_interval / 2

        except Exception as e:
            print(f"警告: グラフ軸設定失敗 ({chart_name}): {e}")


def generate_excel(
    aggregated_df: pd.DataFrame,
    hassei: str,
    start_date: date,
    end_date: date,
    output_path: Path = None,
) -> Path:
    """
    テンプレートを使ってExcelファイルを生成（win32com版）

    Parameters
    ----------
    aggregated_df : pd.DataFrame
        集計済みデータ
    hassei : str
        発生フィルタ値（タイトル用）
    start_date : date
        開始日（タイトル用）
    end_date : date
        終了日（タイトル用）
    output_path : Path, optional
        出力パス（指定なしは自動生成）

    Returns
    -------
    Path
        出力ファイルパス
    """
    if output_path is None:
        output_path = get_output_path()

    # テンプレートをコピー
    template_path = get_template_path()
    shutil.copy2(template_path, output_path)
    print(f"テンプレートコピー: {template_path} -> {output_path}")

    # win32comでExcelを操作
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(str(output_path))
        ws = wb.Worksheets(SHEET_NAME)

        # データ書き込み & 最大値収集
        max_values = []
        chart_names = []

        for (al_noa, fr_rr), config in TABLE_CHART_MAP.items():
            table_name = config["table"]
            chart_name = config["chart"]
            chart_names.append(chart_name)

            # ピボットデータ取得
            pivot_df = pivot_for_chart(aggregated_df, al_noa, fr_rr)

            # テーブルにデータ書き込み
            write_data_to_table(ws, table_name, pivot_df, fr_rr)

            # 最大値を記録（軸統一用）
            try:
                numeric_values = pd.to_numeric(pivot_df.values.flatten(), errors='coerce')
                data_max = float(numeric_values.max()) if len(numeric_values) > 0 else 0
                if pd.isna(data_max):
                    data_max = 0
            except (TypeError, ValueError):
                data_max = 0
            max_values.append(data_max)

            print(f"データ書き込み完了: {table_name} (最大値: {data_max})")

        # グラフ軸を統一設定
        apply_unified_axis(ws, chart_names, max_values)

        # 保存して閉じる
        wb.Save()
        wb.Close()

    finally:
        excel.Quit()

    print(f"Excel出力完了: {output_path}")
    return output_path
