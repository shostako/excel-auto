# データ集計モジュール
"""
データの変換・フィルタ・集計を行う
"""

import pandas as pd
from datetime import date
from typing import Optional, List

from .config import (
    get_hassei,
    get_hakken2,
    get_al_noa,
    get_fr_rr,
    get_lh_rh,
    get_mode2,
)
from .database import fetch_data


def add_derived_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    派生カラムを追加

    追加カラム:
    - 発生: 番号から判定
    - 発見2: 発見から判定
    - アル/ノア: 品番から判定
    - Fr/Rr: 品番から判定
    - LH/RH: 品番から判定
    - モード2: モードから派生（先頭の「裏」を削除）
    """
    df = df.copy()

    # 派生カラム追加
    df["発生"] = df["番号"].apply(get_hassei)
    df["発見2"] = df["発見"].apply(get_hakken2)
    df["アル/ノア"] = df["品番"].apply(get_al_noa)
    df["Fr/Rr"] = df["品番"].apply(get_fr_rr)
    df["LH/RH"] = df["品番"].apply(get_lh_rh)

    # モード2（モードカラムが存在する場合）
    if "モード" in df.columns:
        df["モード2"] = df["モード"].apply(get_mode2)

    return df


def filter_data(
    df: pd.DataFrame,
    hassei: Optional[str] = None,
    hakken2_list: Optional[List[str]] = None,
    mode2: Optional[str] = None,
) -> pd.DataFrame:
    """
    データをフィルタリング

    Parameters
    ----------
    df : pd.DataFrame
        フィルタ対象データ
    hassei : str, optional
        発生フィルタ（成形/塗装/モール/加工）
    hakken2_list : List[str], optional
        発見2フィルタ（複数選択可）
    mode2 : str, optional
        モード2フィルタ

    Returns
    -------
    pd.DataFrame
        フィルタ後のデータ
    """
    df = df.copy()

    if hassei:
        df = df[df["発生"] == hassei]

    if hakken2_list:
        df = df[df["発見2"].isin(hakken2_list)]

    if mode2:
        # モード2カラムが存在する場合のみフィルタ
        if "モード2" in df.columns:
            df = df[df["モード2"] == mode2]

    return df


def aggregate_by_zone(df: pd.DataFrame) -> pd.DataFrame:
    """
    ゾーン別に集計

    Returns
    -------
    pd.DataFrame
        集計結果（アル/ノア, Fr/Rr, LH/RH, ゾーン別の数量合計）
    """
    # グループ化して集計
    grouped = df.groupby(
        ["アル/ノア", "Fr/Rr", "LH/RH", "ゾーン"],
        as_index=False
    ).agg({
        "数量": "sum"
    })

    return grouped


def get_aggregated_data(
    start_date: date,
    end_date: date,
    hassei: str,
    hakken2_list: Optional[List[str]] = None,
    mode2: Optional[str] = None,
    db_path: Optional[str] = None,
) -> pd.DataFrame:
    """
    データ取得→派生カラム追加→フィルタ→集計 の一連の処理

    Parameters
    ----------
    start_date : date
        開始日
    end_date : date
        終了日
    hassei : str
        発生フィルタ
    hakken2_list : List[str], optional
        発見2フィルタ（複数選択可）
    mode2 : str, optional
        モード2フィルタ
    db_path : str, optional
        DBパス（指定なしはデフォルト）

    Returns
    -------
    pd.DataFrame
        集計結果
    """
    # データ取得
    if db_path:
        df = fetch_data(db_path, start_date, end_date)
    else:
        df = fetch_data(start_date=start_date, end_date=end_date)

    # 派生カラム追加
    df = add_derived_columns(df)

    # フィルタ
    df = filter_data(df, hassei=hassei, hakken2_list=hakken2_list, mode2=mode2)

    # 集計
    result = aggregate_by_zone(df)

    return result


def pivot_for_chart(aggregated_df: pd.DataFrame, al_noa: str, fr_rr: str) -> pd.DataFrame:
    """
    グラフ用にピボット変換

    Parameters
    ----------
    aggregated_df : pd.DataFrame
        集計済みデータ
    al_noa : str
        アル/ノアフィルタ（アルヴェル/ノアヴォク）
    fr_rr : str
        Fr/Rrフィルタ

    Returns
    -------
    pd.DataFrame
        ピボット済みデータ（行: ゾーン, 列: LH/RH）
    """
    # フィルタ
    df = aggregated_df[
        (aggregated_df["アル/ノア"] == al_noa) &
        (aggregated_df["Fr/Rr"] == fr_rr)
    ].copy()

    if df.empty:
        # 空の場合はゾーンA-Eの空データを返す
        zones = ["A", "B", "C", "D", "E"]
        return pd.DataFrame({
            "ゾーン": zones,
            "LH": [0] * 5,
            "RH": [0] * 5,
        }).set_index("ゾーン")

    # ピボット
    pivot = df.pivot_table(
        index="ゾーン",
        columns="LH/RH",
        values="数量",
        fill_value=0,
        aggfunc="sum"
    )

    # ゾーンA-Eを確保
    zones = ["A", "B", "C", "D", "E"]
    for zone in zones:
        if zone not in pivot.index:
            pivot.loc[zone] = 0

    # LH/RH列を確保
    for col in ["LH", "RH"]:
        if col not in pivot.columns:
            pivot[col] = 0

    # 並び替え
    pivot = pivot.reindex(zones)
    pivot = pivot[["LH", "RH"]]

    return pivot


def get_max_zone(pivot_df: pd.DataFrame, lh_rh: str) -> str:
    """
    最大値のゾーンを取得

    Parameters
    ----------
    pivot_df : pd.DataFrame
        ピボット済みデータ
    lh_rh : str
        LH or RH

    Returns
    -------
    str
        最大値のゾーン（A-E）、全て0の場合は空文字
    """
    if lh_rh not in pivot_df.columns:
        return ""

    # 数値型に変換（エラー時は0）
    col = pd.to_numeric(pivot_df[lh_rh], errors="coerce").fillna(0)

    if col.max() == 0:
        return ""

    return col.idxmax()
