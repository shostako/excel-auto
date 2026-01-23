# データベース接続モジュール
"""
Access DBからデータを取得する
"""

import pandas as pd
import pyodbc
from pathlib import Path
from datetime import date
from typing import Optional

from .config import DB_PATH, TABLE_NAME


def get_connection_string(db_path: str = DB_PATH) -> str:
    """ODBC接続文字列を生成"""
    # 64bit版 Access ODBC ドライバー
    driver = "{Microsoft Access Driver (*.mdb, *.accdb)}"
    return f"Driver={driver};DBQ={db_path};"


def fetch_bangou_table(db_path: str = DB_PATH) -> pd.DataFrame:
    """
    _番号テーブルを取得（番号→モードの紐付け）

    Returns
    -------
    pd.DataFrame
        番号テーブル
    """
    conn_str = get_connection_string(db_path)
    sql = "SELECT * FROM [_番号]"

    try:
        with pyodbc.connect(conn_str) as conn:
            df = pd.read_sql(sql, conn)
        return df
    except pyodbc.Error as e:
        raise ConnectionError(f"番号テーブル取得エラー: {e}")


def fetch_data(
    db_path: str = DB_PATH,
    start_date: Optional[date] = None,
    end_date: Optional[date] = None,
) -> pd.DataFrame:
    """
    Accessからデータを取得（_番号テーブルとJOIN済み）

    Parameters
    ----------
    db_path : str
        データベースファイルパス
    start_date : date, optional
        開始日（指定時は日付フィルタ適用）
    end_date : date, optional
        終了日（指定時は日付フィルタ適用）

    Returns
    -------
    pd.DataFrame
        取得したデータ（モード列含む）
    """
    conn_str = get_connection_string(db_path)

    # メインデータ取得
    sql = f"SELECT * FROM [{TABLE_NAME}]"
    params = []

    if start_date and end_date:
        sql += " WHERE 日付 >= ? AND 日付 <= ?"
        params = [start_date, end_date]
    elif start_date:
        sql += " WHERE 日付 >= ?"
        params = [start_date]
    elif end_date:
        sql += " WHERE 日付 <= ?"
        params = [end_date]

    try:
        with pyodbc.connect(conn_str) as conn:
            if params:
                df = pd.read_sql(sql, conn, params=params)
            else:
                df = pd.read_sql(sql, conn)

            # 番号テーブル取得
            bangou_df = pd.read_sql("SELECT * FROM [_番号]", conn)

        # 番号でJOIN（モード列を追加）
        if "番号" in df.columns and "番号" in bangou_df.columns:
            # 番号を文字列に統一してマージ
            df["番号_str"] = df["番号"].astype(str).str.strip().str.upper()
            bangou_df["番号_str"] = bangou_df["番号"].astype(str).str.strip().str.upper()

            df = df.merge(
                bangou_df[["番号_str", "モード"]],
                on="番号_str",
                how="left"
            )
            df.drop(columns=["番号_str"], inplace=True)

        return df
    except pyodbc.Error as e:
        raise ConnectionError(f"データベース接続エラー: {e}")


def test_connection(db_path: str = DB_PATH) -> bool:
    """
    データベース接続テスト

    Returns
    -------
    bool
        接続成功ならTrue
    """
    conn_str = get_connection_string(db_path)
    try:
        with pyodbc.connect(conn_str) as conn:
            cursor = conn.cursor()
            cursor.execute(f"SELECT TOP 1 * FROM [{TABLE_NAME}]")
            cursor.fetchone()
        return True
    except Exception as e:
        print(f"接続テスト失敗: {e}")
        return False
