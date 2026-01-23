# 設定ファイル
"""
アプリケーション設定
"""

import os
from pathlib import Path
from datetime import datetime

# データベース設定
DB_PATH = r"Z:\全社共有\オート事業部\日報\不良集計\不良集計表\2026年\不良調査表DB-2026.accdb"
TABLE_NAME = "_不良集計ゾーン別"

# 出力設定
OUTPUT_DIR = Path.home() / "Desktop"
OUTPUT_FILENAME_TEMPLATE = "流出集計_{timestamp}.xlsm"

def get_output_path() -> Path:
    """出力ファイルパスを生成"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = OUTPUT_FILENAME_TEMPLATE.format(timestamp=timestamp)
    return OUTPUT_DIR / filename

# リソースパス
RESOURCES_DIR = Path(__file__).parent / "resources"
IMAGES_DIR = RESOURCES_DIR / "images"

# 画像ファイル名マッピング
IMAGE_FILES = {
    ("アルヴェル", "Fr", "LH"): "アルヴェルFrLH.png",
    ("アルヴェル", "Fr", "RH"): "アルヴェルFrRH.png",
    ("アルヴェル", "Rr", "LH"): "アルヴェルRrLH.png",
    ("アルヴェル", "Rr", "RH"): "アルヴェルRrRH.png",
    ("ノアヴォク", "Fr", "LH"): "ノアヴォクFrLH.png",
    ("ノアヴォク", "Fr", "RH"): "ノアヴォクFrRH.png",
    ("ノアヴォク", "Rr", "LH"): "ノアヴォクRrLH.png",
    ("ノアヴォク", "Rr", "RH"): "ノアヴォクRrRH.png",
}

def get_image_path(al_noa: str, fr_rr: str, lh_rh: str) -> Path:
    """画像ファイルパスを取得"""
    key = (al_noa, fr_rr, lh_rh)
    if key in IMAGE_FILES:
        return IMAGES_DIR / IMAGE_FILES[key]
    raise ValueError(f"画像が見つかりません: {key}")

# 派生カラムのマッピング
# 発見 → 発見2
HAKKEN2_MAP = {
    "S": "成形",
    "T": "塗装",
    "M": "モール",
    "K": "加工",
}

# 番号 → 発生
# 成形: 1-20, X, Y, Z
# 塗装: 21-40
# モール: 41-50
# 加工: A-W (X,Y,Z以外のアルファベット)
SEIKEI_NUMBERS = set(str(i) for i in range(1, 21)) | {"X", "Y", "Z"}
TOSOU_NUMBERS = set(str(i) for i in range(21, 41))
MOULD_NUMBERS = set(str(i) for i in range(41, 51))
KAKOU_LETTERS = set("ABCDEFGHIJKLMNOPQRSTUVW")  # X,Y,Z以外

def get_hassei(bangou) -> str:
    """番号から発生を判定"""
    if bangou is None:
        return "不明"

    # 数値の場合は文字列に変換
    bangou_str = str(bangou).strip().upper()

    # 数値として判定を試みる
    try:
        bangou_int = int(float(bangou_str))  # "1.0" などにも対応
        if 1 <= bangou_int <= 20:
            return "成形"
        elif 21 <= bangou_int <= 40:
            return "塗装"
        elif 41 <= bangou_int <= 50:
            return "モール"
    except (ValueError, TypeError):
        pass

    # 文字列として判定
    if bangou_str in {"X", "Y", "Z"}:
        return "成形"
    elif bangou_str in KAKOU_LETTERS:
        return "加工"

    return "不明"

def get_hakken2(hakken: str) -> str:
    """発見から発見2を判定"""
    if hakken is None:
        return "不明"
    hakken = str(hakken).strip().upper()
    return HAKKEN2_MAP.get(hakken, "不明")

def get_al_noa(hinban: str) -> str:
    """品番からアル/ノアを判定（先頭2文字）"""
    if hinban is None or len(hinban) < 2:
        return "不明"
    prefix = hinban[:2]
    if prefix == "アル":
        return "アルヴェル"
    elif prefix == "ノア":
        return "ノアヴォク"
    else:
        return "不明"

def get_fr_rr(hinban: str) -> str:
    """品番からFr/Rrを判定（3-4文字目）"""
    if hinban is None or len(hinban) < 4:
        return "不明"
    fr_rr = hinban[2:4]
    if fr_rr in ("Fr", "Rr"):
        return fr_rr
    else:
        return "不明"

def get_lh_rh(hinban: str) -> str:
    """品番からLH/RHを判定（末尾2文字）"""
    if hinban is None or len(hinban) < 2:
        return "不明"
    suffix = hinban[-2:]
    if suffix in ("LH", "RH"):
        return suffix
    else:
        return "不明"


def get_mode2(mode: str) -> str:
    """モードからモード2を判定（先頭の「裏」を削除）"""
    if mode is None:
        return ""
    mode = str(mode).strip()
    if mode.startswith("裏"):
        return mode[1:]  # 「裏」を削除
    return mode
