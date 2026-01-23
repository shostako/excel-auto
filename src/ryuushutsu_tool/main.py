#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
流出不良集計ツール - メインエントリーポイント
"""

import sys
from pathlib import Path

# パッケージのルートをパスに追加（exe化対策）
if getattr(sys, 'frozen', False):
    # PyInstallerでexe化された場合
    BASE_DIR = Path(sys._MEIPASS)
else:
    BASE_DIR = Path(__file__).parent


def main():
    """メイン関数"""
    from .gui import run_gui
    run_gui()


if __name__ == "__main__":
    main()
