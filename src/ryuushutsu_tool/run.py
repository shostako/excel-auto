#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
流出不良集計ツール - 起動スクリプト
Windows環境から実行する場合はこのファイルを実行
"""

import sys
from pathlib import Path

# パッケージのパスを追加
sys.path.insert(0, str(Path(__file__).parent.parent))

from ryuushutsu_tool.gui import run_gui

if __name__ == "__main__":
    run_gui()
