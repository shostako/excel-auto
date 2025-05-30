#!/bin/bash

# src/からready/にファイルを移動する補助スクリプト

PROJECT_DIR="/home/shostako/ClaudeCode/ExcelAuto"
SRC_DIR="$PROJECT_DIR/src"
READY_DIR="$PROJECT_DIR/ready"

# 色付きメッセージ
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

# 引数チェック
if [ $# -eq 0 ]; then
    echo "使い方: ./stage.sh ファイル名.bas"
    echo "または: ./stage.sh all  (全.basファイルをステージング)"
    echo ""
    echo "src/内のファイル:"
    ls -1 "$SRC_DIR"/*.bas 2>/dev/null | xargs -n1 basename
    exit 1
fi

# 全ファイルステージング
if [ "$1" == "all" ]; then
    count=0
    for file in "$SRC_DIR"/*.bas; do
        if [ -f "$file" ]; then
            basename=$(basename "$file")
            mv "$file" "$READY_DIR/"
            echo -e "${GREEN}✓ ステージング: $basename${NC}"
            ((count++))
        fi
    done
    echo ""
    echo "合計 $count ファイルをready/に移動しました"
    exit 0
fi

# 個別ファイルステージング
FILE="$1"
FULL_PATH="$SRC_DIR/$FILE"

if [ ! -f "$FULL_PATH" ]; then
    echo -e "${RED}エラー: ファイルが見つかりません: $FILE${NC}"
    echo ""
    echo "src/内のファイル:"
    ls -1 "$SRC_DIR"/*.bas 2>/dev/null | xargs -n1 basename
    exit 1
fi

# ファイル移動
mv "$FULL_PATH" "$READY_DIR/"
echo -e "${GREEN}✓ ステージング完了: $FILE${NC}"
echo "  → ready/に移動しました（自動変換待ち）"