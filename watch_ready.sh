#!/bin/bash

# ready/フォルダを監視して自動的にShift-JIS変換するスクリプト

# 設定
PROJECT_DIR="/home/shostako/ClaudeCode/ExcelAuto"
READY_DIR="$PROJECT_DIR/ready"
WINDOWS_OUTPUT="/mnt/c/Users/shost/Documents/Excelマクロ"
CONVERT_SCRIPT="/home/shostako/ClaudeCode/convert_bas_to_sjis.sh"
LOG_FILE="$PROJECT_DIR/auto_convert.log"

# 色付きメッセージ
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

echo -e "${GREEN}=== ExcelAuto 自動変換監視 ===${NC}"
echo "監視対象: $READY_DIR"
echo "出力先: $WINDOWS_OUTPUT"
echo ""
echo "ready/フォルダに.basファイルを移動すると自動的に変換されます"
echo "終了: Ctrl+C"
echo ""

# inotify-toolsのインストール確認
if ! command -v inotifywait &> /dev/null; then
    echo -e "${RED}エラー: inotify-toolsがインストールされていません${NC}"
    echo "インストールコマンド: sudo apt-get install inotify-tools"
    exit 1
fi

# ログ関数
log_message() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] $1" >> "$LOG_FILE"
}

# 変換関数
convert_and_move() {
    local file="$1"
    local basename=$(basename "$file")
    
    echo -e "${YELLOW}検出: $basename${NC}"
    
    # 少し待つ（ファイル書き込み完了を待つ）
    sleep 0.5
    
    # 変換実行
    if "$CONVERT_SCRIPT" "$file" 2>&1; then
        echo -e "${GREEN}✓ 変換成功: $basename${NC}"
        log_message "SUCCESS: $basename"
        
        # 変換成功したらreadyからsrcに移動（バックアップとして）
        mv "$file" "$PROJECT_DIR/src/"
        echo -e "  → src/に移動しました"
    else
        echo -e "${RED}✗ 変換失敗: $basename${NC}"
        log_message "FAILED: $basename"
    fi
    
    echo ""
}

# 起動時に既存ファイルを処理
echo "既存ファイルをチェック中..."
for file in "$READY_DIR"/*.bas; do
    if [ -f "$file" ]; then
        convert_and_move "$file"
    fi
done

# ファイル監視開始
log_message "監視開始"
echo -e "${GREEN}監視を開始しました${NC}"
echo ""

# inotifywaitで監視（作成、移動を検出）
inotifywait -m -e create -e moved_to --format '%w%f' "$READY_DIR" | while read file; do
    # .basファイルのみ処理
    if [[ "$file" == *.bas ]]; then
        convert_and_move "$file"
    fi
done