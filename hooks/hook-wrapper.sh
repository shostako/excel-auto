#!/bin/bash
# フック共通ラッパー
# Usage: hook-wrapper.sh <event_type>
# stdin: フックイベントのJSONデータ
#
# 動的ディレクトリ検出: JSONのcwdフィールドから作業ディレクトリを取得し、
# そのプロジェクトのPROGRESS.mdを更新する

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROGRESS_UPDATER="$SCRIPT_DIR/progress-updater.py"

EVENT_TYPE="${1:-Unknown}"

# stdinからJSONを読み込み
INPUT=$(cat)

if [ -z "$INPUT" ]; then
    INPUT="{}"
fi

# JSONからcwdを抽出し、イベントタイプを追加
if command -v jq &> /dev/null; then
    PROJECT_DIR=$(echo "$INPUT" | jq -r '.cwd // empty')
    ENHANCED_INPUT=$(echo "$INPUT" | jq --arg event "$EVENT_TYPE" '. + {event: $event}')
else
    # jqがない場合はpythonで処理（タブ区切りで出力）
    PARSED=$(python3 -c "
import json
data = json.loads('''$INPUT''' or '{}')
cwd = data.get('cwd', '')
data['event'] = '$EVENT_TYPE'
print(cwd + '\t' + json.dumps(data))
")
    PROJECT_DIR="${PARSED%%	*}"
    ENHANCED_INPUT="${PARSED#*	}"
fi

# cwdが取得できなかった場合はスキップ（フォールバックなし）
if [ -z "$PROJECT_DIR" ]; then
    echo '{"status": "skipped", "reason": "no_cwd_in_input"}'
    exit 0
fi

# progress-updater.pyを実行
echo "$ENHANCED_INPUT" | python3 "$PROGRESS_UPDATER" "$PROJECT_DIR"
