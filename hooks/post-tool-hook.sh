#!/bin/bash
# PostToolUse用フック
# git commit/push, npm test/build 等を検出してPROGRESS.mdを更新
#
# 動的ディレクトリ検出: JSONのcwdフィールドから作業ディレクトリを取得

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROGRESS_UPDATER="$SCRIPT_DIR/progress-updater.py"

# stdinからJSONを読み込み
INPUT=$(cat)

# コマンド、終了コード、cwdを抽出（jqがなければpythonで）
if command -v jq &> /dev/null; then
    COMMAND=$(echo "$INPUT" | jq -r '.tool_input.command // empty')
    EXIT_CODE=$(echo "$INPUT" | jq -r '.tool_output.exit_code // 0')
    PROJECT_DIR=$(echo "$INPUT" | jq -r '.cwd // empty')
else
    # jqがない場合はpythonで処理（タブ区切りで出力）
    PARSED=$(python3 -c "
import json
try:
    data = json.loads('''$INPUT''')
    cmd = data.get('tool_input', {}).get('command', '')
    exit_code = str(data.get('tool_output', {}).get('exit_code', 0))
    cwd = data.get('cwd', '')
    print(cmd + '\t' + exit_code + '\t' + cwd)
except:
    print('\t0\t')
")
    COMMAND="${PARSED%%	*}"
    REST="${PARSED#*	}"
    EXIT_CODE="${REST%%	*}"
    PROJECT_DIR="${REST#*	}"
fi

# cwdが取得できなかった場合はスキップ
if [ -z "$PROJECT_DIR" ]; then
    exit 0
fi

# 対象コマンドかチェック
SHOULD_UPDATE=false
EVENT_TYPE="PostToolUse"

# git commit/push（成功時のみ）
if [[ "$COMMAND" =~ ^git\ (commit|push) ]] && [[ "$EXIT_CODE" == "0" ]]; then
    SHOULD_UPDATE=true
    EVENT_TYPE="PostToolUse:git"
fi

# npm test/build（成功時のみ）
if [[ "$COMMAND" =~ ^npm\ (test|run\ test|run\ build|build) ]] && [[ "$EXIT_CODE" == "0" ]]; then
    SHOULD_UPDATE=true
    EVENT_TYPE="PostToolUse:npm"
fi

# 対象コマンドでなければ何もしない
if [ "$SHOULD_UPDATE" != "true" ]; then
    exit 0
fi

# PROGRESS.mdを更新
ENHANCED_INPUT=$(python3 -c "
import json
import sys
data = json.loads('''$INPUT''' or '{}')
data['event'] = '$EVENT_TYPE'
data['matched_command'] = '''$COMMAND'''
print(json.dumps(data))
")

echo "$ENHANCED_INPUT" | python3 "$PROGRESS_UPDATER" "$PROJECT_DIR"
