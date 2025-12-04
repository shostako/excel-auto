#!/usr/bin/env python3
"""
PROGRESS.md自動更新スクリプト
フックイベント発火時にPROGRESS.mdのメタ情報を更新する

Usage:
    echo '{"event": "SessionEnd", "reason": "clear"}' | python3 progress-updater.py /path/to/project

イベントタイプ:
    - SessionEnd: セッション終了時
    - PreCompact: コンパクト前
    - PostToolUse: ツール使用後（git commit/push等）
    - SessionStart: セッション開始時（経過時間チェック）
"""

import sys
import json
import os
import subprocess
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

# 日本時間（JST）を明示的に指定
JST = ZoneInfo("Asia/Tokyo")

# 重複発火防止用の最小間隔（秒）
MIN_UPDATE_INTERVAL = 60

def find_progress_md(start_path: str) -> Path | None:
    """PROGRESS.mdを探す（カレントディレクトリから上位へ探索）"""
    current = Path(start_path).resolve()
    while current != current.parent:
        progress_path = current / "PROGRESS.md"
        if progress_path.exists():
            return progress_path
        current = current.parent
    return None

def get_last_update_time(content: str) -> datetime | None:
    """PROGRESS.mdから最終更新日時を抽出（JSTとして解釈）"""
    match = re.search(r'\*\*最終更新\*\*:\s*(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2})', content)
    if match:
        try:
            naive_dt = datetime.strptime(match.group(1), '%Y-%m-%d %H:%M')
            return naive_dt.replace(tzinfo=JST)  # JSTとして解釈
        except ValueError:
            return None
    return None

def get_recent_git_commits(repo_path: str, count: int = 2) -> list[str]:
    """直近のGitコミットを取得"""
    try:
        result = subprocess.run(
            ['git', 'log', f'-{count}', '--oneline'],
            cwd=repo_path,
            capture_output=True,
            text=True,
            timeout=5
        )
        if result.returncode == 0:
            return [line.strip() for line in result.stdout.strip().split('\n') if line.strip()]
    except (subprocess.TimeoutExpired, FileNotFoundError):
        pass
    return []

def update_progress_md(progress_path: Path, event_type: str, event_data: dict) -> dict:
    """PROGRESS.mdを更新"""
    content = progress_path.read_text(encoding='utf-8')
    now = datetime.now(JST)
    now_str = now.strftime('%Y-%m-%d %H:%M')

    # 重複発火チェック
    last_update = get_last_update_time(content)
    if last_update and (now - last_update).total_seconds() < MIN_UPDATE_INTERVAL:
        return {
            "status": "skipped",
            "reason": "recent_update",
            "last_update": last_update.strftime('%Y-%m-%d %H:%M')
        }

    # 最終更新日時を更新
    new_content = re.sub(
        r'(\*\*最終更新\*\*:\s*)\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}',
        f'\\g<1>{now_str}',
        content
    )

    # 直近のGitコミットを更新
    repo_path = progress_path.parent
    commits = get_recent_git_commits(str(repo_path))
    if commits:
        commit_lines = '\n'.join([f'- {commit}' for commit in commits])
        # ## 直近のGitコミット セクションを更新
        new_content = re.sub(
            r'(## 直近のGitコミット\n).*?(?=\n## |\Z)',
            f'\\g<1>{commit_lines}\n',
            new_content,
            flags=re.DOTALL
        )

    # 変更があれば書き込み
    if new_content != content:
        progress_path.write_text(new_content, encoding='utf-8')
        return {
            "status": "updated",
            "event": event_type,
            "timestamp": now_str,
            "commits_updated": len(commits)
        }

    return {"status": "no_change"}

def check_elapsed_time(progress_path: Path, threshold_hours: int = 24) -> dict:
    """SessionStart時の経過時間チェック"""
    content = progress_path.read_text(encoding='utf-8')
    last_update = get_last_update_time(content)

    if not last_update:
        return {
            "warning": True,
            "message": "PROGRESS.mdの最終更新日時が見つかりません。更新を推奨します。"
        }

    elapsed = datetime.now(JST) - last_update
    if elapsed > timedelta(hours=threshold_hours):
        return {
            "warning": True,
            "message": f"PROGRESS.mdが{elapsed.days}日{elapsed.seconds // 3600}時間更新されていません。確認・更新を推奨します。",
            "last_update": last_update.strftime('%Y-%m-%d %H:%M')
        }

    return {"warning": False}

def main():
    # 引数チェック
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: progress-updater.py <project_path>"}))
        sys.exit(1)

    project_path = sys.argv[1]

    # stdinからイベントデータを読み込み
    try:
        event_data = json.load(sys.stdin)
    except json.JSONDecodeError:
        event_data = {}

    event_type = event_data.get('event', 'Unknown')

    # PROGRESS.mdを探す
    progress_path = find_progress_md(project_path)
    if not progress_path:
        print(json.dumps({"error": "PROGRESS.md not found", "search_start": project_path}))
        sys.exit(0)  # エラーでも終了コード0（フックをブロックしない）

    # イベントタイプに応じた処理
    if event_type == 'SessionStart':
        result = check_elapsed_time(progress_path)
        if result.get('warning'):
            # 警告メッセージをstdoutに出力（Claudeへのフィードバック）
            print(json.dumps({
                "type": "progress_reminder",
                **result
            }))
    else:
        # SessionEnd, PreCompact, PostToolUse等
        result = update_progress_md(progress_path, event_type, event_data)
        print(json.dumps(result))

if __name__ == '__main__':
    main()
