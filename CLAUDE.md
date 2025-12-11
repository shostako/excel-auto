# excel-auto プロジェクト設定

## 概要

Excelマクロ（VBA）開発プロジェクト。

## ルール読み込み

詳細ルールは`.claude/rules/`ディレクトリを参照：

| ファイル | 内容 |
|---------|------|
| 01-session-workflow.md | セッション開始/終了ルール |
| 02-monday-behavior.md | Mondayの行動ルール |
| 03-logging.md | 作業ログルール |
| 04-subagent-planmode.md | サブエージェント・Plan Mode運用 |
| 10-vba-required-reading.md | 必読ファイル |
| 11-vba-coding-standards.md | VBAコーディング規約 |
| 12-vba-encoding.md | 文字エンコーディング管理 |
| 13-git-workflow.md | Git/Makefileフロー |

## クイックリファレンス

- **文字コード**: inboxはShift-JIS → iconvで変換必須
- **出力先**: src/(UTF-8) → bas2sjis → macros/(Shift-JIS)
- **禁止**: Activateメソッド使用禁止
- **セッション終了**: `/wrap`コマンドで一括処理

## 関連リソース

- ナレッジ: `docs/excel-knowledge/`
- ログ: `logs/`
- 進捗: `PROGRESS.md`
