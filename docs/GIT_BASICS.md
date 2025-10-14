# Git基本操作ガイド（excel-auto用）

**対象**：Git初心者、最小限の操作で始めたい人向け

---

## はじめに

このガイドは、excel-autoプロジェクトでGitを使い始めるための最小限の知識をまとめたものだ。

**今すぐ覚えること**：たった3つのコマンド
**将来覚えること**：ブランチ、PR（今は不要）

---

## フェーズ1：今日から始める基本操作

### 1. 現在の状態を確認

```bash
cd /home/shostako/ClaudeCode/excel-auto
git status
```

**表示例**:
```
On branch main
Changes not staged for commit:
  modified:   src/m転記_日報_成形.bas

Untracked files:
  src/m転記_日報_塗装.bas
```

**意味**:
- `modified`: 既存ファイルが変更された
- `Untracked files`: 新しいファイルが追加された

---

### 2. ファイルをステージング（記録準備）

```bash
# 特定のファイルを追加
git add src/m転記_日報_成形.bas

# 複数ファイルを追加
git add src/*.bas

# 全ての変更を追加
git add -A
```

**ステージングとは**：「このファイルを次のコミットに含める」という宣言

---

### 3. コミット（記録）

```bash
git commit -m "add: 日報成形マクロ - 9分類集計対応"
```

**コミットメッセージの書き方**:
- `add:` - 新規ファイル追加
- `fix:` - バグ修正
- `update:` - 既存機能の改善
- `docs:` - ドキュメント更新
- `clean:` - ファイル削除・整理

**例**:
```bash
git commit -m "fix: RH系統集計バグ - Exit Forを削除"
git commit -m "update: 合計行の計算ロジック改善"
git commit -m "docs: 10月14日の作業ログ追加"
git commit -m "clean: ワークスペースクリーンアップ"
```

---

### 4. 履歴を確認

```bash
# 簡易表示
git log --oneline

# 詳細表示
git log

# 最近5件
git log --oneline -5

# 特定の文字列を含むコミット検索
git log --oneline --grep="RH系統"
```

**表示例**:
```
b96872e fix: RH系統集計バグ - Exit Forを削除
08a728d add: 日報塗装マクロ
ab6ac1f update: 合計行削除
```

---

### 5. 変更内容を確認

```bash
# まだコミットしていない変更を確認
git diff

# 特定のファイルの変更
git diff src/m転記_日報_成形.bas

# 過去のコミットの変更内容
git show b96872e
```

---

### 6. 過去のファイルを取り出す

```bash
# 過去のコミットからファイルを復元
git checkout b96872e -- src/m転記_日報_成形.bas

# 復元したファイルをコミット
git add src/m転記_日報_成形.bas
git commit -m "restore: 日報成形マクロを復元"
```

---

## 実践：マクロ作成の流れ

### ステップ1：マクロ作成
```bash
# Monday（Claude）がsrc/にマクロ作成
# bas2sjisでmacros/に変換
# Excel取り込み → テスト → OK
```

### ステップ2：コミット
```bash
cd /home/shostako/ClaudeCode/excel-auto

# 状態確認
git status

# ファイル追加
git add src/m転記_日報_成形.bas

# コミット
git commit -m "add: 日報成形マクロ - 9分類集計対応"
```

### ステップ3：作業ログ記録
```bash
# ログファイル編集後
git add logs/2025-10.md
git commit -m "docs: 10月14日の作業ログ追加"
```

### ステップ4：クリーンアップ
```bash
# セッション終了時
./scripts/cleanup-workspace.sh

# 削除確認 → y → 自動でコミット
```

---

## よくある質問

### Q1: コミットし忘れた、どうすればいい？

**A: 気づいた時点でコミットすればOK**

```bash
# 今からコミット
git add src/m転記_日報_成形.bas
git commit -m "add: 日報成形マクロ（作成日: 10/14）"
```

過去に遡ってコミットする必要はない。

---

### Q2: 間違えてコミットした、取り消したい

**A: 最後のコミットだけなら簡単に取り消せる**

```bash
# 最後のコミットを取り消し（ファイルは残る）
git reset --soft HEAD~1

# 再度コミット
git add src/m転記_日報_成形.bas
git commit -m "add: 正しいメッセージ"
```

**注意**: この操作は最後のコミットにのみ有効。

---

### Q3: ファイルを削除したけど、後で必要になった

**A: Git履歴から復元できる**

```bash
# 削除したファイルを含むコミットを探す
git log --oneline --all -- src/m転記_日報_成形.bas

# そのコミットからファイルを復元
git checkout b96872e -- src/m転記_日報_成形.bas
```

---

### Q4: 履歴が大きくなりすぎた

**A: 定期的に圧縮する**

```bash
# Git履歴の圧縮
git gc --aggressive --prune=now
```

---

## フェーズ2：将来のための知識（今は不要）

### ブランチの概念

```bash
# 実験用ブランチ作成
git checkout -b experiment/new-logic

# 好きなだけ変更
# ...

# 元に戻る
git checkout main

# 実験が成功したらマージ
git merge experiment/new-logic

# 失敗したらブランチ削除
git branch -D experiment/new-logic
```

**メリット**:
- 安定版を壊さずに実験できる
- 複数の案を並行して試せる

**今は不要な理由**:
- 基本操作に慣れてから

---

## トラブルシューティング

### エラー: "Please tell me who you are"

```bash
# 初回のみ設定
git config --global user.name "Your Name"
git config --global user.email "your@email.com"
```

---

### エラー: "fatal: not a git repository"

```bash
# プロジェクトルートに移動
cd /home/shostako/ClaudeCode/excel-auto

# または、Gitリポジトリ初期化（初回のみ）
git init
```

---

### 変更を全て破棄したい

```bash
# 全ての変更を破棄（注意：元に戻せない）
git reset --hard HEAD

# 特定のファイルのみ破棄
git checkout -- src/m転記_日報_成形.bas
```

---

## まとめ

**今日から使う3つのコマンド**:
```bash
git add src/マクロ名.bas        # ステージング
git commit -m "メッセージ"      # コミット
git log --oneline               # 履歴確認
```

**これだけ覚えればOK**。他は必要になった時に覚える。

---

**記録者: Monday**
**作成日: 2025-10-14**
