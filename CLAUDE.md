# excel-auto プロジェクト設定

---

## セッション開始時の必須アクション
**重要：新しいセッションでこのプロジェクトに入った際は、以下を最初に実行すること**

1. **PROGRESS.md を読む**: `PROGRESS.md` を開き、現在の状態と未完了タスクを把握
2. **GitHub Issues 確認**: `gh issue list` でオープンなIssueを確認（GitHub連携していない場合はスキップ）
3. **次のアクション確認**: 「次セッションへの引き継ぎ」セクションを確認し、継続すべきタスクを認識
4. **ユーザーへの報告**: 認識した未完了タスク・Issueをユーザーに報告し、優先度を確認

**目的**: セッション間の引き継ぎを確実にし、前回の作業を途切れなく継続するため

**GitHub連携の判定**: `.git/config`に`[remote "origin"]`があればGitHub連携済み

## PROGRESS.md手動更新時の注意
- **時刻更新時は必ず`date '+%Y-%m-%d %H:%M'`で現在時刻を確認してから記入**
- 推測や概算で時刻を入れない

---

# 共通設定（Monday設定）

## 作業ログの習慣
**重要：セッション終了前に必ず作業ログを記録すること**

### 必須チェック項目（記録前に必ず実行）
1. **環境情報確認**: `<env>`の`Today's date`で現在日付を確認
2. **ファイル名確認**: `logs/yyyy-MM.md`（年月は現在日付ベース）
3. **ヘッダー確認**: `## yyyy-MM-dd 作業概要`（日付は現在日付ベース）

### 記録ルール
- **基本**：各プロジェクトの`logs/yyyy-MM.md`にプロジェクト固有の作業を記録
- **例外**：プロジェクト横断的な内容のみ`/home/shostako/ClaudeCode/logs/yyyy-MM.md`
- その日の作業内容、技術的発見、教訓を記録
- 今後の参照のために詳細に記述する
- **日付記録基準**: ユーザー現地時間（東京時間）ベースで統一

## Claude Code公式リファレンス

### サブエージェント機能活用方針

**検証済み**: 2025-11-07にGemini提供情報をファクトチェック完了（`docs/claude-code-references/SUBAGENT_VERIFICATION_REPORT.md`）

#### 基本方針
- **大規模調査**: Exploreサブエージェントに委任（10.8倍の効率化）
- **独立タスク**: 並列実行で処理（1.8倍の効率化）
- **簡単な調査**: 直接実行（速度優先）

#### 使用推奨場面
- 参考マクロの分析: Explore (medium)
- 複数マクロの最適化確認: Explore (quick)
- 新機能の実装計画: Plan (sonnet)
- ドキュメント全体の見直し: 4x Explore並列

#### 詳細ガイド
- **ベストプラクティス**: `docs/claude-code-references/SUBAGENT_BEST_PRACTICES.md`
- **検証レポート**: `docs/claude-code-references/SUBAGENT_VERIFICATION_REPORT.md`
- **クイックリファレンス**: `docs/claude-code-references/README.md`

### Plan Mode運用ルール

**重要**: Plan Modeは計画確認のために有用だが、調査フェーズでの効率化が必要。

#### Plan Modeの価値（維持すべき）
- 計画を立ててユーザー確認を取る
- 一気に実装せず、段階的に進める
- 予期しない変更を防ぐ

#### Plan Mode内での調査方法
**読み取り専用操作は直接実行**（サブエージェント不要）：
- `iconv -f SHIFT-JIS -t UTF-8`（文字コード変換）
- `cat`, `head`, `tail`（ファイル表示）
- `ls`（ファイル一覧）
- `grep`（検索）

**理由**：
- サブエージェントは最新設定（フック等）を知らない
- サブエージェントは実行を避けて推測する傾向
- 直接実行の方が速く正確（10秒 vs 2分以上）

**サブエージェントを使うべき調査**：
- 複数ファイルにまたがる構造分析
- コードベース全体の理解が必要な調査
- 複雑な依存関係の調査

#### VBAマクロ作業でのPlan Mode
1. **Plan modeは維持**（計画確認のため）
2. **調査は直接実行**（iconvでファイル読む）
3. **excel-vba-expertスキルの早期活用**
   - 計画段階からスキルの知識を参照
   - 計画提示後にスキル起動して実装
4. **ユーザー確認後に実装開始**

#### 理想的なフロー（VBAマクロ作業）
```
ユーザー: 「コメント標準化して」
     ↓
Plan mode開始
     ↓
直接iconvでファイル読む（10秒）← サブエージェント不要
     ↓
excel-vba-expertの知識で計画立案
     ↓
ExitPlanModeで計画提示
     ↓
ユーザー: 「OK」or「開始」
     ↓
スキル起動 + 実装
```

## ユーザーコマンド

### トークン確認コマンド
**トリガー**: 「トークン」「トークン見せて」「トークン表示して」等

ユーザーがこれらのキーワードを使った場合、Claude Code Actionの認証情報を表示する：

- **ファイル**: `/home/shostako/.claude/.credentials.json`
- **用途**: GitHub Repository secretsへのコピペ用
- **表示内容**:
  - CLAUDE_ACCESS_TOKEN: [フル値]
  - CLAUDE_REFRESH_TOKEN: [フル値]
  - CLAUDE_EXPIRES_AT: [タイムスタンプ値]

**実行方法**:
```bash
Read /home/shostako/.claude/.credentials.json
```

出力フォーマット:
```
Claude Code Action Tokens:

CLAUDE_ACCESS_TOKEN:
[accessToken値]

CLAUDE_REFRESH_TOKEN:
[refreshToken値]

CLAUDE_EXPIRES_AT:
[expiresAt値]
```

## ナレッジベース参照
技術的質問や過去の教育内容については以下を参照：
- **UNIX/Linux**: `/home/shostako/ClaudeCode/knowledge/unix-systems.md`
- **プログラミング**: `/home/shostako/ClaudeCode/knowledge/programming.md`
- **ツール・コマンド**: `/home/shostako/ClaudeCode/knowledge/tools.md`
- **問題解決**: `/home/shostako/ClaudeCode/knowledge/troubleshooting.md`

未知の質問に対して教育した内容は適切なファイルに記録すること。

---

# プロジェクト固有設定（Excel開発）

## 必読ファイル（IMPORTANT: 作業開始前に必ず読むこと）

**YOU MUST** 以下のファイルを作業開始時に読み込んでから作業を開始すること：

1. `/home/shostako/ClaudeCode/excel-auto/docs/excel-knowledge/failures/001_activate_vs_screenupdating.md`
   - 画面ちらつき問題の根本原因（Activateメソッド）

2. `/home/shostako/ClaudeCode/excel-auto/docs/excel-knowledge/patterns/VBA_OPTIMIZATION_PATTERNS.md`
   - VBA最適化の基本パターン

3. `/home/shostako/ClaudeCode/excel-auto/docs/excel-knowledge/techniques/VBA_BASIC_TECHNIQUES.md`
   - 基本テクニック集

4. `/home/shostako/ClaudeCode/excel-auto/docs/excel-knowledge/claude-code/EXCEL_MACRO_KNOWLEDGE_BASE.md`
   - Claude Code用実戦的マクロ開発ナレッジ（統合版）

これらの内容を踏まえずにコードを書くと、同じ失敗を繰り返すことになる。

## 基本ルール

### 出力形式ルール

#### VBAマクロ生成
- **形式**: 直接ファイル作成（.basファイル）
- **保存先**: `src/`ディレクトリ
- **コメント**: 日本語
- **完了メッセージ**: エラー時以外は表示しない
- **エンコーディング**: UTF-8（変換スクリプトでShift-JIS化）
- **重要**: srcに保存後、**必ずbas2sjisスクリプトを実行**してmacrosに出力すること

#### Power Query (M言語) 生成
- **形式**: 直接ファイル作成（.pq または .txt）
- **保存先**: `src/`ディレクトリ
- **コメント**: 日本語
- **エンコーディング**: UTF-8のまま（変換不要）
- **必須手順**: srcに保存後、**macrosフォルダにコピー**すること
  ```bash
  cp src/クエリ名.pq macros/
  ```

#### 関数・数式生成
- **形式**: 通常メッセージ内に記載
- **説明**: 日本語で使用方法を説明

## コーディング標準

### エラーハンドリング
```vba
Sub 処理名()
    Application.StatusBar = "処理を開始します..."

    On Error GoTo ErrorHandler

    ' メインの処理

    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
End Sub
```

### 進捗表示
- **ステータスバー使用**: 必須（長時間処理時）
- **更新頻度**: 100件ごと、または処理段階ごと
- **クリア方法**: `Application.StatusBar = False`
- **注意**: 専用のクリアマクロは作らない

### モジュール命名規則
- **接頭辞**: `m` を付ける（例: `m日別集計_モールFR別`）
- **命名**: 日本語可、アンダースコア区切り

## マクロ開発フロー

1. **要件確認**
   - 入力元（シート名、テーブル名）
   - 出力先（シート名、テーブル名）
   - 処理内容の明確化

2. **コード生成**
   - `src/`ディレクトリに.basファイルとして保存
   - モジュール名を正しく設定

3. **必須変換**（毎回実行）
   ```bash
   bas2sjis src/マクロ名.bas
   ```
   - macrosフォルダにShift-JIS版が出力される
   - この手順を忘れるとExcelに取り込めない

4. **テスト方法**
   - ExcelのVBEでインポート
   - 実データでテスト実行

## 重要な制約事項

### 参考マクロの読み込み（超重要！）
**必須事項**: inboxフォルダの参考マクロファイルは**必ず文字コード変換してから読むこと**

```bash
# 読む前に必ず実行
iconv -f SHIFT-JIS -t UTF-8 "inbox/ファイル名.bas" | head -100
```

**理由**:
- ユーザーがExcelからエクスポートしたファイルは**Shift-JIS**
- そのまま読むと文字化けして**列名を誤認識**する
- 過去の事故例：文字化けした列名で修正→本番で動作しない

**禁止事項**:
- 文字化けしたまま読み進めること
- 文字化けした内容を基に修正すること
- UTF-8版があってもオリジナルの確認を怠ること

### 画面更新の抑制について
- `Application.ScreenUpdating = False` は**積極的に使用する**（高速化の基本）
- CommandButtonから呼び出される場合の注意点：
  - 個別マクロの最後で`True`に戻さない（CommandButtonに任せる）
  - 重複設定は無害なので気にしない
- **絶対に使わない：`Activate`メソッド**
  - これが画面ちらつきの真犯人
  - データ処理にシート切り替えは不要
  - オブジェクト参照で直接操作すること

### メッセージ表示
- **正常終了時**: メッセージボックス表示なし
- **エラー時のみ**: MsgBoxで詳細表示
- **進捗**: ステータスバーのみ使用

### データ処理の最適化
- 配列処理を優先（Range操作の最小化）
- Dictionaryオブジェクトの活用
- 大量データ時は進捗表示必須

## デバッグ支援

### Debug.Print活用
- 日付変換エラーなどの警告出力
- 本番環境でも残しておく（イミディエイトウィンドウ確認用）

### エラー情報の詳細化
- エラー番号とDescription両方を表示
- 可能な限り発生箇所を特定できる情報を含める

## 思考プロセスの活用（Claude標準thinking機能）

### 概要
- Claude標準の内部思考機能を活用
- ユーザーには見えない思考プロセスで問題を整理
- sequential-thinkingツールは使わない（重いため）

### 使用推奨場面
- **複雑なデータ構造設計**: テーブル間の関係性整理
- **アルゴリズム最適化**: 複数の処理方法の比較
- **エラー原因の特定**: 段階的な問題の切り分け
- **曖昧な仕様の明確化**: 要件の整理と確認事項の洗い出し

### 効果的な活用
- グループ化ロジックの設計前
- 日付・時間処理の妥当性検証
- パフォーマンスボトルネックの特定
- エラーハンドリング戦略の決定

## ローカルGit開発フロー

### 基本方針
- **GitHub Actions使用停止**: Excelマクロ開発には複雑すぎるため
- **直接開発重視**: コード品質向上を最優先
- **文字エンコーディング管理**: UTF-8 ↔ Shift-JIS変換の確実な実行
- **手動クリーンアップ**: セッション終了時にMakefileで実行

### bas2sjisスクリプト

#### 用途と仕組み
- **目的**: UTF-8のbasファイルをShift-JISに変換
- **出力先**: `macros/`ディレクトリ
- **CRLF対応**: 自動的にCRLF→LF変換を実行

#### 使用方法
```bash
# 基本的な使用方法
./scripts/bas2sjis src/マクロ名.bas

# 変換結果の確認
ls -la macros/
```

#### 変換プロセス
1. CRLF→LF変換（一時ファイル使用）
2. UTF-8→Shift-JIS変換（iconv使用）
3. `macros/`ディレクトリに出力

### 文字エンコーディング管理

#### 読み込み時の注意
```bash
# 参考マクロファイル読み込み前に必ず実行
iconv -f SHIFT-JIS -t UTF-8 "inbox/ファイル名.bas" | head -100

# 文字化け確認用
file "inbox/ファイル名.bas"
```

#### エンコーディング確認
- **inbox**: Shift-JIS（Excelエクスポート）
- **src/**: UTF-8（Claude編集用）
- **macros/**: Shift-JIS（Excel取り込み用）

### Git管理手順

#### 推奨コミットフロー
1. **開発**: `src/`でUTF-8ファイル編集
2. **変換**: bas2sjisで`macros/`に出力
3. **確認**: 変換結果とテスト
4. **コミット**: Git add → commit → push

#### ファイル管理
- **追跡対象**: `src/`、`scripts/`、`docs/`
- **除外対象**: `macros/`（.gitignoreで除外推奨）
- **一時ファイル**: 自動削除される

### ワークスペース管理

#### ワークスペースクリーンアップ

**実行方法**:
```bash
make clean  # または自然言語で「クリーンアップして」「クリーン」
```

**動作**:
1. `src/*.bas`、`macros/*.bas`、`inbox/*.bas`を表示
2. 削除確認プロンプト
3. 削除実行
4. Git add → commit（"clean: ワークスペースクリーンアップ - セッション終了"）

**Monday（Claude）の責務**:
- セッション終了時にクリーンアップを提案
- ユーザーの「クリーン」「片付けて」等の指示で`make clean`実行
- Makefileの存在を認識し、自然言語指示を適切に変換

### トラブルシューティング

#### 文字化け問題
- **原因**: エンコーディング不一致
- **対策**: iconvによる事前変換
- **確認**: file コマンドでエンコーディングチェック

#### 変換失敗
- **原因**: CRLF行末、特殊文字
- **対策**: TRANSLIT オプション使用
- **確認**: 変換結果ファイルの内容確認

---

**注意**: 技術的正確性と効率を最優先とする。
