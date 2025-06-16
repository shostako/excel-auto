# excel-auto プロジェクト設定

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
- **自動変換**: 有効（生成・修正後に自動的にShift-JIS変換を実行）

#### Power Query (M言語) 生成
- **形式**: 直接ファイル作成（.pq または .txt）
- **保存先**: `src/`ディレクトリ
- **コメント**: 日本語

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
   - **自動変換が実行される**（Shift-JISでWindows側に保存）

3. **手動変換**（必要な場合のみ）
   ```bash
   bas2sjis src/マクロ名.bas
   ```

4. **テスト方法**
   - ExcelのVBEでインポート
   - 実データでテスト実行

## 重要な制約事項

### 参考マクロの読み込み（超重要！）
**必須事項**: 参考マクロフォルダのファイルは**必ず文字コード変換してから読むこと**

```bash
# 読む前に必ず実行
iconv -f SHIFT-JIS -t UTF-8 "参考マクロ/ファイル名.bas" | head -100
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
iconv -f SHIFT-JIS -t UTF-8 "参考マクロ/ファイル名.bas" | head -100

# 文字化け確認用
file "参考マクロ/ファイル名.bas"
```

#### エンコーディング確認
- **参考マクロ**: Shift-JIS（Excelエクスポート）
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