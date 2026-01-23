# プロジェクト進捗状況

## 現在の状態
- **最終更新**: 2026-01-23 16:05
- **アクティブタスク**: ryuushutsu_tool（Python流出不良集計ツール）- 保留中

## 進行中
- ryuushutsu_tool: Python流出不良集計ツール（**保留中・不具合あり**）

## 保留中: ryuushutsu_tool

### 概要
VBAマクロのピボットテーブルフィルタリングが遅すぎる（3分以上）ため、PythonでAccess DBに直接アクセスして集計・Excel出力するツールを開発中。

### 完了した部分
- [x] プロジェクト構造 (`src/ryuushutsu_tool/`)
- [x] Access DB接続モジュール (`database.py`)
- [x] 派生カラム計算 (`config.py` - 発生/発見2/アル・ノア/Fr・Rr/LH・RH/モード2)
- [x] 集計ロジック (`aggregator.py`)
- [x] GUI (`gui.py` - tkinter, 日付範囲, 発生, 発見2複数選択, モード2)
- [x] テンプレート方式Excel出力 (`excel_writer.py`)
- [x] 縦軸統一ロジック移植 (`calc_nice_max_value`, `calc_nice_tick_interval`)

### 未解決の問題
1. **win32com SaveエラーでTypeError**
   - `wb.Save()` で `'bool' object is not callable` エラー
   - xlwings → win32com直接に変更したが同じエラー
   - 原因不明、COMオブジェクトの状態が異常？

2. **モード2フィルタ時のエラー**
   - モード2を空白にすると動作
   - 「キズ」等を入れるとエラー
   - pivot_dfの型問題は修正済みだが、保存エラーと複合している可能性

### テンプレート情報
- **ファイル**: `inbox/閲覧_FrRrゾーン特化.xlsm`
- **シート名**: `ゾーンFrRr流出`
- **テーブル名**: `_アルヴェルFr`, `_アルヴェルRr`, `_ノアヴォクFr`, `_ノアヴォクRr`
- **グラフ名**: `グラフ1`, `グラフ2`, `グラフ3`, `グラフ4`（スペースなし）
- **テーブル列**: 品番, ゾーン, 数量

### 次回の作業
1. win32com Saveエラーの調査・修正
2. モード2フィルタの動作確認
3. Windows環境でのテスト
4. PyInstallerでexe化

### 起動方法（PowerShell）
```powershell
cd \\wsl.localhost\Ubuntu-22.04\home\shostako\ClaudeCode\excel-auto\src
python -m ryuushutsu_tool.run
```

---

## 完了済み（今回セッション・夜2）
- [x] 成形P統合マクロ・シート作成
  - m転記_統合_成形P.bas新規作成
  - G→T→Hテーブルをコピー（タイトル行含む）
  - 期間ごとに改ページ、印刷範囲設定
- [x] 加工P統合マクロ・シート作成
  - m転記_統合_加工P.bas新規作成
- [x] 塗装P統合マクロ・シート作成
  - m転記_統合_塗装P.bas新規作成
- [x] 各一括マクロにP追加
  - m転記_一括_成形.bas, m転記_一括_加工.bas, m転記_一括_塗装.bas
- [x] 印刷レイアウト修正
  - CenterVertically = False, CenterHorizontally = False追加
  - 3つのPマクロすべてに適用

## 完了済み（今回セッション・夜）
- [x] m転記_見せる表_加工.bas - 加工NWからショット数取得追加
  - 期間検証追加（加工G vs 加工NW）
  - 7行目ショット数の転記元を流出G→加工NWに変更
- [x] m転記_一括_加工.bas - 加工NWシート追加
  - 対象シート・テーブル・マクロ配列に加工NW追加
- [x] mクエリ参照元変更_複数月対応.bas - 日報加工クエリ追加
  - 対象シート配列に「加工NW」追加
  - `Create日報加工複数月結合Query`関数追加

## 完了済み（今回セッション・夕方）
- [x] m転記_日報_加工NW.bas - 新規作成（塗装NWベース）
  - パターンA: 62-xxxxx形式 → FrLH/RH両方（×1）
  - パターンB: ロット=「単」かつ特定数字含む → 特定1グループ（×2）
  - 補給品: 上記以外（ロット=「単」→×2、それ以外→末尾LH/RH判定）

## 完了済み（今回セッション・午後2）
- [x] m転記_廃棄_成形H/塗装H/加工H.bas - ショット数処理全削除、ロット数量テーブル不参照
- [x] m転記_流出_成形G/塗装G/加工G.bas - ショット数処理全削除、ロット数量テーブル不参照
- [x] m転記_流出_成形G.bas - BubbleSortDesc → QuickSortDescに統一

## 完了済み（今回セッション・午後）
- [x] 手直しクエリ.pq - 年単位ファイル単独参照、品番末尾・注番月列除外
- [x] mクエリ参照元変更_複数月対応.bas - ロット数量/番号固定クエリ削除、手直しシート除外
- [x] m転記_手直し_成形T/塗装T/加工T.bas - ショット数処理全削除、ロット数量テーブル不参照

## 完了済み（今回セッション・午前）
- [x] UserForm5 ID範囲指定対応（22-25 → 22,23,24,25）
- [x] UserForm5 DataModifiedフラグ追加（クエリ更新判定用）
- [x] UserForm5 削除モード対応（データ無効化、実行後フォーム閉じる）
- [x] mゾーン別データ転送ADO.bas - 品番末尾・注番月フィールド削除
- [x] mクエリ西暦更新.bas - 新規作成（Power Query接続先を動的変更）
- [x] 不良集計ゾーン別ADO.pq - 年間DBのみ参照に変更
- [x] m自動採番リセットAccess.bas - DBパス動的構築対応

## 完了済み（過去）
- [x] VBAマクロ開発基盤（src/macros分離、bas2sjis変換スクリプト）
- [x] 各種転記マクロ、フィルターマクロ、クエリ関連
- [x] UserForm2〜5
- [x] /wrapコマンド、/vbaコマンド

## 未完了・保留
- ryuushutsu_tool（上記参照）

## 次セッションへの引き継ぎ
- **ryuushutsu_tool保留中**: win32com Saveエラー未解決
- **P統合シート完成**: 成形P/加工P/塗装P
- **加工NW関連マクロ整備完了**: すべてに加工NW統合済み
- **参照すべきリソース**:
  - `src/ryuushutsu_tool/` - Python流出集計ツール
  - `inbox/閲覧_FrRrゾーン特化.xlsm` - テンプレートExcel
  - `inbox/mグラフ軸設定.bas` - 縦軸統一ロジック参考

## 直近のGitコミット
- 85a95ef feat: 番号転送ADOマクロ改造（全とっかえ方式、DBパス動的構築）
- d7d1854 feat: P統合マクロ追加（成形P/加工P/塗装P）、一括マクロ改修
