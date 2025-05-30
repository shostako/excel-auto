# ExcelAuto

Excel VBAマクロの開発・管理プロジェクト

## ディレクトリ構成

- `src/` - 開発中のVBAソースコード（UTF-8）
- `ready/` - 変換待ちステージングエリア（自動監視）
- `converted/` - Shift-JIS変換済みファイル（一時保存用）
- `templates/` - よく使うマクロのテンプレート
- `docs/` - ドキュメント・仕様書など
- `test/` - テスト用ファイル

## 使い方

### 方法1: ステージング方式（推奨）
1. `src/`ディレクトリでVBAコードを開発
2. 完成したら`ready/`に移動してステージング
   ```bash
   ./stage.sh マクロ名.bas     # 個別ファイル
   ./stage.sh all              # 全ファイル
   ```
3. 自動変換されてWindows側に保存される

### 方法2: 手動変換
1. `src/`ディレクトリでVBAコードを開発
2. 変換スクリプトで手動変換
   ```bash
   bas2sjis src/マクロ名.bas
   ```
3. ExcelのVBEでインポート

## 自動変換監視の起動

```bash
./watch_ready.sh
```
- `ready/`フォルダを監視
- .basファイルが入ると自動的にShift-JIS変換
- 変換後はWindows側に保存され、元ファイルは`src/`に移動

## 変換先

デフォルトでWindows側の以下のフォルダに出力：
`C:\Users\shost\Documents\Excelマクロ`

## 注意事項

- ソースコードはUTF-8で記述
- VBEインポート用はShift-JISに自動変換される
- モジュール名（Attribute VB_Name）は必ず記載すること