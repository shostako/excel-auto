# Excel マクロ生成支援ツール

自然言語でExcel操作を説明すると、VBA・Python・JavaScriptのコードを自動生成するAI支援ツール

## 🚀 特徴

- **多言語対応**: VBA、Python (openpyxl)、JavaScript (Office Scripts)
- **自動コード生成**: Issueで要求→Claudeが自動実装→PR作成
- **安全なコード**: エラーハンドリング・セキュリティチェック付き
- **豊富なサンプル**: よくあるExcel操作のテンプレート集

## 📝 使い方

### 1. 新しいマクロをリクエスト

[新しいIssue](https://github.com/shostako/excel_macro_trial/issues/new)を作成して、以下のフォーマットで記載:

```markdown
## やりたいこと
（例：売上データを月別に集計してグラフを作成したい）

## 詳細な要件
- 対象となるデータの範囲
- 期待する出力形式
- 使用したい言語（VBA/Python/JavaScript）

@claude 上記の要件でマクロを作成してください
```

### 2. Claudeが自動で対応

1. 要件を分析
2. 適切なコードを生成
3. Pull Requestを作成
4. 使用方法を説明

## 🗂 リポジトリ構成

```
excel_macro_trial/
├── README.md           # このファイル
├── CLAUDE.md          # AI向けガイドライン
├── src/               # 生成されたコード
│   ├── vba/          # VBAマクロ
│   ├── python/       # Pythonスクリプト
│   └── javascript/   # Office Scripts
├── examples/          # サンプル集
└── docs/             # ドキュメント
```

## 🛠 セットアップ

### VBA
1. Excelで「開発」タブを有効化
2. Visual Basic Editorを開く
3. 生成されたコードをコピー＆ペースト

### Python
```bash
pip install openpyxl pandas xlwings
```

### JavaScript (Office Scripts)
Excel for Webで「自動化」タブから使用

## 📚 よくある要求例

- データクリーニング（重複削除、空白行削除）
- レポート自動生成（月次・週次・日次）
- グラフ・チャート作成
- データ統合・分析
- メール送信自動化

## 🤝 コントリビューション

新しいテンプレートやアイデアがあれば、Issueやプルリクエストでお知らせください。

## 📄 ライセンス

MIT License