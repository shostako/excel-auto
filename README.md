# excel-auto

AI駆動Excel自動化ツール - 自然言語でExcel操作を説明すると、VBA・Python・JavaScriptのコードを自動生成

## 🚀 特徴

- **AI駆動コード生成**: 自然言語での要求から適切なコードを自動生成
- **多言語対応**: VBA、Python (openpyxl)、JavaScript (Office Scripts)
- **安全なコード**: エラーハンドリング・セキュリティチェック付き
- **豊富な知識ベース**: VBA最適化パターン、失敗事例、テクニック集

## 📝 使い方

### 1. 新しいマクロをリクエスト

[新しいIssue](https://github.com/shostako/excel-auto/issues/new)を作成して、以下のフォーマットで記載:

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

1. 要件を分析し、最適な実装方法を提案
2. 高品質なコードを生成（最適化・エラーハンドリング込み）
3. Pull Requestを作成
4. 詳細な使用方法を説明

## 🗂 リポジトリ構成

```
excel-auto/
├── README.md                   # このファイル
├── CLAUDE.md                  # AI向けガイドライン
├── src/                       # 生成されたソースコード
│   ├── vba/                  # VBAマクロ
│   ├── python/               # Pythonスクリプト
│   └── javascript/           # Office Scripts
├── docs/                      # ドキュメント
│   ├── excel-knowledge/      # 知識ベース
│   │   ├── failures/        # 失敗事例・対策
│   │   ├── patterns/        # 最適化パターン
│   │   └── techniques/      # 基本テクニック
│   └── logs/                # 開発ログ
├── templates/                 # テンプレート集
├── inbox/                     # 参考マクロ受け取り場所
└── converted/                # 変換済みファイル
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

- **データ処理**: 重複削除、空白行削除、データクリーニング
- **レポート生成**: 月次・週次・日次の自動レポート作成
- **グラフ・チャート**: データの可視化、動的グラフ作成
- **ピボットテーブル**: 高度な集計・分析処理
- **データ統合**: 複数ファイルのマージ、外部データ取り込み

## 🎯 高品質コードの特徴

### VBA最適化
- **画面ちらつき防止**: `Application.ScreenUpdating = False`
- **高速処理**: Activate/Selectメソッドを使わない最適化
- **メモリ効率**: 適切な変数管理とオブジェクト解放
- **エラーハンドリング**: 堅牢なエラー処理とクリーンアップ

### 安全性
- 破壊的操作の事前警告
- 大量データ処理時のメモリ使用量考慮
- 実行前の処理内容・影響範囲説明

## 🤝 コントリビューション

新しいテンプレートやアイデアがあれば、Issueやプルリクエストでお知らせください。

## 📄 ライセンス

MIT License

---

**Note**: このプロジェクトは、AI（Claude）との協働によるExcel自動化を目指しています。技術的正確性と実用性を最優先に、継続的に改善していきます。