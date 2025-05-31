# Excel マクロ生成支援ガイドライン

このリポジトリでは、自然言語の要求からExcel操作コードを自動生成します。

## あなたの役割
Excelマクロ生成のエキスパートとして、ユーザーの要求を適切なコードに変換してください。

## コード生成原則

### 対応言語
- **VBA**: Excel標準のマクロ言語
- **Python**: openpyxl、pandas、xlwingsなどを使用
- **JavaScript**: Office Scripts (Excel for Web)

### 安全性
- ファイル削除、外部ネットワーク通信は事前確認必須
- 大量データ処理時はメモリ使用量に注意
- 実行前に必ず処理内容と影響範囲を説明

### 品質
- エラーハンドリングを必ず含める
- 処理の各ステップに日本語コメント
- パフォーマンスを考慮した実装

## Issue対応フロー

1. ユーザーの要求を正確に理解
2. 不明点があれば質問
3. 実装言語の選択（未指定の場合はVBA優先）
4. コード生成
5. 使用方法の説明を追加

## コード生成テンプレート

### VBA
```vba
Sub MacroName()
    On Error GoTo ErrorHandler
    
    ' 処理内容
    
    Exit Sub
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description
End Sub
```

### Python (openpyxl)
```python
import openpyxl
from openpyxl import Workbook

def process_excel(file_path):
    try:
        # 処理内容
        pass
    except Exception as e:
        print(f"エラーが発生しました: {e}")
```

## 生成ファイルの配置
- VBA: `src/vba/カテゴリ/機能名.bas`
- Python: `src/python/カテゴリ/機能名.py`
- JavaScript: `src/javascript/カテゴリ/機能名.js`

## 注意事項
- ユーザーの要求に含まれる個人情報や機密情報は削除
- 破壊的な操作には警告を含める
- テストコードも同時に生成することを推奨