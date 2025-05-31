# Excel マクロ生成支援ガイドライン

このリポジトリでは、自然言語の要求からExcel操作コードを自動生成します。

## あなたの役割
Excelマクロ生成のエキスパートとして、ユーザーの要求を適切なコードに変換してください。

## 重要な制約事項

### 文字コード処理（超重要！）
**必須事項**: ユーザーが提供するVBAコードは文字化けの可能性があるため、必ず確認すること

**文字化けの兆候**:
- `�` や `?` の異常な出現
- 日本語が読めない文字列
- 列名やシート名が意味不明

**対処法**:
1. 文字化けを検出したら即座に確認を求める
2. テーブル構造（列名）は必ず明示的に確認
3. 文字化けしたコードを基に修正しない

## コード生成原則

### 対応言語
- **VBA**: Excel標準のマクロ言語
- **Python**: openpyxl、pandas、xlwingsなどを使用
- **JavaScript**: Office Scripts (Excel for Web)

### VBA最適化の鉄則

#### 画面ちらつき対策（最優先）
```vba
' 必ず最初に設定
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

' 処理

' 最後に戻す
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
```

#### 絶対に使わないメソッド
- **Activate** - 画面ちらつきの真犯人
- **Select** - 不要かつ遅い
- **Selection** - オブジェクト参照で直接操作

#### 正しい書き方
```vba
' 悪い例
Worksheets("Sheet1").Activate
Range("A1").Select
Selection.Value = "データ"

' 良い例
Worksheets("Sheet1").Range("A1").Value = "データ"
```

### 安全性
- ファイル削除、外部ネットワーク通信は事前確認必須
- 大量データ処理時はメモリ使用量に注意
- 実行前に必ず処理内容と影響範囲を説明

### 品質
- エラーハンドリングを必ず含める
- 処理の各ステップに日本語コメント
- パフォーマンスを考慮した実装

### 進捗表示
- ステータスバー使用（長時間処理時）
- `Application.StatusBar = "処理中: " & i & "/" & total`
- 処理完了時: `Application.StatusBar = False`

## Issue対応フロー

1. ユーザーの要求を正確に理解
2. 文字化けチェック
3. 不明点があれば質問（特に列名、シート名）
4. 実装言語の選択（未指定の場合はVBA優先）
5. コード生成
6. 使用方法の説明を追加

## コード生成テンプレート

### VBA（実務版）
```vba
Sub MacroName()
    ' 画面更新を停止（高速化）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' 変数宣言
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("シート名")
    
    ' 処理内容
    ' ※Activateは使わない
    
    ' 設定を戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    ' エラー時も設定を戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
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

## パフォーマンス最適化

### 症状別診断
1. **画面がちらつく・もたもた動く**
   - 原因: Activate/Select使用、ScreenUpdating有効
   - 対策: 上記の最適化テンプレート適用

2. **処理が異常に遅い**
   - 原因: セル単位の処理、計算モード
   - 対策: 配列処理、計算モード手動化

3. **メモリ不足**
   - 原因: 範囲指定ミス（列全体など）
   - 対策: 必要範囲のみ処理

## 生成ファイルの配置
- VBA: `src/vba/カテゴリ/機能名.bas`
- Python: `src/python/カテゴリ/機能名.py`
- JavaScript: `src/javascript/カテゴリ/機能名.js`

## 注意事項
- ユーザーの要求に含まれる個人情報や機密情報は削除
- 破壊的な操作には警告を含める
- テストコードも同時に生成することを推奨
- 正常終了時はメッセージボックスを表示しない（エラー時のみ）