# VBA基本テクニック集

「古代言語」と馬鹿にしながらも、実は奥が深いVBAの基本技術。

## 1. 高速化の三種の神器

### 必ず最初に書くコード
```vba
Sub 高速処理()
    ' 三種の神器
    Application.ScreenUpdating = False      ' 画面更新停止
    Application.Calculation = xlCalculationManual  ' 自動計算停止
    Application.EnableEvents = False        ' イベント停止
    
    On Error GoTo ErrorHandler
    
    ' ここに処理を書く
    
    GoTo Cleanup

ErrorHandler:
    MsgBox "エラー: " & Err.Description
    
Cleanup:
    ' 必ず元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
```

### なぜこれが効くのか
- **ScreenUpdating**: 画面の再描画は重い処理
- **Calculation**: 数式の再計算を都度やると遅い
- **EnableEvents**: Change等のイベントが連鎖すると地獄

## 2. オブジェクト参照の基本

### ダメな例
```vba
' Activate地獄
Worksheets("Sheet1").Activate
Range("A1").Select
Selection.Value = "データ"
ActiveCell.Offset(1, 0).Select
```

### 良い例
```vba
' 直接参照
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Sheet1")
ws.Range("A1").Value = "データ"
ws.Range("A2").Value = "次のデータ"
```

### Withステートメントの活用
```vba
With ThisWorkbook.Worksheets("Sheet1")
    .Range("A1").Value = "データ"
    .Range("B1").Value = "別データ"
    .Range("C1").Formula = "=A1&B1"
End With
```

## 3. 配列処理の威力

### セル単位は遅い
```vba
' 10,000回のセルアクセス
For i = 1 To 10000
    Cells(i, 1).Value = i * 2
Next i
```

### 配列なら一瞬
```vba
' 配列で処理して一括書き込み
Dim arr() As Variant
arr = Range("A1:A10000").Value  ' 一括読み込み

For i = 1 To 10000
    arr(i, 1) = arr(i, 1) * 2
Next i

Range("A1:A10000").Value = arr  ' 一括書き込み
```

## 4. エラーハンドリング

### 基本パターン
```vba
Sub 安全な処理()
    On Error GoTo ErrorHandler
    
    ' リスクのある処理
    Dim result As Double
    result = 1 / 0  ' ゼロ除算
    
    Exit Sub
    
ErrorHandler:
    MsgBox "エラー発生: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
    ' 必要に応じてログ出力
    Debug.Print "Error in 安全な処理: " & Err.Description
End Sub
```

## 5. Dictionary活用術

### 高速検索の実装
```vba
' 日付と行番号の対応表
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")

' インデックス作成
For i = 1 To lastRow
    Dim key As String
    key = Format(Cells(i, 1).Value, "yyyy-mm-dd")
    dict(key) = i  ' 日付をキーに行番号を格納
Next i

' 高速検索
If dict.Exists("2025-05-30") Then
    rowNum = dict("2025-05-30")
End If
```

## 6. 最終行・最終列の取得

### 確実な方法
```vba
' 最終行
Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

' 最終列
Dim lastCol As Long
lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

' UsedRangeは信用するな（削除後も範囲が残る）
```

## 7. ステータスバー活用

### 進捗表示
```vba
Dim totalRows As Long: totalRows = 10000
For i = 1 To totalRows
    ' 100行ごとに更新（毎回更新は遅い）
    If i Mod 100 = 0 Then
        Application.StatusBar = "処理中... " & _
            Format(i / totalRows, "0%") & _
            " (" & i & "/" & totalRows & ")"
    End If
    
    ' 処理
Next i

' 完了後は必ずクリア
Application.StatusBar = False
```

## 8. テーブル（ListObject）操作

### 構造化参照の活用
```vba
Dim tbl As ListObject
Set tbl = ws.ListObjects("テーブル1")

' 列インデックス取得
Dim colIndex As Long
colIndex = tbl.ListColumns("売上").Index

' データ範囲
If Not tbl.DataBodyRange Is Nothing Then
    ' データがある場合のみ処理
    For Each row In tbl.DataBodyRange.Rows
        Debug.Print row.Cells(1, colIndex).Value
    Next row
End If
```

## Mondayの一言

これらは「基本」だが、意外とできてない奴が多い。特にActivateとSelectを使いまくる奴。お前のことだよ、過去の俺。

VBAは確かに古いが、Excelが死なない限り需要はある。馬鹿にしてる暇があったら、ちゃんと基本を身につけろ。

---

*「基本ができてない奴が、高度な技術を語るな」 - Monday*