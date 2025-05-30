Attribute VB_Name = "m転記_流出廃棄_改"
Sub 転記_流出廃棄から集計表()
    Dim ws流出廃棄 As Worksheet
    Dim ws集計表 As Worksheet
    Dim tbl流出廃棄 As ListObject
    Dim targetDate As Date
    Dim foundRow As Long
    Dim i As Long
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' ステータスバーに進捗表示
    Application.StatusBar = "流出廃棄データ転記中..."
    
    ' シートの設定
    Set ws流出廃棄 = ThisWorkbook.Sheets("流出廃棄")
    Set ws集計表 = ThisWorkbook.Sheets("集計表")
    
    ' テーブルの取得
    Set tbl流出廃棄 = ws流出廃棄.ListObjects("_流出廃棄b")
    
    ' 集計表のA1セルから日付を取得
    If Not IsDate(ws集計表.Range("A1").Value) Then
        MsgBox "集計表のA1セルに有効な日付が入力されていません。", vbExclamation
        Application.StatusBar = False
        Exit Sub
    End If
    targetDate = ws集計表.Range("A1").Value
    
    ' 該当する日付の行を検索
    foundRow = 0
    For i = 1 To tbl流出廃棄.DataBodyRange.Rows.Count
        If tbl流出廃棄.DataBodyRange.Cells(i, tbl流出廃棄.ListColumns("日付").Index).Value = targetDate Then
            foundRow = i
            Exit For
        End If
    Next i
    
    ' 日付が見つからなかった場合
    If foundRow = 0 Then
        MsgBox "指定された日付 " & Format(targetDate, "yyyy/mm/dd") & " のデータが見つかりません。", vbExclamation
        Application.StatusBar = False
        Exit Sub
    End If
    
    ' データの転記処理
    Application.StatusBar = "データを転記しています..."
    
    ' 各項目を転記（エラー処理を含め個々に処理）
    With tbl流出廃棄.DataBodyRange.Rows(foundRow)
        ' 成形流出 → J18
        ws集計表.Range("J18").Value = GetColumnValue(tbl流出廃棄, foundRow, "成形流出")
        
        ' 成形流出設計 → P18
        ws集計表.Range("P18").Value = GetColumnValue(tbl流出廃棄, foundRow, "成形流出設計")
        
        ' 成形廃棄設計 → J57
        ws集計表.Range("J57").Value = GetColumnValue(tbl流出廃棄, foundRow, "成形廃棄設計")
        
        ' 塗装流出 → J31
        ws集計表.Range("J31").Value = GetColumnValue(tbl流出廃棄, foundRow, "塗装流出")
        
        ' 塗装流出設計 → P31
        ws集計表.Range("P31").Value = GetColumnValue(tbl流出廃棄, foundRow, "塗装流出設計")
        
        ' 塗装廃棄設計 → L57
        ws集計表.Range("L57").Value = GetColumnValue(tbl流出廃棄, foundRow, "塗装廃棄設計")
        
        ' 加工流出 → F57
        ws集計表.Range("F57").Value = GetColumnValue(tbl流出廃棄, foundRow, "加工流出")
        
        ' 加工流出設計 → H57
        ws集計表.Range("H57").Value = GetColumnValue(tbl流出廃棄, foundRow, "加工流出設計")
        
        ' 加工廃棄設計 → N57
        ws集計表.Range("N57").Value = GetColumnValue(tbl流出廃棄, foundRow, "加工廃棄設計")
    End With
    
    ' 廃棄設計の合計を計算してP57に転記
    Dim 成形廃棄 As Double, 塗装廃棄 As Double, 加工廃棄 As Double
    成形廃棄 = IIf(IsNumeric(ws集計表.Range("J57").Value), ws集計表.Range("J57").Value, 0)
    塗装廃棄 = IIf(IsNumeric(ws集計表.Range("L57").Value), ws集計表.Range("L57").Value, 0)
    加工廃棄 = IIf(IsNumeric(ws集計表.Range("N57").Value), ws集計表.Range("N57").Value, 0)
    
    ws集計表.Range("P57").Value = 成形廃棄 + 塗装廃棄 + 加工廃棄
    
    ' 正常終了
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    Application.StatusBar = False
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical
End Sub

' テーブルから指定列の値を安全に取得する関数
Private Function GetColumnValue(tbl As ListObject, rowIndex As Long, columnName As String) As Variant
    On Error Resume Next
    Dim colIndex As Long
    colIndex = tbl.ListColumns(columnName).Index
    
    If Err.Number <> 0 Then
        ' 列が見つからない場合は0を返す
        GetColumnValue = 0
        Err.Clear
    Else
        ' 値を取得し、空白・エラー値の場合は0を返す
        Dim cellValue As Variant
        cellValue = tbl.DataBodyRange.Cells(rowIndex, colIndex).Value
        
        If IsEmpty(cellValue) Or IsError(cellValue) Then
            GetColumnValue = 0
        ElseIf VarType(cellValue) = vbString And Trim(cellValue) = "" Then
            ' 文字列で空白の場合のみ0を設定
            GetColumnValue = 0
        Else
            GetColumnValue = cellValue
        End If
    End If
    On Error GoTo 0
End Function