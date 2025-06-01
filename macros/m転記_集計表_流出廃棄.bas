Attribute VB_Name = "m転記_集計表_流出廃棄"
Sub 転記_集計表_流出廃棄()
    ' 変数宣言
    Dim ws流出廃棄 As Worksheet
    Dim ws集計表 As Worksheet
    Dim tbl流出廃棄 As ListObject
    Dim targetDate As Date
    Dim foundRow As Long
    Dim i As Long
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' 高速化設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 進捗表示開始
    Application.StatusBar = "流出廃棄データの転記処理を開始します..."
    
    ' シートの設定
    On Error Resume Next
    Set ws集計表 = ThisWorkbook.Sheets("集計表")
    If ws集計表 Is Nothing Then
        MsgBox "「集計表」シートが見つかりません。", vbCritical, "シートエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' 集計表のA1セルから日付を取得
    If Not IsDate(ws集計表.Range("A1").Value) Then
        MsgBox "集計表のA1セルに有効な日付が入力されていません。", vbCritical, "日付エラー"
        GoTo CleanupAndExit
    End If
    targetDate = ws集計表.Range("A1").Value
    
    ' 流出廃棄シートの設定
    On Error Resume Next
    Set ws流出廃棄 = ThisWorkbook.Sheets("流出廃棄")
    If ws流出廃棄 Is Nothing Then
        MsgBox "「流出廃棄」シートが見つかりません。", vbCritical, "シートエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' テーブルの取得
    On Error Resume Next
    Set tbl流出廃棄 = ws流出廃棄.ListObjects("_流出廃棄b")
    If tbl流出廃棄 Is Nothing Then
        MsgBox "「流出廃棄」シートに「_流出廃棄b」テーブルが見つかりません。", vbCritical, "テーブルエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' データ範囲の確認
    If tbl流出廃棄.DataBodyRange Is Nothing Then
        MsgBox "「_流出廃棄b」テーブルにデータがありません。", vbCritical, "データエラー"
        GoTo CleanupAndExit
    End If
    
    ' 該当する日付の行を検索
    Application.StatusBar = "流出廃棄データを検索中..."
    foundRow = 0
    For i = 1 To tbl流出廃棄.DataBodyRange.Rows.Count
        If tbl流出廃棄.DataBodyRange.Cells(i, tbl流出廃棄.ListColumns("日付").Index).Value = targetDate Then
            foundRow = i
            Exit For
        End If
    Next i
    
    ' 日付が見つからなかった場合
    If foundRow = 0 Then
        MsgBox "日付 " & Format(targetDate, "yyyy/mm/dd") & " のデータが見つかりません。", vbCritical, "データエラー"
        GoTo CleanupAndExit
    End If
    
    ' データの転記処理
    Application.StatusBar = "流出廃棄データを転記中..."
    
    ' 各項目を転記（エラー回避のため個別に処理）
    With tbl流出廃棄.DataBodyRange.Rows(foundRow)
        ' 成形流出 → J18
        ws集計表.Range("J18").Value = GetColumnValue(tbl流出廃棄, foundRow, "成形流出")
        
        ' 成形流出累計 → P18
        ws集計表.Range("P18").Value = GetColumnValue(tbl流出廃棄, foundRow, "成形流出累計")
        
        ' 成形廃棄累計 → J31
        ws集計表.Range("J31").Value = GetColumnValue(tbl流出廃棄, foundRow, "成形廃棄累計")
        
        ' 塗装流出 → P31
        ws集計表.Range("P31").Value = GetColumnValue(tbl流出廃棄, foundRow, "塗装流出")
        
        ' 塗装流出累計 → F57
        ws集計表.Range("F57").Value = GetColumnValue(tbl流出廃棄, foundRow, "塗装流出累計")
        
        ' 塗装廃棄累計 → H57
        ws集計表.Range("H57").Value = GetColumnValue(tbl流出廃棄, foundRow, "塗装廃棄累計")
        
        ' 加工流出 → J57
        ws集計表.Range("J57").Value = GetColumnValue(tbl流出廃棄, foundRow, "加工流出")
        
        ' 加工流出累計 → L57
        ws集計表.Range("L57").Value = GetColumnValue(tbl流出廃棄, foundRow, "加工流出累計")
        
        ' 加工廃棄累計 → N57
        ws集計表.Range("N57").Value = GetColumnValue(tbl流出廃棄, foundRow, "加工廃棄累計")
    End With
    
    ' 廃棄累計の合計を計算してP57に転記
    Dim 成形廃棄 As Double, 塗装廃棄 As Double, 加工廃棄 As Double
    成形廃棄 = IIf(IsNumeric(ws集計表.Range("J31").Value), ws集計表.Range("J31").Value, 0)
    塗装廃棄 = IIf(IsNumeric(ws集計表.Range("H57").Value), ws集計表.Range("H57").Value, 0)
    加工廃棄 = IIf(IsNumeric(ws集計表.Range("N57").Value), ws集計表.Range("N57").Value, 0)
    
    ws集計表.Range("P57").Value = 成形廃棄 + 塗装廃棄 + 加工廃棄
    
    ' 正常終了（エラー時以外はメッセージ非表示）
    GoTo CleanupAndExit
    
ErrorHandler:
    MsgBox "転記処理中にエラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "転記エラー"
           
CleanupAndExit:
    ' 後処理
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False  ' ステータスバーをクリア
End Sub

' テーブルから指定列の値を安全に取得する関数
Private Function GetColumnValue(tbl As ListObject, rowIndex As Long, columnName As String) As Variant
    On Error Resume Next
    Dim colIndex As Long
    colIndex = tbl.ListColumns(columnName).Index
    
    If Err.Number <> 0 Then
        ' 列が見つからない場合
        GetColumnValue = 0
        Debug.Print "警告: 列「" & columnName & "」が見つかりません。"
        Err.Clear
    Else
        ' 値を取得（空白の場合は0を返す）
        Dim cellValue As Variant
        cellValue = tbl.DataBodyRange.Cells(rowIndex, colIndex).Value
        If IsEmpty(cellValue) Or cellValue = "" Or IsNull(cellValue) Then
            GetColumnValue = 0
        Else
            GetColumnValue = cellValue
        End If
    End If
    On Error GoTo 0
End Function