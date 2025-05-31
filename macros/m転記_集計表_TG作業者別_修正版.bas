Attribute VB_Name = "m転記_集計表_TG作業者別_修正版"

' TG作業者別集計転記マクロ（空白値→0に変換、修正版）
Sub 転記_集計表_TG作業者別()
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim tblSource As ListObject
    Dim tblOutput As ListObject
    Dim dict As Object 'Scripting.Dictionary
    Dim outputArray() As Variant
    Dim dataArray As Variant
    
    On Error GoTo ErrorHandler
    
    ' 高速化設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 基本設定
    Set wb = ThisWorkbook
    Dim sourceSheetName As String
    Dim sourceTableName As String
    Dim outputSheetName As String
    Dim outputTableName As String
    Dim outputStartCellAddress As String
    Dim outputHeader As Range
    
    sourceSheetName = "全工程"
    sourceTableName = "全工程テーブル"
    outputSheetName = "集計表_TG作業者別"
    outputTableName = "集計表_TG作業者別テーブル"
    outputStartCellAddress = "A1"
    
    ' データソース確認
    Set wsSource = wb.Sheets(sourceSheetName)
    Set tblSource = wsSource.ListObjects(sourceTableName)
    
    ' 列インデックス取得
    Dim colDate As Long, colProcess As Long, colWorker As Long
    Dim colJisseki As Long, colDandori As Long, colKadou As Long, colFuryo As Long
    
    colDate = GetColumnIndex(tblSource, "日付")
    colProcess = GetColumnIndex(tblSource, "工程")
    colWorker = GetColumnIndex(tblSource, "作業者")
    colJisseki = GetColumnIndex(tblSource, "実績時間")
    colDandori = GetColumnIndex(tblSource, "段取時間")
    colKadou = GetColumnIndex(tblSource, "稼働時間")
    colFuryo = GetColumnIndex(tblSource, "不良数")
    
    ' データ読み込み
    dataArray = tblSource.DataBodyRange.Value
    
    ' Dictionary初期化
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 集計処理
    Dim i As Long, j As Long
    For i = 1 To UBound(dataArray, 1)
        ' TG工程のみ処理
        If dataArray(i, colProcess) = "TG" Then
            Dim currentDate As Date
            Dim currentWorker As String
            Dim dictKey As String
            
            currentDate = dataArray(i, colDate)
            currentWorker = dataArray(i, colWorker)
            dictKey = Format(currentDate, "yyyy/mm/dd") & "_" & currentWorker
            
            ' 空白値→0に変換
            Dim jissekiVal As Double, dandoriVal As Double, kadouVal As Double, furyoVal As Double
            jissekiVal = IIf(IsEmpty(dataArray(i, colJisseki)) Or dataArray(i, colJisseki) = "", 0, CDbl(dataArray(i, colJisseki)))
            dandoriVal = IIf(IsEmpty(dataArray(i, colDandori)) Or dataArray(i, colDandori) = "", 0, CDbl(dataArray(i, colDandori)))
            kadouVal = IIf(IsEmpty(dataArray(i, colKadou)) Or dataArray(i, colKadou) = "", 0, CDbl(dataArray(i, colKadou)))
            furyoVal = IIf(IsEmpty(dataArray(i, colFuryo)) Or dataArray(i, colFuryo) = "", 0, CDbl(dataArray(i, colFuryo)))
            
            If dict.Exists(dictKey) Then
                Dim item As Variant
                item = dict(dictKey)
                item(2) = item(2) + jissekiVal
                item(3) = item(3) + dandoriVal
                item(4) = item(4) + kadouVal
                item(5) = item(5) + furyoVal
                dict(dictKey) = item
            Else
                dict(dictKey) = Array(currentDate, currentWorker, jissekiVal, dandoriVal, kadouVal, furyoVal)
            End If
        End If
    Next i
    
    ' 出力配列作成
    ReDim outputArray(0 To dict.Count, 0 To 5)
    
    ' ヘッダー行
    outputArray(0, 0) = "日付"
    outputArray(0, 1) = "作業者"
    outputArray(0, 2) = "実績時間"
    outputArray(0, 3) = "段取時間"
    outputArray(0, 4) = "稼働時間"
    outputArray(0, 5) = "不良数"
    
    ' データ行
    Dim k As Long
    k = 1
    Dim dictKey As Variant
    For Each dictKey In dict.Keys
        Dim item As Variant
        item = dict(dictKey)
        outputArray(k, 0) = item(0)
        outputArray(k, 1) = item(1)
        outputArray(k, 2) = item(2)
        outputArray(k, 3) = item(3)
        outputArray(k, 4) = item(4)
        outputArray(k, 5) = item(5)
        k = k + 1
    Next
    
    ' 出力先準備
    On Error Resume Next
    Set wsOutput = wb.Sheets(outputSheetName)
    If wsOutput Is Nothing Then
        Set wsOutput = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsOutput.Name = outputSheetName
    End If
    On Error GoTo ErrorHandler
    
    ' 既存のテーブルを削除
    On Error Resume Next
    Set tblOutput = wsOutput.ListObjects(outputTableName)
    If Not tblOutput Is Nothing Then
        tblOutput.Delete
    End If
    On Error GoTo ErrorHandler
    
    ' シート内容をクリア
    wsOutput.Cells.Clear
    
    ' データ出力
    wsOutput.Range(outputStartCellAddress).Resize(UBound(outputArray, 1) + 1, UBound(outputArray, 2) + 1).Value = outputArray
    
    ' テーブル作成
    Set outputHeader = wsOutput.Range(outputStartCellAddress).Resize(1, UBound(outputArray, 2) + 1)
    Set tblOutput = wsOutput.ListObjects.Add(xlSrcRange, _
        wsOutput.Range(outputStartCellAddress).Resize(UBound(outputArray, 1) + 1, UBound(outputArray, 2) + 1), , xlYes)
    tblOutput.Name = outputTableName
    
    MsgBox "転記が完了しました。" & vbCrLf & "出力件数: " & dict.Count & "件", vbInformation
    
Cleanup:
    ' 設定を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    Set wb = Nothing
    Set wsSource = Nothing
    Set wsOutput = Nothing
    Set tblSource = Nothing
    Set tblOutput = Nothing
    Set dict = Nothing
    Set outputHeader = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    GoTo Cleanup
End Sub

' 列インデックス取得ヘルパー関数
Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    Dim i As Long
    GetColumnIndex = 0
    For i = 1 To tbl.ListColumns.Count
        If tbl.ListColumns(i).Name = columnName Then
            GetColumnIndex = i
            Exit For
        End If
    Next i
End Function