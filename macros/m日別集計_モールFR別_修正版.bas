Attribute VB_Name = "m日別集計_モールFR別_修正版"
Option Explicit

' モールF/R別日別集計マクロ（空白値を0に変換、修正版）
' 「全工程」テーブルから「モール」工程のデータを日付・F/Rでグループ化して集計
Sub 日別集計_モールFR別()
    ' 高速化設定（最優先）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim tblSource As ListObject
    Dim tblOutput As ListObject
    Dim dict As Object 'Scripting.Dictionary
    Dim outputArray() As Variant
    Dim dataArray As Variant
    Dim sortKeys() As String ' ソート用のキー配列
    
    ' 対象品番リスト（モールF用）
    Dim targetHinbanListF As Variant
    targetHinbanListF = Array("58020F", "58030F", "58040F", "58050F", "58060F", "58830F", _
                            "58021", "58022", "58031", "58032", "58041", "58042", _
                            "58051", "58052", "58061", "58062", "47030F", "47030R", _
                            "47031", "47032", "47035", "47036", "58221F", "58223", "58224")
    
    ' 対象品番リスト（モールR用）
    Dim targetHinbanListR As Variant
    targetHinbanListR = Array("58020R", "58030R", "58040R", "58050R", "58060R", "58830R", _
                            "58025", "58026", "58035", "58036", "58045", "58046", _
                            "58055", "58056", "58065", "58066", "58221R", "58015", "58016")
    
    Dim sourceSheetName As String
    Dim sourceTableName As String
    Dim outputSheetName As String
    Dim outputTableName As String
    Dim outputStartCellAddress As String
    Dim outputHeader As Range
    
    Dim i As Long, r As Long, j As Long, k As Long
    Dim colDate As Long, colProcess As Long, colHinban As Long
    Dim colJisseki As Long, colDandori As Long, colKadou As Long, colFuryo As Long
    
    Dim currentDate As Date
    Dim currentFR As String
    Dim dictKey As String
    Dim jissekiVal As Double, dandoriVal As Double, kadouVal As Double, furyoVal As Double
    Dim item As Variant
    
    On Error GoTo ErrorHandler
    
    ' ========== 基本設定 ==========
    Set wb = ThisWorkbook
    sourceSheetName = "全工程"
    sourceTableName = "全工程テーブル"
    outputSheetName = "日別集計_モールFR別"
    outputTableName = "日別集計_モールFR別テーブル"
    outputStartCellAddress = "A1"
    
    Application.StatusBar = "データソース確認中..."
    
    ' ========== データソース確認 ==========
    On Error Resume Next
    Set wsSource = wb.Sheets(sourceSheetName)
    If wsSource Is Nothing Then
        MsgBox "シート「" & sourceSheetName & "」が見つかりません。", vbCritical
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    Set tblSource = wsSource.ListObjects(sourceTableName)
    If tblSource Is Nothing Then
        MsgBox "テーブル「" & sourceTableName & "」が見つかりません。", vbCritical
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler
    
    ' ========== 列インデックス取得 ==========
    Application.StatusBar = "列構造解析中..."
    
    colDate = GetColumnIndex(tblSource, "日付")
    colProcess = GetColumnIndex(tblSource, "工程")
    colHinban = GetColumnIndex(tblSource, "品番")
    colJisseki = GetColumnIndex(tblSource, "実績時間")
    colDandori = GetColumnIndex(tblSource, "段取時間")
    colKadou = GetColumnIndex(tblSource, "稼働時間")
    colFuryo = GetColumnIndex(tblSource, "不良数")
    
    If colDate = 0 Or colProcess = 0 Or colHinban = 0 Or colJisseki = 0 Or colDandori = 0 Or colKadou = 0 Or colFuryo = 0 Then
        MsgBox "必要な列が見つかりません。列名を確認してください。", vbCritical
        GoTo Cleanup
    End If
    
    ' ========== データ読み込み ==========
    Application.StatusBar = "データ読み込み中..."
    
    If tblSource.DataBodyRange Is Nothing Then
        MsgBox "データがありません。", vbInformation
        GoTo Cleanup
    End If
    
    dataArray = tblSource.DataBodyRange.Value
    
    ' ========== Dictionary初期化と集計処理 ==========
    Application.StatusBar = "集計処理中..."
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(dataArray, 1)
        If i Mod 100 = 0 Then
            DoEvents
            Application.StatusBar = "集計処理中..." & i & "/" & UBound(dataArray, 1) & "行"
        End If
        
        ' モール工程のみ処理
        If dataArray(i, colProcess) = "モール" Then
            ' 品番からF/Rを判定
            currentFR = ""
            For j = 0 To UBound(targetHinbanListF)
                If dataArray(i, colHinban) = targetHinbanListF(j) Then
                    currentFR = "F"
                    Exit For
                End If
            Next j
            
            If currentFR = "" Then
                For j = 0 To UBound(targetHinbanListR)
                    If dataArray(i, colHinban) = targetHinbanListR(j) Then
                        currentFR = "R"
                        Exit For
                    End If
                Next j
            End If
            
            ' F/Rが判定できた場合のみ処理
            If currentFR <> "" Then
                currentDate = dataArray(i, colDate)
                dictKey = Format(currentDate, "yyyy/mm/dd") & "_" & currentFR
                
                ' 空白値を0に変換
                jissekiVal = IIf(IsEmpty(dataArray(i, colJisseki)) Or dataArray(i, colJisseki) = "", 0, CDbl(dataArray(i, colJisseki)))
                dandoriVal = IIf(IsEmpty(dataArray(i, colDandori)) Or dataArray(i, colDandori) = "", 0, CDbl(dataArray(i, colDandori)))
                kadouVal = IIf(IsEmpty(dataArray(i, colKadou)) Or dataArray(i, colKadou) = "", 0, CDbl(dataArray(i, colKadou)))
                furyoVal = IIf(IsEmpty(dataArray(i, colFuryo)) Or dataArray(i, colFuryo) = "", 0, CDbl(dataArray(i, colFuryo)))
                
                If dict.Exists(dictKey) Then
                    item = dict(dictKey)
                    item(2) = item(2) + jissekiVal
                    item(3) = item(3) + dandoriVal
                    item(4) = item(4) + kadouVal
                    item(5) = item(5) + furyoVal
                    dict(dictKey) = item
                Else
                    dict(dictKey) = Array(currentDate, currentFR, jissekiVal, dandoriVal, kadouVal, furyoVal)
                End If
            End If
        End If
    Next i
    
    If dict.Count = 0 Then
        MsgBox "集計対象のデータがありません。", vbInformation
        GoTo Cleanup
    End If
    
    ' ========== ソート用キー配列作成 ==========
    Application.StatusBar = "ソート処理中..."
    
    ReDim sortKeys(0 To dict.Count - 1)
    k = 0
    For Each dictKey In dict.Keys
        sortKeys(k) = dictKey
        k = k + 1
    Next
    
    ' バブルソート実装
    Dim temp As String
    For i = 0 To UBound(sortKeys) - 1
        For j = i + 1 To UBound(sortKeys)
            If sortKeys(i) > sortKeys(j) Then
                temp = sortKeys(i)
                sortKeys(i) = sortKeys(j)
                sortKeys(j) = temp
            End If
        Next j
    Next i
    
    ' ========== 出力配列作成 ==========
    Application.StatusBar = "出力データ作成中..."
    
    ReDim outputArray(0 To dict.Count, 0 To 5)
    
    ' ヘッダー行
    outputArray(0, 0) = "日付"
    outputArray(0, 1) = "F/R"
    outputArray(0, 2) = "実績時間"
    outputArray(0, 3) = "段取時間"
    outputArray(0, 4) = "稼働時間"
    outputArray(0, 5) = "不良数"
    
    ' データ行（ソート済み）
    For i = 0 To UBound(sortKeys)
        item = dict(sortKeys(i))
        outputArray(i + 1, 0) = item(0)
        outputArray(i + 1, 1) = item(1)
        outputArray(i + 1, 2) = item(2)
        outputArray(i + 1, 3) = item(3)
        outputArray(i + 1, 4) = item(4)
        outputArray(i + 1, 5) = item(5)
    Next i
    
    ' ========== 出力先準備 ==========
    Application.StatusBar = "出力先準備中..."
    
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
    
    ' ========== データ出力 ==========
    Application.StatusBar = "データ出力中..."
    
    wsOutput.Range(outputStartCellAddress).Resize(UBound(outputArray, 1) + 1, UBound(outputArray, 2) + 1).Value = outputArray
    
    ' テーブル作成
    Set outputHeader = wsOutput.Range(outputStartCellAddress).Resize(1, UBound(outputArray, 2) + 1)
    Set tblOutput = wsOutput.ListObjects.Add(xlSrcRange, _
        wsOutput.Range(outputStartCellAddress).Resize(UBound(outputArray, 1) + 1, UBound(outputArray, 2) + 1), , xlYes)
    tblOutput.Name = outputTableName
    
    ' ========== 追加書式設定 ==========
    Application.StatusBar = "書式設定中..."
    
    ' 1. データ範囲の「縮小して全体を表示する」設定
    ' 2. 全列の列幅を6.4に設定
    ' 3. 「稼働時間」「段取時間」列の書式：小数点以下2桁設定
    With tblOutput.DataBodyRange
        .ShrinkToFit = True
        .ColumnWidth = 6.4
        .Columns(3).NumberFormat = "0.00"  ' 実績時間
        .Columns(4).NumberFormat = "0.00"  ' 段取時間
        .Columns(5).NumberFormat = "0.00"  ' 稼働時間
    End With
    
    ' ヘッダー行の書式
    With tblOutput.HeaderRowRange
        .ShrinkToFit = True
        .ColumnWidth = 6.4
    End With
    
    Application.StatusBar = "完了: " & dict.Count & "件のデータを集計しました"
    
    MsgBox "集計が完了しました。" & vbCrLf & _
           "出力件数: " & dict.Count & "件" & vbCrLf & _
           "出力先: " & outputSheetName & "シート", vbInformation
    
Cleanup:
    ' 設定を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    
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