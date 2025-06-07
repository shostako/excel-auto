Attribute VB_Name = "m転記_シートTG品番別"
Option Explicit

' ==========================================
' TG品番別からシートへの転記マクロ
' 「_TG品番別a」テーブルから品番別シートへデータを転記
' ==========================================
Sub 転記_シートTG品番別()
    ' ==========================================
    ' 変数宣言
    ' ==========================================
    Dim wsSource As Worksheet
    Dim sourceTable As ListObject
    Dim sourceData As Range
    Dim i As Long, j As Long, k As Long
    Dim targetDate As Date
    Dim dateColIndex As Long
    Dim foundRow As Long
    Dim targetSheet As Worksheet
    Dim processedCount As Long
    Dim totalDates As Long
    
    ' 品番リスト定義
    Dim partNumbers() As Variant
    Dim sheetNames() As Variant
    partNumbers = Array("53827-60050", "53828-60080")
    sheetNames = Array("53827-60050 RH", "53828-60080 LH")
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' ==========================================
    ' 高速化設定
    ' ==========================================
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 進捗表示開始
    Application.StatusBar = "TG品番別シート転記処理を開始します..."
    
    ' ==========================================
    ' ソースシート・テーブル取得
    ' ==========================================
    ' TG品番別シート取得
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("TG品番別")
    If wsSource Is Nothing Then
        MsgBox "「TG品番別」シートが見つかりません。", vbCritical, "シートエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' ソーステーブル取得
    On Error Resume Next
    Set sourceTable = wsSource.ListObjects("_TG品番別a")
    If sourceTable Is Nothing Then
        MsgBox "「_TG品番別a」テーブルが見つかりません。", vbCritical, "テーブルエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' データ範囲取得
    If sourceTable.DataBodyRange Is Nothing Then
        MsgBox "「_TG品番別a」テーブルにデータがありません。", vbInformation, "データなし"
        GoTo CleanupAndExit
    End If
    Set sourceData = sourceTable.DataBodyRange
    
    ' 日付列のインデックス取得
    On Error Resume Next
    dateColIndex = sourceTable.ListColumns("日付").Index
    If Err.Number <> 0 Then
        MsgBox "「_TG品番別a」テーブルに「日付」列が見つかりません。", vbCritical, "列エラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' ==========================================
    ' メイン処理: 各品番のデータを転記
    ' ==========================================
    totalDates = sourceData.Rows.Count
    processedCount = 0
    
    ' 各品番について処理
    For k = 0 To UBound(partNumbers)
        ' 転記先シートの存在確認
        On Error Resume Next
        Set targetSheet = ThisWorkbook.Worksheets(sheetNames(k))
        On Error GoTo ErrorHandler
        
        If targetSheet Is Nothing Then
            Debug.Print "警告: シート「" & sheetNames(k) & "」が見つかりません。"
            GoTo NextPart
        End If
        
        Application.StatusBar = "転記処理中... (" & sheetNames(k) & ")"
        
        ' 各日付のデータを処理
        For i = 1 To sourceData.Rows.Count
            processedCount = processedCount + 1
            
            ' 日付を取得
            targetDate = sourceData.Cells(i, dateColIndex).Value
            
            ' 空白行はスキップ
            If IsEmpty(targetDate) Then
                GoTo NextDate
            End If
            
            ' 転記先シートで該当日付の行を検索
            foundRow = 0
            For j = 2 To targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
                If targetSheet.Cells(j, 1).Value = targetDate Then
                    foundRow = j
                    Exit For
                End If
            Next j
            
            ' 該当日付が見つからない場合は新規行追加
            If foundRow = 0 Then
                foundRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row + 1
                targetSheet.Cells(foundRow, 1).Value = targetDate
            End If
            
            ' データ転記
            ' 実績列の転記（品番に応じた列から）
            Dim colName As String
            colName = partNumbers(k) & "実績"
            On Error Resume Next
            Dim colIndex As Long
            colIndex = sourceTable.ListColumns(colName).Index
            If Err.Number = 0 Then
                If Not IsEmpty(sourceData.Cells(i, colIndex).Value) Then
                    targetSheet.Cells(foundRow, 2).Value = sourceData.Cells(i, colIndex).Value
                End If
            End If
            Err.Clear
            
            ' 不良列の転記
            colName = partNumbers(k) & "不良"
            colIndex = sourceTable.ListColumns(colName).Index
            If Err.Number = 0 Then
                If Not IsEmpty(sourceData.Cells(i, colIndex).Value) Then
                    targetSheet.Cells(foundRow, 3).Value = sourceData.Cells(i, colIndex).Value
                End If
            End If
            Err.Clear
            On Error GoTo ErrorHandler
            
            ' 進捗更新
            If processedCount Mod 10 = 0 Then
                Application.StatusBar = "転記処理中... (" & processedCount & "/" & (totalDates * UBound(partNumbers) + 1) & ")"
            End If
            
NextDate:
        Next i
        
NextPart:
        Set targetSheet = Nothing
    Next k
    
    ' 正常終了
    GoTo CleanupAndExit
    
ErrorHandler:
    ' エラー処理
    MsgBox "転記処理中にエラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "転記エラー"
    
CleanupAndExit:
    ' 後処理
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub