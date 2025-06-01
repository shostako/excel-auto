Attribute VB_Name = "m転記_シート加工作業者別"
Option Explicit

' ==========================================
' 加工作業者別からシートへの転記マクロ
' 「_加工作業者別a」テーブルから各作業者シートへデータを転記
' ==========================================
Sub 転記_シート加工作業者別()
    ' ==========================================
    ' 変数宣言
    ' ==========================================
    Dim wsSource As Worksheet
    Dim sourceTable As ListObject
    Dim sourceData As Range
    Dim i As Long, j As Long
    Dim workerName As String
    Dim targetSheet As Worksheet
    Dim lastRow As Long
    Dim foundRow As Long
    Dim targetDate As Date
    Dim processedCount As Long
    Dim totalWorkers As Long
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' ==========================================
    ' 高速化設定
    ' ==========================================
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 進捗表示開始
    Application.StatusBar = "加工作業者別シート転記処理を開始します..."
    
    ' ==========================================
    ' ソースシート・テーブル取得
    ' ==========================================
    ' 加工作業者別シート取得
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("加工作業者別")
    If wsSource Is Nothing Then
        MsgBox "「加工作業者別」シートが見つかりません。", vbCritical, "シートエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' ソーステーブル取得
    On Error Resume Next
    Set sourceTable = wsSource.ListObjects("_加工作業者別a")
    If sourceTable Is Nothing Then
        MsgBox "「_加工作業者別a」テーブルが見つかりません。", vbCritical, "テーブルエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' データ範囲取得
    If sourceTable.DataBodyRange Is Nothing Then
        MsgBox "「_加工作業者別a」テーブルにデータがありません。", vbInformation, "データなし"
        GoTo CleanupAndExit
    End If
    Set sourceData = sourceTable.DataBodyRange
    
    ' ==========================================
    ' メイン処理: 各作業者のデータを転記
    ' ==========================================
    ' 作業者数をカウント（進捗表示用）
    totalWorkers = sourceData.Rows.Count
    processedCount = 0
    
    ' 各行（作業者）のデータを処理
    For i = 1 To sourceData.Rows.Count
        ' 日付と作業者名を取得
        targetDate = sourceData.Cells(i, 1).Value  ' 日付列
        workerName = sourceData.Cells(i, 2).Value  ' 作業者列
        
        ' 空白行はスキップ
        If workerName = "" Or IsEmpty(workerName) Then
            GoTo NextWorker
        End If
        
        processedCount = processedCount + 1
        Application.StatusBar = "転記処理中... (" & processedCount & "/" & totalWorkers & ") " & workerName
        
        ' 作業者シートの存在確認
        On Error Resume Next
        Set targetSheet = ThisWorkbook.Worksheets(workerName)
        On Error GoTo ErrorHandler
        
        If targetSheet Is Nothing Then
            ' シートが存在しない場合はスキップ（エラーにはしない）
            Debug.Print "警告: 作業者「" & workerName & "」のシートが見つかりません。"
            GoTo NextWorker
        End If
        
        ' 転記先シートで該当日付の行を検索
        lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
        foundRow = 0
        
        For j = 2 To lastRow  ' ヘッダー行をスキップ
            If targetSheet.Cells(j, 1).Value = targetDate Then
                foundRow = j
                Exit For
            End If
        Next j
        
        ' 該当日付が見つからない場合は新規行追加
        If foundRow = 0 Then
            foundRow = lastRow + 1
            targetSheet.Cells(foundRow, 1).Value = targetDate
        End If
        
        ' データ転記（3列目以降のデータ）
        ' 稼動時間
        If Not IsEmpty(sourceData.Cells(i, 3).Value) Then
            targetSheet.Cells(foundRow, 2).Value = sourceData.Cells(i, 3).Value
        End If
        
        ' 段取時間
        If Not IsEmpty(sourceData.Cells(i, 4).Value) Then
            targetSheet.Cells(foundRow, 3).Value = sourceData.Cells(i, 4).Value
        End If
        
        ' 実績
        If Not IsEmpty(sourceData.Cells(i, 5).Value) Then
            targetSheet.Cells(foundRow, 4).Value = sourceData.Cells(i, 5).Value
        End If
        
        ' 不良
        If Not IsEmpty(sourceData.Cells(i, 6).Value) Then
            targetSheet.Cells(foundRow, 5).Value = sourceData.Cells(i, 6).Value
        End If
        
NextWorker:
        Set targetSheet = Nothing
    Next i
    
    ' 正常終了
    GoTo CleanupAndExit
    
ErrorHandler:
    ' エラー処理
    MsgBox "転記処理中にエラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "処理中の作業者: " & workerName, vbCritical, "転記エラー"
    
CleanupAndExit:
    ' 後処理
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub