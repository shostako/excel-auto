Attribute VB_Name = "m転記_シート加工作業者別"
Sub 転記_シート加工作業者別()
    ' 変数宣言
    Dim ws As Worksheet
    Dim sourceTable As ListObject
    Dim TargetTable As ListObject
    Dim sourceData As Range
    Dim TargetData As Range
    Dim i As Long, j As Long
    Dim sourceRow As Long, targetRow As Long
    Dim sourceDate As Date, targetDate As Date
    Dim workerName As String
    Dim workTimeColName As String, resultColName As String
    Dim workTimeCol As ListColumn, resultCol As ListColumn
    Dim workTimeColIndex As Long, resultColIndex As Long
    Dim dateColSourceIndex As Long, dateColTargetIndex As Long
    Dim workerColIndex As Long, workTimeSourceColIndex As Long, resultSourceColIndex As Long
    Dim totalRows As Long
    Dim processedCount As Long
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' 進捗表示開始
    Application.StatusBar = "転記処理を開始します..."
    Application.ScreenUpdating = False
    
    ' ワークシート取得
    Set ws = ThisWorkbook.Worksheets("加工作業者別")
    
    ' テーブル取得
    Set sourceTable = ws.ListObjects("_加工作業者別a")
    Set TargetTable = ws.ListObjects("_加工作業者別b")
    
    ' データ範囲取得（ヘッダー除く）
    Set sourceData = sourceTable.DataBodyRange
    Set TargetData = TargetTable.DataBodyRange
    
    ' 列インデックス取得
    dateColSourceIndex = sourceTable.ListColumns("日付").Index
    dateColTargetIndex = TargetTable.ListColumns("日付").Index
    workerColIndex = sourceTable.ListColumns("作業者").Index
    workTimeSourceColIndex = sourceTable.ListColumns("稼働時間").Index
    resultSourceColIndex = sourceTable.ListColumns("実績").Index
    
    ' 総行数取得
    totalRows = sourceData.Rows.Count
    processedCount = 0
    
    ' ソーステーブルの各行を処理
    For i = 1 To totalRows
        ' 進捗表示更新
        processedCount = processedCount + 1
        Application.StatusBar = "転記処理中... (" & processedCount & "/" & totalRows & ")"
        
        ' ソースデータ取得
        sourceDate = sourceData.Cells(i, dateColSourceIndex).Value
        workerName = Trim(sourceData.Cells(i, workerColIndex).Value)
        
        ' 作業者名が空白の場合はスキップ
        If workerName = "" Then
            GoTo NextSourceRow
        End If
        
        ' 転記先の対応日付行を検索
        targetRow = 0
        For j = 1 To TargetData.Rows.Count
            If TargetData.Cells(j, dateColTargetIndex).Value = sourceDate Then
                targetRow = j
                Exit For
            End If
        Next j
        
        ' 対応する日付が見つからない場合はスキップ
        If targetRow = 0 Then
            GoTo NextSourceRow
        End If
        
        ' 稼働時間転記処理
        workTimeColName = workerName & "稼働時間"
        Set workTimeCol = Nothing
        
        ' 稼働時間列の存在確認
        For Each workTimeCol In TargetTable.ListColumns
            If workTimeCol.Name = workTimeColName Then
                workTimeColIndex = workTimeCol.Index
                ' 稼働時間値を転記
                TargetData.Cells(targetRow, workTimeColIndex).Value = sourceData.Cells(i, workTimeSourceColIndex).Value
                Exit For
            End If
        Next workTimeCol
        
        ' 実績転記処理
        resultColName = workerName & "実績"
        Set resultCol = Nothing
        
        ' 実績列の存在確認
        For Each resultCol In TargetTable.ListColumns
            If resultCol.Name = resultColName Then
                resultColIndex = resultCol.Index
                ' 実績値を転記
                TargetData.Cells(targetRow, resultColIndex).Value = sourceData.Cells(i, resultSourceColIndex).Value
                Exit For
            End If
        Next resultCol
        
NextSourceRow:
    Next i
    
    ' 処理完了
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ' エラー時の処理
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "転記処理中にエラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "転記エラー"
    
End Sub

