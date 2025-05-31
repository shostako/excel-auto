Attribute VB_Name = "m転記_シート加工品番別"
Option Explicit

Sub 転記_シート加工品番別()
    ' 高速化設定
    Application.ScreenUpdating = False
    Application.StatusBar = "加工品番別 転記処理を開始..."
    
    On Error GoTo ErrorHandler
    
    ' テーブル取得
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("加工品番別")
    
    Dim srcTable As ListObject, tgtTable As ListObject
    Set srcTable = ws.ListObjects("_加工品番別a")
    Set tgtTable = ws.ListObjects("_加工品番別b")
    
    ' データ範囲チェック
    If srcTable.DataBodyRange Is Nothing Then
        Application.StatusBar = "ソーステーブルにデータなし"
        GoTo Cleanup
    End If
    
    ' 必要な列インデックス取得
    Dim srcCols As Object
    Set srcCols = CreateObject("Scripting.Dictionary")
    srcCols("日付") = srcTable.ListColumns("日付").Index
    srcCols("通称") = srcTable.ListColumns("通称").Index
    srcCols("実績") = srcTable.ListColumns("実績").Index
    srcCols("不良") = srcTable.ListColumns("不良").Index
    srcCols("稼働時間") = srcTable.ListColumns("稼働時間").Index
    
    Dim tgtDateCol As Long
    tgtDateCol = tgtTable.ListColumns("日付").Index
    
    ' データ転記処理
    Dim srcData As Range, tgtData As Range
    Set srcData = srcTable.DataBodyRange
    Set tgtData = tgtTable.DataBodyRange
    
    If tgtData Is Nothing Then
        Application.StatusBar = "転記先テーブルが空"
        GoTo Cleanup
    End If
    
    ' 日付ごとの合計値を格納する辞書
    Dim dailyTotals As Object
    Set dailyTotals = CreateObject("Scripting.Dictionary")
    
    ' 日付・通称ごとの合計値を格納する辞書（これが修正のポイント）
    Dim dateNicknameTotals As Object
    Set dateNicknameTotals = CreateObject("Scripting.Dictionary")
    
    ' 日付ごとに出現した通称を記録する辞書（転記先クリア用）
    Dim dailyNicknames As Object
    Set dailyNicknames = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, j As Long
    Dim srcDate As Date, nickname As String
    Dim transferred As Long: transferred = 0
    Dim totalRows As Long: totalRows = srcData.Rows.Count
    
    ' まず、日付・通称ごとの合計値を計算
    Application.StatusBar = "データを集計中..."
    For i = 1 To totalRows
        srcDate = srcData.Cells(i, srcCols("日付")).Value
        nickname = Trim(srcData.Cells(i, srcCols("通称")).Value)
        
        If nickname <> "" Then
            Dim dateKey As String
            dateKey = Format(srcDate, "yyyy-mm-dd")
            
            ' 日付・通称をキーとした複合キー
            Dim compositeKey As String
            compositeKey = dateKey & "|" & nickname
            
            ' 日付・通称ごとの集計
            If Not dateNicknameTotals.Exists(compositeKey) Then
                dateNicknameTotals(compositeKey) = Array(0, 0, 0) ' 実績、不良、稼働時間の順
            End If
            
            Dim nicknameTotals As Variant
            nicknameTotals = dateNicknameTotals(compositeKey)
            nicknameTotals(0) = nicknameTotals(0) + srcData.Cells(i, srcCols("実績")).Value
            nicknameTotals(1) = nicknameTotals(1) + srcData.Cells(i, srcCols("不良")).Value
            nicknameTotals(2) = nicknameTotals(2) + srcData.Cells(i, srcCols("稼働時間")).Value
            dateNicknameTotals(compositeKey) = nicknameTotals
            
            ' 日付ごとの合計値も同時に計算
            If Not dailyTotals.Exists(dateKey) Then
                dailyTotals(dateKey) = Array(0, 0, 0) ' 実績、不良、稼働時間の順
                Set dailyNicknames(dateKey) = CreateObject("Scripting.Dictionary")
            End If
            
            Dim totals As Variant
            totals = dailyTotals(dateKey)
            totals(0) = totals(0) + srcData.Cells(i, srcCols("実績")).Value
            totals(1) = totals(1) + srcData.Cells(i, srcCols("不良")).Value
            totals(2) = totals(2) + srcData.Cells(i, srcCols("稼働時間")).Value
            dailyTotals(dateKey) = totals
            
            ' 通称を記録
            dailyNicknames(dateKey)(nickname) = True
        End If
    Next i
    
    ' 転記先テーブルを一旦クリア（該当する通称と合計のみ）
    Application.StatusBar = "転記先をクリア中..."
    For j = 1 To tgtData.Rows.Count
        Dim tgtDate As Date
        tgtDate = tgtData.Cells(j, tgtDateCol).Value
        dateKey = Format(tgtDate, "yyyy-mm-dd")
        
        ' 該当日付に通称が存在する場合のみクリア
        If dailyNicknames.Exists(dateKey) Then
            Dim nick As Variant
            For Each nick In dailyNicknames(dateKey).Keys
                ' 通称の実績、不良、稼働時間をクリア
                ClearValue tgtTable, tgtData, j, CStr(nick) & "日実績"
                ClearValue tgtTable, tgtData, j, CStr(nick) & "日不良実績"
                ClearValue tgtTable, tgtData, j, CStr(nick) & "日稼働時間"
            Next nick
            
            ' 合計列もクリア
            ClearValue tgtTable, tgtData, j, "合計日実績"
            ClearValue tgtTable, tgtData, j, "合計日不良実績"
            ClearValue tgtTable, tgtData, j, "合計日稼働時間"
        End If
    Next j
    
    ' 集計されたデータを転記
    Application.StatusBar = "集計データを転記中..."
    Dim key As Variant
    Dim processedKeys As Long: processedKeys = 0
    Dim totalKeys As Long: totalKeys = dateNicknameTotals.Count
    
    For Each key In dateNicknameTotals.Keys
        processedKeys = processedKeys + 1
        
        ' 進捗表示
        If processedKeys Mod 10 = 0 Or processedKeys = totalKeys Then
            Application.StatusBar = "加工品番別 転記処理中... " & Format(processedKeys / totalKeys, "0%") & _
                                  " (" & processedKeys & "/" & totalKeys & "件)"
            DoEvents
        End If
        
        ' キーを分解して日付と通称を取得
        Dim keyParts() As String
        keyParts = Split(key, "|")
        dateKey = keyParts(0)
        nickname = keyParts(1)
        
        ' 転記先の日付検索
        For j = 1 To tgtData.Rows.Count
            If Format(tgtData.Cells(j, tgtDateCol).Value, "yyyy-mm-dd") = dateKey Then
                ' 集計されたデータを転記
                nicknameTotals = dateNicknameTotals(key)
                TransferValue tgtTable, tgtData, j, nickname & "日実績", nicknameTotals(0)
                TransferValue tgtTable, tgtData, j, nickname & "日不良実績", nicknameTotals(1)
                TransferValue tgtTable, tgtData, j, nickname & "日稼働時間", nicknameTotals(2)
                transferred = transferred + 1
                Exit For
            End If
        Next j
    Next key
    
    ' 合計値の転記
    Application.StatusBar = "合計値を転記中..."
    Dim totalTransferred As Long: totalTransferred = 0
    For j = 1 To tgtData.Rows.Count
        tgtDate = tgtData.Cells(j, tgtDateCol).Value
        dateKey = Format(tgtDate, "yyyy-mm-dd")
        
        If dailyTotals.Exists(dateKey) Then
            totals = dailyTotals(dateKey)
            
            ' 合計値を転記
            TransferValue tgtTable, tgtData, j, "合計日実績", totals(0)
            TransferValue tgtTable, tgtData, j, "合計日不良実績", totals(1)
            TransferValue tgtTable, tgtData, j, "合計日稼働時間", totals(2)
            totalTransferred = totalTransferred + 1
        End If
    Next j
    
    ' 完了時のステータスバー表示
    Application.StatusBar = "加工品番別転記完了: " & transferred & "件の品番別データ、" & _
                           totalTransferred & "件の合計データを転記"
    
    ' 1秒待機してからクリア
    Application.Wait Now + TimeValue("0:00:01")
    Application.StatusBar = False
    
Cleanup:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    ' エラー時だけメッセージボックス
    MsgBox "加工品番別 転記処理でエラー発生" & vbCrLf & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "転記エラー"
    Application.StatusBar = False
    Resume Cleanup
End Sub

' 値転記用ヘルパー関数
Private Sub TransferValue(tbl As ListObject, data As Range, _
                         row As Long, colName As String, val As Variant)
    On Error Resume Next
    Dim colIdx As Long
    colIdx = tbl.ListColumns(colName).Index
    If colIdx > 0 Then data.Cells(row, colIdx).Value = val
    On Error GoTo 0
End Sub

' 値クリア用ヘルパー関数
Private Sub ClearValue(tbl As ListObject, data As Range, _
                      row As Long, colName As String)
    On Error Resume Next
    Dim colIdx As Long
    colIdx = tbl.ListColumns(colName).Index
    If colIdx > 0 Then data.Cells(row, colIdx).ClearContents
    On Error GoTo 0
End Sub

