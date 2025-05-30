Attribute VB_Name = "m転記_シートモールFR別"
Option Explicit

' モールFR別転記マクロ
' 「_モールFR別a」テーブルから「_モールFR別b」テーブルへデータを転記
Sub 転記_シートモールFR別()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim srcTable As ListObject
    Dim tgtTable As ListObject
    Dim srcData As Range
    Dim tgtData As Range
    Dim srcCols As Object
    Dim tgtCols As Object
    
    ' 基本設定
    Set wb = ThisWorkbook
    
    ' ステータスバー表示
    Application.StatusBar = "モールFR別転記処理を開始..."
    
    On Error GoTo ErrorHandler
    
    ' シート取得（シート名は「モールFR別」と想定）
    Set ws = wb.Worksheets("モールFR別")
    
    ' テーブル取得
    Set srcTable = ws.ListObjects("_モールFR別a")
    Set tgtTable = ws.ListObjects("_モールFR別b")
    
    ' データ範囲チェック
    If srcTable.DataBodyRange Is Nothing Then
        Application.StatusBar = "ソーステーブルにデータがありません"
        GoTo Cleanup
    End If
    
    If tgtTable.DataBodyRange Is Nothing Then
        Application.StatusBar = "転記先テーブルにデータがありません"
        GoTo Cleanup
    End If
    
    ' ソーステーブルの列インデックス取得
    Set srcCols = CreateObject("Scripting.Dictionary")
    srcCols("日付") = srcTable.ListColumns("日付").Index
    srcCols("F/R") = srcTable.ListColumns("F/R").Index
    srcCols("実績") = srcTable.ListColumns("実績").Index
    srcCols("不良") = srcTable.ListColumns("不良").Index
    srcCols("稼働時間") = srcTable.ListColumns("稼働時間").Index
    
    ' 転記先の列インデックス取得
    Set tgtCols = CreateObject("Scripting.Dictionary")
    tgtCols("日付") = tgtTable.ListColumns("日付").Index
    
    ' モールF列
    tgtCols("モールF日実績") = GetColumnIndexSafe(tgtTable, "モールF日実績")
    tgtCols("モールF日不良数") = GetColumnIndexSafe(tgtTable, "モールF日不良数")
    tgtCols("モールF日稼働時間") = GetColumnIndexSafe(tgtTable, "モールF日稼働時間")
    
    ' モールR列
    tgtCols("モールR日実績") = GetColumnIndexSafe(tgtTable, "モールR日実績")
    tgtCols("モールR日不良数") = GetColumnIndexSafe(tgtTable, "モールR日不良数")
    tgtCols("モールR日稼働時間") = GetColumnIndexSafe(tgtTable, "モールR日稼働時間")
    
    ' データ範囲取得
    Set srcData = srcTable.DataBodyRange
    Set tgtData = tgtTable.DataBodyRange
    
    Dim i As Long, j As Long
    Dim srcDate As Date, frType As String
    Dim transferred As Long: transferred = 0
    Dim totalRows As Long: totalRows = srcData.Rows.Count
    
    ' 転記先テーブルを全クリア
    Application.StatusBar = "転記先をクリア中..."
    For j = 1 To tgtData.Rows.Count
        ' モールF列のクリア
        If tgtCols("モールF日実績") > 0 Then tgtData.Cells(j, tgtCols("モールF日実績")).ClearContents
        If tgtCols("モールF日不良数") > 0 Then tgtData.Cells(j, tgtCols("モールF日不良数")).ClearContents
        If tgtCols("モールF日稼働時間") > 0 Then tgtData.Cells(j, tgtCols("モールF日稼働時間")).ClearContents
        
        ' モールR列のクリア
        If tgtCols("モールR日実績") > 0 Then tgtData.Cells(j, tgtCols("モールR日実績")).ClearContents
        If tgtCols("モールR日不良数") > 0 Then tgtData.Cells(j, tgtCols("モールR日不良数")).ClearContents
        If tgtCols("モールR日稼働時間") > 0 Then tgtData.Cells(j, tgtCols("モールR日稼働時間")).ClearContents
    Next j
    
    ' データの転記
    For i = 1 To totalRows
        ' 進捗表示（10行ごとに更新）
        If i Mod 10 = 0 Or i = totalRows Then
            Application.StatusBar = "モールFR別転記処理中... " & Format(i / totalRows, "0%") & _
                                  " (" & i & "/" & totalRows & "行)"
            DoEvents ' 画面更新
        End If
        
        ' ソースデータ取得
        srcDate = srcData.Cells(i, srcCols("日付")).Value
        frType = Trim(srcData.Cells(i, srcCols("F/R")).Value)
        
        ' 転記先の日付検索
        For j = 1 To tgtData.Rows.Count
            If tgtData.Cells(j, tgtCols("日付")).Value = srcDate Then
                ' F/Rタイプに応じて転記
                If frType = "F" Then
                    ' モールF列への転記
                    If tgtCols("モールF日実績") > 0 Then
                        tgtData.Cells(j, tgtCols("モールF日実績")).Value = srcData.Cells(i, srcCols("実績")).Value
                    End If
                    If tgtCols("モールF日不良数") > 0 Then
                        tgtData.Cells(j, tgtCols("モールF日不良数")).Value = srcData.Cells(i, srcCols("不良")).Value
                    End If
                    If tgtCols("モールF日稼働時間") > 0 Then
                        tgtData.Cells(j, tgtCols("モールF日稼働時間")).Value = srcData.Cells(i, srcCols("稼働時間")).Value
                    End If
                    transferred = transferred + 1
                    
                ElseIf frType = "R" Then
                    ' モールR列への転記
                    If tgtCols("モールR日実績") > 0 Then
                        tgtData.Cells(j, tgtCols("モールR日実績")).Value = srcData.Cells(i, srcCols("実績")).Value
                    End If
                    If tgtCols("モールR日不良数") > 0 Then
                        tgtData.Cells(j, tgtCols("モールR日不良数")).Value = srcData.Cells(i, srcCols("不良")).Value
                    End If
                    If tgtCols("モールR日稼働時間") > 0 Then
                        tgtData.Cells(j, tgtCols("モールR日稼働時間")).Value = srcData.Cells(i, srcCols("稼働時間")).Value
                    End If
                    transferred = transferred + 1
                End If
                
                Exit For ' 日付が見つかったら次の行へ
            End If
        Next j
    Next i
    
    ' 小数点以下2桁の書式設定（稼働時間列のみ）
    Application.StatusBar = "書式設定中..."
    For j = 1 To tgtData.Rows.Count
        ' モールF日稼働時間
        If tgtCols("モールF日稼働時間") > 0 Then
            tgtData.Cells(j, tgtCols("モールF日稼働時間")).NumberFormatLocal = "0.00"
        End If
        
        ' モールR日稼働時間
        If tgtCols("モールR日稼働時間") > 0 Then
            tgtData.Cells(j, tgtCols("モールR日稼働時間")).NumberFormatLocal = "0.00"
        End If
    Next j
    
    ' 完了処理
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ' エラー時の処理
    Application.StatusBar = False
    MsgBox "モールFR別転記処理でエラーが発生しました" & vbCrLf & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "転記エラー"
    
Cleanup:
    ' 後処理
    Application.StatusBar = False
    Set srcCols = Nothing
    Set tgtCols = Nothing
    Set srcData = Nothing
    Set tgtData = Nothing
    Set srcTable = Nothing
    Set tgtTable = Nothing
    Set ws = Nothing
    Set wb = Nothing
End Sub

' 列インデックスを安全に取得するヘルパー関数
Private Function GetColumnIndexSafe(tbl As ListObject, colName As String) As Long
    On Error Resume Next
    GetColumnIndexSafe = tbl.ListColumns(colName).Index
    If Err.Number <> 0 Then
        GetColumnIndexSafe = 0
        Debug.Print "警告: 列「" & colName & "」が見つかりません"
    End If
    On Error GoTo 0
End Function