Attribute VB_Name = "m転記_シート成形号機別"
Option Explicit

Sub 転記_シート成形号機別()
    ' 高速化設定
    Application.ScreenUpdating = False
    Application.StatusBar = "成形号機別転記処理を開始..."
    
    On Error GoTo ErrorHandler
    
    ' テーブル取得
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("成形号機別")
    
    Dim srcTable As ListObject, tgtTable As ListObject
    Set srcTable = ws.ListObjects("_成形号機別a")
    Set tgtTable = ws.ListObjects("_成形号機別b")
    
    ' データ範囲チェック
    If srcTable.DataBodyRange Is Nothing Then
        Application.StatusBar = "ソーステーブルにデータなし"
        GoTo Cleanup
    End If
    
    ' 必要な列インデックス取得（ソース）
    Dim srcCols As Object
    Set srcCols = CreateObject("Scripting.Dictionary")
    srcCols("日付") = srcTable.ListColumns("日付").Index
    srcCols("機械") = srcTable.ListColumns("機械").Index
    srcCols("実績") = srcTable.ListColumns("実績").Index
    srcCols("不良") = srcTable.ListColumns("不良").Index
    srcCols("稼働時間") = srcTable.ListColumns("稼働時間").Index
    
    ' 転記先の列インデックス取得
    Dim tgtCols As Object
    Set tgtCols = CreateObject("Scripting.Dictionary")
    tgtCols("日付") = tgtTable.ListColumns("日付").Index
    
    ' 1号機用列
    tgtCols("1号機日実績") = GetColumnIndexSafe(tgtTable, "1号機日実績")
    tgtCols("1号機日不良実績") = GetColumnIndexSafe(tgtTable, "1号機日不良実績")
    tgtCols("1号機日稼働時間") = GetColumnIndexSafe(tgtTable, "1号機日稼働時間")
    
    ' 2号機用列
    tgtCols("2号機日実績") = GetColumnIndexSafe(tgtTable, "2号機日実績")
    tgtCols("2号機日不良実績") = GetColumnIndexSafe(tgtTable, "2号機日不良実績")
    tgtCols("2号機日稼働時間") = GetColumnIndexSafe(tgtTable, "2号機日稼働時間")
    
    ' 3号機用列
    tgtCols("3号機日実績") = GetColumnIndexSafe(tgtTable, "3号機日実績")
    tgtCols("3号機日不良実績") = GetColumnIndexSafe(tgtTable, "3号機日不良実績")
    tgtCols("3号機日稼働時間") = GetColumnIndexSafe(tgtTable, "3号機日稼働時間")
    
    ' 4号機用列
    tgtCols("4号機日実績") = GetColumnIndexSafe(tgtTable, "4号機日実績")
    tgtCols("4号機日不良実績") = GetColumnIndexSafe(tgtTable, "4号機日不良実績")
    tgtCols("4号機日稼働時間") = GetColumnIndexSafe(tgtTable, "4号機日稼働時間")
    
    ' 5号機用列
    tgtCols("5号機日実績") = GetColumnIndexSafe(tgtTable, "5号機日実績")
    tgtCols("5号機日不良実績") = GetColumnIndexSafe(tgtTable, "5号機日不良実績")
    tgtCols("5号機日稼働時間") = GetColumnIndexSafe(tgtTable, "5号機日稼働時間")
    
    ' データ転記処理
    Dim srcData As Range, tgtData As Range
    Set srcData = srcTable.DataBodyRange
    Set tgtData = tgtTable.DataBodyRange
    
    If tgtData Is Nothing Then
        Application.StatusBar = "転記先テーブルが空"
        GoTo Cleanup
    End If
    
    Dim i As Long, j As Long
    Dim srcDate As Date, machine As String
    Dim transferred As Long: transferred = 0
    Dim totalRows As Long: totalRows = srcData.Rows.Count
    
    ' 転記先テーブルを一旦クリア
    Application.StatusBar = "転記先をクリア中..."
    For j = 1 To tgtData.Rows.Count
        ' 1号機列のクリア
        If tgtCols("1号機日実績") > 0 Then tgtData.Cells(j, tgtCols("1号機日実績")).ClearContents
        If tgtCols("1号機日不良実績") > 0 Then tgtData.Cells(j, tgtCols("1号機日不良実績")).ClearContents
        If tgtCols("1号機日稼働時間") > 0 Then tgtData.Cells(j, tgtCols("1号機日稼働時間")).ClearContents
        
        ' 2号機列のクリア
        If tgtCols("2号機日実績") > 0 Then tgtData.Cells(j, tgtCols("2号機日実績")).ClearContents
        If tgtCols("2号機日不良実績") > 0 Then tgtData.Cells(j, tgtCols("2号機日不良実績")).ClearContents
        If tgtCols("2号機日稼働時間") > 0 Then tgtData.Cells(j, tgtCols("2号機日稼働時間")).ClearContents
        
        ' 3号機列のクリア
        If tgtCols("3号機日実績") > 0 Then tgtData.Cells(j, tgtCols("3号機日実績")).ClearContents
        If tgtCols("3号機日不良実績") > 0 Then tgtData.Cells(j, tgtCols("3号機日不良実績")).ClearContents
        If tgtCols("3号機日稼働時間") > 0 Then tgtData.Cells(j, tgtCols("3号機日稼働時間")).ClearContents
        
        ' 4号機列のクリア
        If tgtCols("4号機日実績") > 0 Then tgtData.Cells(j, tgtCols("4号機日実績")).ClearContents
        If tgtCols("4号機日不良実績") > 0 Then tgtData.Cells(j, tgtCols("4号機日不良実績")).ClearContents
        If tgtCols("4号機日稼働時間") > 0 Then tgtData.Cells(j, tgtCols("4号機日稼働時間")).ClearContents
        
        ' 5号機列のクリア
        If tgtCols("5号機日実績") > 0 Then tgtData.Cells(j, tgtCols("5号機日実績")).ClearContents
        If tgtCols("5号機日不良実績") > 0 Then tgtData.Cells(j, tgtCols("5号機日不良実績")).ClearContents
        If tgtCols("5号機日稼働時間") > 0 Then tgtData.Cells(j, tgtCols("5号機日稼働時間")).ClearContents
    Next j
    
    ' 機械別データの転記
    For i = 1 To totalRows
        ' 進捗表示（10行ごとに更新して処理速度優先）
        If i Mod 10 = 0 Or i = totalRows Then
            Application.StatusBar = "成形号機別転記処理中... " & Format(i / totalRows, "0%") & _
                                  " (" & i & "/" & totalRows & "行)"
            DoEvents ' 画面更新
        End If
        
        ' ソースデータ取得
        srcDate = srcData.Cells(i, srcCols("日付")).Value
        machine = Trim(srcData.Cells(i, srcCols("機械")).Value)
        
        ' 機械が対象外ならスキップ
        If machine <> "SS01" And machine <> "SS02" And machine <> "SS03" And _
           machine <> "SS04" And machine <> "SS05" Then
            GoTo NextRow
        End If
        
        ' 転記先の日付検索
        For j = 1 To tgtData.Rows.Count
            If tgtData.Cells(j, tgtCols("日付")).Value = srcDate Then
                ' 機械に応じて転記
                Select Case machine
                    Case "SS01"
                        ' 1号機の転記
                        If tgtCols("1号機日実績") > 0 Then
                            tgtData.Cells(j, tgtCols("1号機日実績")).Value = srcData.Cells(i, srcCols("実績")).Value
                        End If
                        If tgtCols("1号機日不良実績") > 0 Then
                            tgtData.Cells(j, tgtCols("1号機日不良実績")).Value = srcData.Cells(i, srcCols("不良")).Value
                        End If
                        If tgtCols("1号機日稼働時間") > 0 Then
                            tgtData.Cells(j, tgtCols("1号機日稼働時間")).Value = srcData.Cells(i, srcCols("稼働時間")).Value
                        End If
                        transferred = transferred + 1
                        
                    Case "SS02"
                        ' 2号機の転記
                        If tgtCols("2号機日実績") > 0 Then
                            tgtData.Cells(j, tgtCols("2号機日実績")).Value = srcData.Cells(i, srcCols("実績")).Value
                        End If
                        If tgtCols("2号機日不良実績") > 0 Then
                            tgtData.Cells(j, tgtCols("2号機日不良実績")).Value = srcData.Cells(i, srcCols("不良")).Value
                        End If
                        If tgtCols("2号機日稼働時間") > 0 Then
                            tgtData.Cells(j, tgtCols("2号機日稼働時間")).Value = srcData.Cells(i, srcCols("稼働時間")).Value
                        End If
                        transferred = transferred + 1
                        
                    Case "SS03"
                        ' 3号機の転記
                        If tgtCols("3号機日実績") > 0 Then
                            tgtData.Cells(j, tgtCols("3号機日実績")).Value = srcData.Cells(i, srcCols("実績")).Value
                        End If
                        If tgtCols("3号機日不良実績") > 0 Then
                            tgtData.Cells(j, tgtCols("3号機日不良実績")).Value = srcData.Cells(i, srcCols("不良")).Value
                        End If
                        If tgtCols("3号機日稼働時間") > 0 Then
                            tgtData.Cells(j, tgtCols("3号機日稼働時間")).Value = srcData.Cells(i, srcCols("稼働時間")).Value
                        End If
                        transferred = transferred + 1
                        
                    Case "SS04"
                        ' 4号機の転記
                        If tgtCols("4号機日実績") > 0 Then
                            tgtData.Cells(j, tgtCols("4号機日実績")).Value = srcData.Cells(i, srcCols("実績")).Value
                        End If
                        If tgtCols("4号機日不良実績") > 0 Then
                            tgtData.Cells(j, tgtCols("4号機日不良実績")).Value = srcData.Cells(i, srcCols("不良")).Value
                        End If
                        If tgtCols("4号機日稼働時間") > 0 Then
                            tgtData.Cells(j, tgtCols("4号機日稼働時間")).Value = srcData.Cells(i, srcCols("稼働時間")).Value
                        End If
                        transferred = transferred + 1
                        
                    Case "SS05"
                        ' 5号機の転記
                        If tgtCols("5号機日実績") > 0 Then
                            tgtData.Cells(j, tgtCols("5号機日実績")).Value = srcData.Cells(i, srcCols("実績")).Value
                        End If
                        If tgtCols("5号機日不良実績") > 0 Then
                            tgtData.Cells(j, tgtCols("5号機日不良実績")).Value = srcData.Cells(i, srcCols("不良")).Value
                        End If
                        If tgtCols("5号機日稼働時間") > 0 Then
                            tgtData.Cells(j, tgtCols("5号機日稼働時間")).Value = srcData.Cells(i, srcCols("稼働時間")).Value
                        End If
                        transferred = transferred + 1
                End Select
                
                Exit For ' 日付が見つかったら次の行へ
            End If
        Next j
        
NextRow:
    Next i
    
    ' 完了時のステータスバー表示
    Application.StatusBar = "成形号機別転記完了: " & transferred & "件のデータを転記"
    
    ' 1秒待機してからクリア
    Application.Wait Now + TimeValue("0:00:01")
    Application.StatusBar = False
    
Cleanup:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    ' エラー時だけメッセージボックス
    MsgBox "成形号機別転記処理でエラー発生" & vbCrLf & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "転記エラー"
    Application.StatusBar = False
    Resume Cleanup
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

