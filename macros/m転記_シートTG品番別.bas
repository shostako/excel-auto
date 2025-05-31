Attribute VB_Name = "m転記_シートTG品番別"
Option Explicit

Sub 転記_シートTG品番別()
    ' 高速化設定
    Application.ScreenUpdating = False
    Application.StatusBar = "TG品番別転記処理を開始..."
    
    On Error GoTo ErrorHandler
    
    ' テーブル取得
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("TG品番別")
    
    Dim srcTable As ListObject, tgtTable As ListObject
    Set srcTable = ws.ListObjects("_TG品番別a")
    Set tgtTable = ws.ListObjects("_TG品番別b")
    
    ' データ範囲チェック
    If srcTable.DataBodyRange Is Nothing Then
        Application.StatusBar = "ソーステーブルにデータなし"
        GoTo Cleanup
    End If
    
    ' 必要な列インデックス取得
    Dim srcCols As Object
    Set srcCols = CreateObject("Scripting.Dictionary")
    srcCols("日付") = srcTable.ListColumns("日付").Index
    srcCols("品番") = srcTable.ListColumns("品番").Index
    srcCols("実績") = srcTable.ListColumns("実績").Index
    srcCols("不良") = srcTable.ListColumns("不良").Index
    srcCols("稼働時間") = srcTable.ListColumns("稼働時間").Index
    
    ' 転記先の列インデックス取得
    Dim tgtCols As Object
    Set tgtCols = CreateObject("Scripting.Dictionary")
    tgtCols("日付") = tgtTable.ListColumns("日付").Index
    
    ' RH用列
    tgtCols("RH日実績") = GetColumnIndexSafe(tgtTable, "RH日実績")
    tgtCols("RH日不良実績") = GetColumnIndexSafe(tgtTable, "RH日不良実績")
    tgtCols("RH日稼働時間") = GetColumnIndexSafe(tgtTable, "RH日稼働時間")
    
    ' LH用列
    tgtCols("LH日実績") = GetColumnIndexSafe(tgtTable, "LH日実績")
    tgtCols("LH日不良実績") = GetColumnIndexSafe(tgtTable, "LH日不良実績")
    tgtCols("LH日稼働時間") = GetColumnIndexSafe(tgtTable, "LH日稼働時間")
    
    ' 合計用列
    tgtCols("合計日実績") = GetColumnIndexSafe(tgtTable, "合計日実績")
    tgtCols("合計日不良実績") = GetColumnIndexSafe(tgtTable, "合計日不良実績")
    tgtCols("合計日稼働時間") = GetColumnIndexSafe(tgtTable, "合計日稼働時間")
    
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
    
    Dim i As Long, j As Long
    Dim srcDate As Date, hinban As String
    Dim transferred As Long: transferred = 0
    Dim totalRows As Long: totalRows = srcData.Rows.Count
    
    ' まず、日付ごとの合計を計算
    Application.StatusBar = "合計値を計算中..."
    For i = 1 To totalRows
        srcDate = srcData.Cells(i, srcCols("日付")).Value
        hinban = Trim(srcData.Cells(i, srcCols("品番")).Value)
        
        ' 対象品番のみ処理
        If hinban = "53827-60050" Or hinban = "53828-60080" Then
            Dim dateKey As String
            dateKey = Format(srcDate, "yyyy-mm-dd")
            
            ' 日付キーが存在しない場合は初期化
            If Not dailyTotals.Exists(dateKey) Then
                dailyTotals(dateKey) = Array(0, 0, 0) ' 実績、不良、稼働時間の順
            End If
            
            ' 合計値を加算
            Dim totals As Variant
            totals = dailyTotals(dateKey)
            totals(0) = totals(0) + srcData.Cells(i, srcCols("実績")).Value
            totals(1) = totals(1) + srcData.Cells(i, srcCols("不良")).Value
            totals(2) = totals(2) + srcData.Cells(i, srcCols("稼働時間")).Value
            dailyTotals(dateKey) = totals
        End If
    Next i
    
    ' 転記先テーブルを一旦クリア（品番別データ）
    Application.StatusBar = "転記先をクリア中..."
    For j = 1 To tgtData.Rows.Count
        ' RH列のクリア
        If tgtCols("RH日実績") > 0 Then tgtData.Cells(j, tgtCols("RH日実績")).ClearContents
        If tgtCols("RH日不良実績") > 0 Then tgtData.Cells(j, tgtCols("RH日不良実績")).ClearContents
        If tgtCols("RH日稼働時間") > 0 Then tgtData.Cells(j, tgtCols("RH日稼働時間")).ClearContents
        
        ' LH列のクリア
        If tgtCols("LH日実績") > 0 Then tgtData.Cells(j, tgtCols("LH日実績")).ClearContents
        If tgtCols("LH日不良実績") > 0 Then tgtData.Cells(j, tgtCols("LH日不良実績")).ClearContents
        If tgtCols("LH日稼働時間") > 0 Then tgtData.Cells(j, tgtCols("LH日稼働時間")).ClearContents
        
        ' 合計列のクリア
        If tgtCols("合計日実績") > 0 Then tgtData.Cells(j, tgtCols("合計日実績")).ClearContents
        If tgtCols("合計日不良実績") > 0 Then tgtData.Cells(j, tgtCols("合計日不良実績")).ClearContents
        If tgtCols("合計日稼働時間") > 0 Then tgtData.Cells(j, tgtCols("合計日稼働時間")).ClearContents
    Next j
    
    ' 品番別データの転記
    For i = 1 To totalRows
        ' 進捗表示（10行ごとに更新して処理速度優先）
        If i Mod 10 = 0 Or i = totalRows Then
            Application.StatusBar = "TG品番別転記処理中... " & Format(i / totalRows, "0%") & _
                                  " (" & i & "/" & totalRows & "行)"
            DoEvents ' 画面更新
        End If
        
        ' ソースデータ取得
        srcDate = srcData.Cells(i, srcCols("日付")).Value
        hinban = Trim(srcData.Cells(i, srcCols("品番")).Value)
        
        ' 品番が対象外ならスキップ
        If hinban <> "53827-60050" And hinban <> "53828-60080" Then
            GoTo NextRow
        End If
        
        ' 転記先の日付検索
        For j = 1 To tgtData.Rows.Count
            If tgtData.Cells(j, tgtCols("日付")).Value = srcDate Then
                ' 品番に応じて転記
                If hinban = "53827-60050" Then
                    ' RH品番の転記
                    If tgtCols("RH日実績") > 0 Then
                        tgtData.Cells(j, tgtCols("RH日実績")).Value = srcData.Cells(i, srcCols("実績")).Value
                    End If
                    If tgtCols("RH日不良実績") > 0 Then
                        tgtData.Cells(j, tgtCols("RH日不良実績")).Value = srcData.Cells(i, srcCols("不良")).Value
                    End If
                    If tgtCols("RH日稼働時間") > 0 Then
                        tgtData.Cells(j, tgtCols("RH日稼働時間")).Value = srcData.Cells(i, srcCols("稼働時間")).Value
                    End If
                    transferred = transferred + 1
                    
                ElseIf hinban = "53828-60080" Then
                    ' LH品番の転記
                    If tgtCols("LH日実績") > 0 Then
                        tgtData.Cells(j, tgtCols("LH日実績")).Value = srcData.Cells(i, srcCols("実績")).Value
                    End If
                    If tgtCols("LH日不良実績") > 0 Then
                        tgtData.Cells(j, tgtCols("LH日不良実績")).Value = srcData.Cells(i, srcCols("不良")).Value
                    End If
                    If tgtCols("LH日稼働時間") > 0 Then
                        tgtData.Cells(j, tgtCols("LH日稼働時間")).Value = srcData.Cells(i, srcCols("稼働時間")).Value
                    End If
                    transferred = transferred + 1
                End If
                
                Exit For ' 日付が見つかったら次の行へ
            End If
        Next j
        
NextRow:
    Next i
    
    ' 合計値の転記
    Application.StatusBar = "合計値を転記中..."
    Dim totalTransferred As Long: totalTransferred = 0
    For j = 1 To tgtData.Rows.Count
        Dim tgtDate As Date
        tgtDate = tgtData.Cells(j, tgtCols("日付")).Value
        dateKey = Format(tgtDate, "yyyy-mm-dd")
        
        If dailyTotals.Exists(dateKey) Then
            totals = dailyTotals(dateKey)
            
            ' 合計値を転記
            If tgtCols("合計日実績") > 0 Then
                tgtData.Cells(j, tgtCols("合計日実績")).Value = totals(0)
            End If
            If tgtCols("合計日不良実績") > 0 Then
                tgtData.Cells(j, tgtCols("合計日不良実績")).Value = totals(1)
            End If
            If tgtCols("合計日稼働時間") > 0 Then
                tgtData.Cells(j, tgtCols("合計日稼働時間")).Value = totals(2)
            End If
            totalTransferred = totalTransferred + 1
        End If
    Next j
    
    ' 完了時のステータスバー表示
    Application.StatusBar = "TG品番別転記完了: " & transferred & "件の品番別データ、" & _
                           totalTransferred & "件の合計データを転記"
    
    ' 1秒待機してからクリア
    Application.Wait Now + TimeValue("0:00:01")
    Application.StatusBar = False
    
Cleanup:
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    ' エラー時だけメッセージボックス
    MsgBox "TG品番別転記処理でエラー発生" & vbCrLf & vbCrLf & _
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

