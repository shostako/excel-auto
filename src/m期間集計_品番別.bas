Attribute VB_Name = "m期間集計_品番別"
Option Explicit

Sub 期間集計_品番別()
    Application.StatusBar = "品番別期間集計を開始します..."
    
    ' 高速化の三種の神器
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler
    
    ' ワークシート参照
    Dim wsSrc As Worksheet, wsTgt As Worksheet
    Set wsSrc = ThisWorkbook.Worksheets("品番別")
    Set wsTgt = ThisWorkbook.Worksheets("品番別a")
    
    ' テーブル参照
    Dim tblSrc As ListObject, tblTgt As ListObject
    Set tblSrc = wsSrc.ListObjects("_品番別")
    Set tblTgt = wsTgt.ListObjects("_品番別a")
    
    ' 期間取得
    Dim startDate As Date, endDate As Date
    startDate = wsTgt.Range("B1").Value
    endDate = wsTgt.Range("B2").Value
    
    If startDate = 0 Or endDate = 0 Then
        MsgBox "開始日（B1）または終了日（B2）が設定されていません。", vbCritical
        GoTo Cleanup
    End If
    
    Application.StatusBar = "データを読み込み中..."
    
    ' ソースデータ確認
    If tblSrc.DataBodyRange Is Nothing Then
        MsgBox "ソーステーブル「_品番別」にデータがありません。", vbCritical
        GoTo Cleanup
    End If
    
    ' 列インデックス取得（ソース）
    Dim srcCols As Object
    Set srcCols = CreateObject("Scripting.Dictionary")
    srcCols("日付") = tblSrc.ListColumns("日付").Index
    srcCols("品番") = tblSrc.ListColumns("品番").Index
    srcCols("型替") = tblSrc.ListColumns("型替").Index
    srcCols("稼働") = tblSrc.ListColumns("稼働").Index
    srcCols("サイクル") = tblSrc.ListColumns("サイクル").Index
    srcCols("ショット数") = tblSrc.ListColumns("ショット数").Index
    srcCols("不良数") = tblSrc.ListColumns("不良数").Index
    srcCols("打出し") = tblSrc.ListColumns("打出し").Index
    srcCols("ショート") = tblSrc.ListColumns("ショート").Index
    srcCols("ウエルド") = tblSrc.ListColumns("ウエルド").Index
    srcCols("シワ") = tblSrc.ListColumns("シワ").Index
    srcCols("異物") = tblSrc.ListColumns("異物").Index
    srcCols("シルバー") = tblSrc.ListColumns("シルバー").Index
    srcCols("フローマーク") = tblSrc.ListColumns("フローマーク").Index
    srcCols("ゴミ押し") = tblSrc.ListColumns("ゴミ押し").Index
    srcCols("GCカス") = tblSrc.ListColumns("GCカス").Index
    srcCols("キズ") = tblSrc.ListColumns("キズ").Index
    srcCols("ヒケ") = tblSrc.ListColumns("ヒケ").Index
    srcCols("糸引き") = tblSrc.ListColumns("糸引き").Index
    srcCols("型汚れ") = tblSrc.ListColumns("型汚れ").Index
    srcCols("マクレ") = tblSrc.ListColumns("マクレ").Index
    srcCols("取出不良") = tblSrc.ListColumns("取出不良").Index
    srcCols("割れ白化") = tblSrc.ListColumns("割れ白化").Index
    srcCols("コアカス") = tblSrc.ListColumns("コアカス").Index
    srcCols("その他") = tblSrc.ListColumns("その他").Index
    srcCols("チョコ停打出し") = tblSrc.ListColumns("チョコ停打出し").Index
    srcCols("検査") = tblSrc.ListColumns("検査").Index
    srcCols("流出不良") = tblSrc.ListColumns("流出不良").Index
    
    ' データ配列取得
    Dim srcData As Variant
    srcData = tblSrc.DataBodyRange.Value
    
    Application.StatusBar = "データを集計中..."
    
    ' 品番別集計用Dictionary
    Dim summaryDict As Object
    Set summaryDict = CreateObject("Scripting.Dictionary")
    
    ' データ集計処理
    Dim i As Long, hinban As String, targetDate As Date
    For i = 1 To UBound(srcData, 1)
        ' 日付チェック（型変換エラー対策）
        On Error Resume Next
        targetDate = srcData(i, srcCols("日付"))
        On Error GoTo ErrorHandler
        
        If targetDate = 0 Then
            Debug.Print "日付変換エラー: " & i & "行目 値=" & srcData(i, srcCols("日付"))
            GoTo NextRow
        End If
        
        ' 期間チェック
        If targetDate >= startDate And targetDate <= endDate Then
            hinban = CStr(srcData(i, srcCols("品番")))
            
            ' 初回なら初期化
            If Not summaryDict.Exists(hinban) Then
                Dim newRow As Object
                Set newRow = CreateObject("Scripting.Dictionary")
                newRow("品番") = hinban
                newRow("型替") = 0
                newRow("稼働") = 0
                newRow("サイクル") = 0
                newRow("ショット数") = 0
                newRow("不良数") = 0
                newRow("打出し") = 0
                newRow("ショート") = 0
                newRow("ウエルド") = 0
                newRow("シワ") = 0
                newRow("異物") = 0
                newRow("シルバー") = 0
                newRow("フローマーク") = 0
                newRow("ゴミ押し") = 0
                newRow("GCカス") = 0
                newRow("キズ") = 0
                newRow("ヒケ") = 0
                newRow("糸引き") = 0
                newRow("型汚れ") = 0
                newRow("マクレ") = 0
                newRow("取出不良") = 0
                newRow("割れ白化") = 0
                newRow("コアカス") = 0
                newRow("その他") = 0
                newRow("チョコ停打出し") = 0
                newRow("検査") = 0
                newRow("流出不良") = 0
                Set summaryDict(hinban) = newRow
            End If
            
            ' 集計値加算
            Dim currentRow As Object
            Set currentRow = summaryDict(hinban)
            currentRow("型替") = currentRow("型替") + Val(srcData(i, srcCols("型替")))
            currentRow("稼働") = currentRow("稼働") + Val(srcData(i, srcCols("稼働")))
            currentRow("サイクル") = currentRow("サイクル") + Val(srcData(i, srcCols("サイクル")))
            currentRow("ショット数") = currentRow("ショット数") + Val(srcData(i, srcCols("ショット数")))
            currentRow("不良数") = currentRow("不良数") + Val(srcData(i, srcCols("不良数")))
            currentRow("打出し") = currentRow("打出し") + Val(srcData(i, srcCols("打出し")))
            currentRow("ショート") = currentRow("ショート") + Val(srcData(i, srcCols("ショート")))
            currentRow("ウエルド") = currentRow("ウエルド") + Val(srcData(i, srcCols("ウエルド")))
            currentRow("シワ") = currentRow("シワ") + Val(srcData(i, srcCols("シワ")))
            currentRow("異物") = currentRow("異物") + Val(srcData(i, srcCols("異物")))
            currentRow("シルバー") = currentRow("シルバー") + Val(srcData(i, srcCols("シルバー")))
            currentRow("フローマーク") = currentRow("フローマーク") + Val(srcData(i, srcCols("フローマーク")))
            currentRow("ゴミ押し") = currentRow("ゴミ押し") + Val(srcData(i, srcCols("ゴミ押し")))
            currentRow("GCカス") = currentRow("GCカス") + Val(srcData(i, srcCols("GCカス")))
            currentRow("キズ") = currentRow("キズ") + Val(srcData(i, srcCols("キズ")))
            currentRow("ヒケ") = currentRow("ヒケ") + Val(srcData(i, srcCols("ヒケ")))
            currentRow("糸引き") = currentRow("糸引き") + Val(srcData(i, srcCols("糸引き")))
            currentRow("型汚れ") = currentRow("型汚れ") + Val(srcData(i, srcCols("型汚れ")))
            currentRow("マクレ") = currentRow("マクレ") + Val(srcData(i, srcCols("マクレ")))
            currentRow("取出不良") = currentRow("取出不良") + Val(srcData(i, srcCols("取出不良")))
            currentRow("割れ白化") = currentRow("割れ白化") + Val(srcData(i, srcCols("割れ白化")))
            currentRow("コアカス") = currentRow("コアカス") + Val(srcData(i, srcCols("コアカス")))
            currentRow("その他") = currentRow("その他") + Val(srcData(i, srcCols("その他")))
            currentRow("チョコ停打出し") = currentRow("チョコ停打出し") + Val(srcData(i, srcCols("チョコ停打出し")))
            currentRow("検査") = currentRow("検査") + Val(srcData(i, srcCols("検査")))
            currentRow("流出不良") = currentRow("流出不良") + Val(srcData(i, srcCols("流出不良")))
        End If
        
NextRow:
        ' 進捗表示（100件ごと）
        If i Mod 100 = 0 Then
            Application.StatusBar = "集計中... " & Format(i / UBound(srcData, 1), "0%") & " (" & i & "/" & UBound(srcData, 1) & ")"
        End If
    Next i
    
    Application.StatusBar = "出力データを準備中..."
    
    ' 出力用配列準備
    Dim outputCount As Long
    outputCount = summaryDict.Count
    
    If outputCount = 0 Then
        MsgBox "指定期間内にデータがありません。" & vbCrLf & _
               "期間: " & Format(startDate, "yyyy/mm/dd") & " ～ " & Format(endDate, "yyyy/mm/dd"), vbInformation
        GoTo Cleanup
    End If
    
    ' 出力配列作成（不良率列を追加）
    Dim outputData() As Variant
    ReDim outputData(1 To outputCount, 1 To 28) ' 27列 + 不良率
    
    ' 列名配列
    Dim colNames As Variant
    colNames = Array("品番", "型替", "稼働", "サイクル", "ショット数", "不良数", "不良率", _
                     "打出し", "ショート", "ウエルド", "シワ", "異物", "シルバー", "フローマーク", _
                     "ゴミ押し", "GCカス", "キズ", "ヒケ", "糸引き", "型汚れ", "マクレ", "取出不良", _
                     "割れ白化", "コアカス", "その他", "チョコ停打出し", "検査", "流出不良")
    
    ' データを配列に転記
    Dim rowIndex As Long
    rowIndex = 1
    Dim key As Variant
    For Each key In summaryDict.Keys
        Set currentRow = summaryDict(key)
        outputData(rowIndex, 1) = currentRow("品番")
        outputData(rowIndex, 2) = currentRow("型替")
        outputData(rowIndex, 3) = currentRow("稼働")
        outputData(rowIndex, 4) = currentRow("サイクル")
        outputData(rowIndex, 5) = currentRow("ショット数")
        outputData(rowIndex, 6) = currentRow("不良数")
        
        ' 不良率計算（ゼロ除算対策）
        If currentRow("ショット数") > 0 Then
            outputData(rowIndex, 7) = currentRow("不良数") / currentRow("ショット数")
        Else
            outputData(rowIndex, 7) = 0
        End If
        
        outputData(rowIndex, 8) = currentRow("打出し")
        outputData(rowIndex, 9) = currentRow("ショート")
        outputData(rowIndex, 10) = currentRow("ウエルド")
        outputData(rowIndex, 11) = currentRow("シワ")
        outputData(rowIndex, 12) = currentRow("異物")
        outputData(rowIndex, 13) = currentRow("シルバー")
        outputData(rowIndex, 14) = currentRow("フローマーク")
        outputData(rowIndex, 15) = currentRow("ゴミ押し")
        outputData(rowIndex, 16) = currentRow("GCカス")
        outputData(rowIndex, 17) = currentRow("キズ")
        outputData(rowIndex, 18) = currentRow("ヒケ")
        outputData(rowIndex, 19) = currentRow("糸引き")
        outputData(rowIndex, 20) = currentRow("型汚れ")
        outputData(rowIndex, 21) = currentRow("マクレ")
        outputData(rowIndex, 22) = currentRow("取出不良")
        outputData(rowIndex, 23) = currentRow("割れ白化")
        outputData(rowIndex, 24) = currentRow("コアカス")
        outputData(rowIndex, 25) = currentRow("その他")
        outputData(rowIndex, 26) = currentRow("チョコ停打出し")
        outputData(rowIndex, 27) = currentRow("検査")
        outputData(rowIndex, 28) = currentRow("流出不良")
        
        rowIndex = rowIndex + 1
    Next key
    
    Application.StatusBar = "データを出力中..."
    
    ' 既存テーブルをクリア
    If Not tblTgt.DataBodyRange Is Nothing Then
        tblTgt.DataBodyRange.Delete
    End If
    
    ' タイトル設定
    wsTgt.Range("A4").Value = "成形品番別不良集計：" & Format(startDate, "yyyy/mm/dd") & "～" & Format(endDate, "yyyy/mm/dd")
    
    ' ヘッダー設定
    Dim headerRange As Range
    Set headerRange = wsTgt.Range("A5").Resize(1, UBound(colNames) + 1)
    headerRange.Value = colNames
    
    ' データ出力
    Dim dataRange As Range
    Set dataRange = wsTgt.Range("A6").Resize(outputCount, 28)
    dataRange.Value = outputData
    
    ' テーブル範囲の更新
    tblTgt.Resize wsTgt.Range("A5").Resize(outputCount + 1, 28)
    
    Application.StatusBar = "書式設定中..."
    
    ' 書式設定
    With wsTgt.Cells
        .Font.Name = "Yu Gothic UI"
        .Font.Size = 11
    End With
    
    ' テーブル書式
    With tblTgt.Range
        .ShrinkToFit = True ' 縮小して全体を表示
    End With
    
    ' ヘッダー書式
    With tblTgt.HeaderRowRange
        .HorizontalAlignment = xlCenter ' 中央揃え
    End With
    
    ' 不良率列のみ%書式
    Dim defectRateCol As Range
    Set defectRateCol = tblTgt.ListColumns("不良率").DataBodyRange
    defectRateCol.NumberFormat = "0.00%"
    
    Application.StatusBar = "品番別期間集計が完了しました（" & outputCount & "件）"
    
    GoTo Cleanup

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
    
Cleanup:
    ' 設定を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
End Sub