Attribute VB_Name = "m期間集計_通称別a_最適化版"
Sub 期間集計_通称別a_最適化版()
    ' 「品番別」シートの「_品番別」テーブルから直接通称別集計を行い、
    ' 「品番別aa」シートの「_品番別aa」テーブルに出力するマクロ
    ' 品番を通称に変換してグループ化し、各項目を集計する
    
    ' ★最適化：画面更新の停止（画面ちらつき防止の最重要設定）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' ステータスバーに処理状況を表示
    Application.StatusBar = "通称別直接集計: 処理を開始します..."
    
    ' 変数宣言
    Dim srcSheet As Worksheet, destSheet As Worksheet
    Dim srcTable As ListObject, destTable As ListObject
    Dim srcData As Variant
    Dim StartDate As Double, EndDate As Double  ' 日付をシリアル値として扱う
    Dim dictGroups As Object
    Dim dictSums As Object
    Dim dictCounts As Object
    Dim headerRow As Range
    Dim i As Long, j As Long
    Dim key As Variant
    Dim destRow As Long
    Dim tempValue As Variant
    Dim tsushoArr() As Variant
    Dim useFilter As Boolean
    Dim rowDateValue As Double
    Dim dataStartRow As Long
    Dim dataEndRow As Long
    Dim tableRange As Range
    Dim hinban As String
    Dim tsusho As String
    Dim isInDateRange As Boolean
    Dim tableFound As Boolean
    Dim dataRng As Range
    Dim lastRow As Long, lastCol As Long
    
    ' 列インデックス用変数
    Dim hinbanCol As Integer, dateCol As Integer
    Dim kataKaeCol As Integer, kadoCol As Integer, cycleCol As Integer
    Dim shotCol As Integer, furyoCol As Integer
    Dim uchidashiCol As Integer, shortCol As Integer, weldCol As Integer
    Dim shiwaCol As Integer, ibutsuCol As Integer, silverCol As Integer
    Dim flowCol As Integer, gomiCol As Integer, gcKasuCol As Integer
    Dim kizuCol As Integer, hikeCol As Integer, itohikiCol As Integer
    Dim kataYogoreCol As Integer, makureCol As Integer, toridashiFuryoCol As Integer
    Dim wareHakukaCol As Integer, coreKasuCol As Integer, sonotaCol As Integer
    Dim chocoCol As Integer, kensaCol As Integer, ryushutuCol As Integer
    
    ' ★最適化：エラーハンドリング設定（設定復旧を確実に行う）
    On Error GoTo ErrorHandler
    
    ' 入力元シートの取得
    On Error Resume Next
    Set srcSheet = ThisWorkbook.Worksheets("品番別")
    On Error GoTo ErrorHandler
    
    If srcSheet Is Nothing Then
        MsgBox "「品番別」シートが見つかりません。", vbCritical
        GoTo Cleanup
    End If
    
    ' 出力先シートの設定
    On Error Resume Next
    Set destSheet = ThisWorkbook.Worksheets("品番別aa")
    On Error GoTo ErrorHandler
    
    If destSheet Is Nothing Then
        ' シートが存在しない場合は新規作成
        Set destSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        destSheet.Name = "品番別aa"
    End If
    
    ' ステータスバーを更新
    Application.StatusBar = "通称別直接集計: 日付範囲を取得中..."
    
    ' 開始日と終了日を取得（品番別aaシートから）
    On Error Resume Next
    StartDate = CDbl(destSheet.Range("B1").Value)
    EndDate = CDbl(destSheet.Range("B2").Value)
    On Error GoTo ErrorHandler
    
    ' 日付が設定されているかチェック
    useFilter = (StartDate > 0) And (EndDate > 0)
    
    ' テーブルの検索
    tableFound = False
    On Error Resume Next
    Set srcTable = srcSheet.ListObjects("_品番別")
    On Error GoTo ErrorHandler
    
    If Not srcTable Is Nothing Then
        tableFound = True
    End If
    
    ' テーブルが見つからない場合は、データ範囲を探して変換
    If Not tableFound Then
        ' データ範囲を探す
        If Not IsEmpty(srcSheet.Range("A1").Value) Then
            lastRow = srcSheet.Cells(srcSheet.Rows.Count, 1).End(xlUp).Row
            lastCol = srcSheet.Cells(1, srcSheet.Columns.Count).End(xlToLeft).Column
            
            If lastRow > 1 Then  ' ヘッダー行とデータが少なくとも1行ある
                Set dataRng = srcSheet.Range(srcSheet.Cells(1, 1), srcSheet.Cells(lastRow, lastCol))
                
                ' データ範囲をテーブルに変換
                On Error Resume Next
                Set srcTable = srcSheet.ListObjects.Add(xlSrcRange, dataRng, , xlYes)
                If Err.Number = 0 Then
                    srcTable.Name = "_品番別"
                    tableFound = True
                End If
                On Error GoTo ErrorHandler
            End If
        End If
    End If
    
    ' それでもテーブルが見つからない場合は処理中止
    If Not tableFound Then
        MsgBox "テーブル「_品番別」が見つかりません。", vbCritical
        GoTo Cleanup
    End If
    
    ' ステータスバーを更新
    Application.StatusBar = "通称別直接集計: データ取得中..."
    
    ' ★最適化：元データを配列として一括取得
    srcData = srcTable.DataBodyRange.Value
    
    ' 列のインデックスを取得
    Set headerRow = srcTable.HeaderRowRange
    
    For i = 1 To headerRow.Cells.Count
        Select Case headerRow.Cells(1, i).Value
            Case "品番"
                hinbanCol = i
            Case "日付"
                dateCol = i
            Case "型替"
                kataKaeCol = i
            Case "稼動"
                kadoCol = i
            Case "サイクル"
                cycleCol = i
            Case "ショット数"
                shotCol = i
            Case "不良数"
                furyoCol = i
            Case "打出調"
                uchidashiCol = i
            Case "ショート"
                shortCol = i
            Case "ウエルド"
                weldCol = i
            Case "シワ"
                shiwaCol = i
            Case "異物"
                ibutsuCol = i
            Case "シルバー"
                silverCol = i
            Case "フローマーク"
                flowCol = i
            Case "ゴミ付着"
                gomiCol = i
            Case "GCカス"
                gcKasuCol = i
            Case "キズ"
                kizuCol = i
            Case "ヒケ"
                hikeCol = i
            Case "糸引き"
                itohikiCol = i
            Case "型汚れ"
                kataYogoreCol = i
            Case "マクレ"
                makureCol = i
            Case "取出不良"
                toridashiFuryoCol = i
            Case "割れ剥化"
                wareHakukaCol = i
            Case "コアカス"
                coreKasuCol = i
            Case "その他"
                sonotaCol = i
            Case "チョコ停打出調"
                chocoCol = i
            Case "検査"
                kensaCol = i
            Case "流出不良"
                ryushutuCol = i
        End Select
    Next i
    
    ' 必要な列が見つからない場合は処理中止
    If hinbanCol = 0 Or dateCol = 0 Then
        MsgBox "必要な列が見つかりません。", vbCritical
        GoTo Cleanup
    End If
    
    ' ステータスバーを更新
    Application.StatusBar = "通称別直接集計: データ集計中..."
    
    ' Dictionaryオブジェクトを作成
    Set dictGroups = CreateObject("Scripting.Dictionary")
    Set dictSums = CreateObject("Scripting.Dictionary")
    Set dictCounts = CreateObject("Scripting.Dictionary")
    
    ' データをグループ化して集計
    For i = 1 To UBound(srcData, 1)
        ' ★最適化：100行ごとに進捗更新（頻繁な更新は避ける）
        If i Mod 100 = 0 Then
            Application.StatusBar = "通称別直接集計: データ集計中... " & Format(i / UBound(srcData, 1), "0%")
        End If
        
        ' 日付の確認
        isInDateRange = True
        
        If useFilter Then
            ' 日付のシリアル値を取得
            rowDateValue = CDbl(srcData(i, dateCol))
            
            ' 日付範囲内かどうかチェック
            isInDateRange = (rowDateValue >= StartDate And rowDateValue <= EndDate)
        End If
        
        ' 日付範囲内のデータのみ処理
        If isInDateRange Then
            ' 品番から通称を判定
            hinban = srcData(i, hinbanCol)
            tsusho = 品番通称判定(hinban)
            
            ' 新しい通称の場合、Dictionaryに追加
            If Not dictGroups.Exists(tsusho) Then
                dictGroups.Add tsusho, tsusho
                
                ' 集計用のDictionaryを初期化
                Set dictSums(tsusho) = CreateObject("Scripting.Dictionary")
                Set dictCounts(tsusho) = CreateObject("Scripting.Dictionary")
                
                ' 各集計項目を初期化
                dictSums(tsusho)("型替") = 0
                dictSums(tsusho)("稼動") = 0
                dictSums(tsusho)("サイクル") = 0
                dictCounts(tsusho)("サイクル") = 0  ' サイクル平均計算用
                dictSums(tsusho)("ショット数") = 0
                dictSums(tsusho)("不良数") = 0
                dictSums(tsusho)("打出調") = 0
                dictSums(tsusho)("ショート") = 0
                dictSums(tsusho)("ウエルド") = 0
                dictSums(tsusho)("シワ") = 0
                dictSums(tsusho)("異物") = 0
                dictSums(tsusho)("シルバー") = 0
                dictSums(tsusho)("フローマーク") = 0
                dictSums(tsusho)("ゴミ付着") = 0
                dictSums(tsusho)("GCカス") = 0
                dictSums(tsusho)("キズ") = 0
                dictSums(tsusho)("ヒケ") = 0
                dictSums(tsusho)("糸引き") = 0
                dictSums(tsusho)("型汚れ") = 0
                dictSums(tsusho)("マクレ") = 0
                dictSums(tsusho)("取出不良") = 0
                dictSums(tsusho)("割れ剥化") = 0
                dictSums(tsusho)("コアカス") = 0
                dictSums(tsusho)("その他") = 0
                dictSums(tsusho)("チョコ停打出調") = 0
                dictSums(tsusho)("検査") = 0
                dictSums(tsusho)("流出不良") = 0
            End If
            
            ' 各項目の合計値を更新
            If kataKaeCol > 0 And IsNumeric(srcData(i, kataKaeCol)) Then
                dictSums(tsusho)("型替") = dictSums(tsusho)("型替") + CDbl(srcData(i, kataKaeCol))
            End If
            
            If kadoCol > 0 And IsNumeric(srcData(i, kadoCol)) Then
                dictSums(tsusho)("稼動") = dictSums(tsusho)("稼動") + CDbl(srcData(i, kadoCol))
            End If
            
            If cycleCol > 0 And IsNumeric(srcData(i, cycleCol)) And srcData(i, cycleCol) <> 0 Then
                dictSums(tsusho)("サイクル") = dictSums(tsusho)("サイクル") + CDbl(srcData(i, cycleCol))
                dictCounts(tsusho)("サイクル") = dictCounts(tsusho)("サイクル") + 1
            End If
            
            If shotCol > 0 And IsNumeric(srcData(i, shotCol)) Then
                dictSums(tsusho)("ショット数") = dictSums(tsusho)("ショット数") + CDbl(srcData(i, shotCol))
            End If
            
            If furyoCol > 0 And IsNumeric(srcData(i, furyoCol)) Then
                dictSums(tsusho)("不良数") = dictSums(tsusho)("不良数") + CDbl(srcData(i, furyoCol))
            End If
            
            ' 不良項目の集計
            If uchidashiCol > 0 And IsNumeric(srcData(i, uchidashiCol)) Then
                dictSums(tsusho)("打出調") = dictSums(tsusho)("打出調") + CDbl(srcData(i, uchidashiCol))
            End If
            
            If shortCol > 0 And IsNumeric(srcData(i, shortCol)) Then
                dictSums(tsusho)("ショート") = dictSums(tsusho)("ショート") + CDbl(srcData(i, shortCol))
            End If
            
            If weldCol > 0 And IsNumeric(srcData(i, weldCol)) Then
                dictSums(tsusho)("ウエルド") = dictSums(tsusho)("ウエルド") + CDbl(srcData(i, weldCol))
            End If
            
            If shiwaCol > 0 And IsNumeric(srcData(i, shiwaCol)) Then
                dictSums(tsusho)("シワ") = dictSums(tsusho)("シワ") + CDbl(srcData(i, shiwaCol))
            End If
            
            If ibutsuCol > 0 And IsNumeric(srcData(i, ibutsuCol)) Then
                dictSums(tsusho)("異物") = dictSums(tsusho)("異物") + CDbl(srcData(i, ibutsuCol))
            End If
            
            If silverCol > 0 And IsNumeric(srcData(i, silverCol)) Then
                dictSums(tsusho)("シルバー") = dictSums(tsusho)("シルバー") + CDbl(srcData(i, silverCol))
            End If
            
            If flowCol > 0 And IsNumeric(srcData(i, flowCol)) Then
                dictSums(tsusho)("フローマーク") = dictSums(tsusho)("フローマーク") + CDbl(srcData(i, flowCol))
            End If
            
            If gomiCol > 0 And IsNumeric(srcData(i, gomiCol)) Then
                dictSums(tsusho)("ゴミ付着") = dictSums(tsusho)("ゴミ付着") + CDbl(srcData(i, gomiCol))
            End If
            
            If gcKasuCol > 0 And IsNumeric(srcData(i, gcKasuCol)) Then
                dictSums(tsusho)("GCカス") = dictSums(tsusho)("GCカス") + CDbl(srcData(i, gcKasuCol))
            End If
            
            If kizuCol > 0 And IsNumeric(srcData(i, kizuCol)) Then
                dictSums(tsusho)("キズ") = dictSums(tsusho)("キズ") + CDbl(srcData(i, kizuCol))
            End If
            
            If hikeCol > 0 And IsNumeric(srcData(i, hikeCol)) Then
                dictSums(tsusho)("ヒケ") = dictSums(tsusho)("ヒケ") + CDbl(srcData(i, hikeCol))
            End If
            
            If itohikiCol > 0 And IsNumeric(srcData(i, itohikiCol)) Then
                dictSums(tsusho)("糸引き") = dictSums(tsusho)("糸引き") + CDbl(srcData(i, itohikiCol))
            End If
            
            If kataYogoreCol > 0 And IsNumeric(srcData(i, kataYogoreCol)) Then
                dictSums(tsusho)("型汚れ") = dictSums(tsusho)("型汚れ") + CDbl(srcData(i, kataYogoreCol))
            End If
            
            If makureCol > 0 And IsNumeric(srcData(i, makureCol)) Then
                dictSums(tsusho)("マクレ") = dictSums(tsusho)("マクレ") + CDbl(srcData(i, makureCol))
            End If
            
            If toridashiFuryoCol > 0 And IsNumeric(srcData(i, toridashiFuryoCol)) Then
                dictSums(tsusho)("取出不良") = dictSums(tsusho)("取出不良") + CDbl(srcData(i, toridashiFuryoCol))
            End If
            
            If wareHakukaCol > 0 And IsNumeric(srcData(i, wareHakukaCol)) Then
                dictSums(tsusho)("割れ剥化") = dictSums(tsusho)("割れ剥化") + CDbl(srcData(i, wareHakukaCol))
            End If
            
            If coreKasuCol > 0 And IsNumeric(srcData(i, coreKasuCol)) Then
                dictSums(tsusho)("コアカス") = dictSums(tsusho)("コアカス") + CDbl(srcData(i, coreKasuCol))
            End If
            
            If sonotaCol > 0 And IsNumeric(srcData(i, sonotaCol)) Then
                dictSums(tsusho)("その他") = dictSums(tsusho)("その他") + CDbl(srcData(i, sonotaCol))
            End If
            
            If chocoCol > 0 And IsNumeric(srcData(i, chocoCol)) Then
                dictSums(tsusho)("チョコ停打出調") = dictSums(tsusho)("チョコ停打出調") + CDbl(srcData(i, chocoCol))
            End If
            
            If kensaCol > 0 And IsNumeric(srcData(i, kensaCol)) Then
                dictSums(tsusho)("検査") = dictSums(tsusho)("検査") + CDbl(srcData(i, kensaCol))
            End If
            
            If ryushutuCol > 0 And IsNumeric(srcData(i, ryushutuCol)) Then
                dictSums(tsusho)("流出不良") = dictSums(tsusho)("流出不良") + CDbl(srcData(i, ryushutuCol))
            End If
        End If
    Next i
    
    ' ステータスバーを更新
    Application.StatusBar = "通称別直接集計: データ出力準備中..."
    
    ' ★最適化：出力先シートの操作をWithブロックで効率化
    With destSheet
        ' 出力先シートを4行目以降をクリア（1-3行目は残す）
        .Range("A4:AB31").Clear
        
        ' 4行目以降の書式設定
        With .Range("A4:AB" & .Rows.Count)
            .Font.Name = "Yu Gothic UI"
            .Font.Size = 11
        End With
        
        ' タイトル行の作成（4行目）
        destRow = 4
        If useFilter Then
            .Range("A" & destRow).Value = "期間別通称別不良集計：" & Format(StartDate, "yyyy/mm/dd") & "～" & Format(EndDate, "yyyy/mm/dd")
        Else
            .Range("A" & destRow).Value = "期間別通称別不良集計：全期間"
        End If
        .Range("A" & destRow).Font.Bold = True
        
        ' ヘッダー行の作成（5行目）
        destRow = 5
        .Range("A" & destRow).Value = "通称"
        .Range("B" & destRow).Value = "型替"
        .Range("C" & destRow).Value = "稼動"
        .Range("D" & destRow).Value = "サイクル"
        .Range("E" & destRow).Value = "ショット数"
        .Range("F" & destRow).Value = "不良数"
        .Range("G" & destRow).Value = "不良率"
        .Range("H" & destRow).Value = "打出調"
        .Range("I" & destRow).Value = "ショート"
        .Range("J" & destRow).Value = "ウエルド"
        .Range("K" & destRow).Value = "シワ"
        .Range("L" & destRow).Value = "異物"
        .Range("M" & destRow).Value = "シルバー"
        .Range("N" & destRow).Value = "フローマーク"
        .Range("O" & destRow).Value = "ゴミ付着"
        .Range("P" & destRow).Value = "GCカス"
        .Range("Q" & destRow).Value = "キズ"
        .Range("R" & destRow).Value = "ヒケ"
        .Range("S" & destRow).Value = "糸引き"
        .Range("T" & destRow).Value = "型汚れ"
        .Range("U" & destRow).Value = "マクレ"
        .Range("V" & destRow).Value = "取出不良"
        .Range("W" & destRow).Value = "割れ剥化"
        .Range("X" & destRow).Value = "コアカス"
        .Range("Y" & destRow).Value = "その他"
        .Range("Z" & destRow).Value = "チョコ停打出調"
        .Range("AA" & destRow).Value = "検査"
        .Range("AB" & destRow).Value = "流出不良"
        
        ' ヘッダー行の書式設定
        With .Range("A" & destRow & ":AB" & destRow)
            .HorizontalAlignment = xlCenter  ' 中央揃え
            .Font.Bold = True
            .ShrinkToFit = True  ' 縮小して全体を表示
        End With
    End With
    
    destRow = destRow + 1
    dataStartRow = destRow
    
    ' データがない場合の処理
    If dictGroups.Count = 0 Then
        destSheet.Cells(dataStartRow, 1).Value = "該当データなし"
        For j = 2 To 28  ' B列からAB列まで0で埋める
            destSheet.Cells(dataStartRow, j).Value = 0
        Next j
        
        dataEndRow = dataStartRow
    Else
        ' 通称の配列を作成してソート
        ReDim tsushoArr(0 To dictGroups.Count - 1)
        i = 0
        For Each key In dictGroups.Keys
            tsushoArr(i) = key
            i = i + 1
        Next key
        
        ' 配列をソート（指定の順番で）
        ' TG → 62-28030Fr → 62-28030Rr → 62-58050Fr → 62-58050Rr → 補給品
        Dim sortedArr() As Variant
        ReDim sortedArr(0 To UBound(tsushoArr))
        Dim sortIdx As Integer
        sortIdx = 0
        
        ' 順番に配列作成
        Dim orderList As Variant
        orderList = Array("TG", "62-28030Fr", "62-28030Rr", "62-58050Fr", "62-58050Rr", "補給品")
        
        For i = 0 To UBound(orderList)
            For j = 0 To UBound(tsushoArr)
                If tsushoArr(j) = orderList(i) Then
                    sortedArr(sortIdx) = tsushoArr(j)
                    sortIdx = sortIdx + 1
                    Exit For
                End If
            Next j
        Next i
        
        ' ステータスバーを更新
        Application.StatusBar = "通称別直接集計: データ出力中..."
        
        ' ★最適化：出力データを配列に準備してから一括書き込み
        Dim outputData() As Variant
        ReDim outputData(1 To sortIdx, 1 To 28)
        
        ' データの書き込み準備
        For i = 0 To sortIdx - 1
            key = sortedArr(i)
            
            ' 基本データを配列に格納
            outputData(i + 1, 1) = key
            outputData(i + 1, 2) = dictSums(key)("型替")
            outputData(i + 1, 3) = dictSums(key)("稼動")
            
            ' サイクルの平均値を計算
            If dictCounts(key)("サイクル") > 0 Then
                outputData(i + 1, 4) = dictSums(key)("サイクル") / dictCounts(key)("サイクル")
            Else
                outputData(i + 1, 4) = 0
            End If
            
            outputData(i + 1, 5) = dictSums(key)("ショット数")
            outputData(i + 1, 6) = dictSums(key)("不良数")
            
            ' 不良率の計算（不良数÷ショット数）
            If dictSums(key)("ショット数") > 0 Then
                outputData(i + 1, 7) = dictSums(key)("不良数") / dictSums(key)("ショット数")
            Else
                outputData(i + 1, 7) = 0
            End If
            
            ' 不良項目データを配列に格納
            outputData(i + 1, 8) = dictSums(key)("打出調")
            outputData(i + 1, 9) = dictSums(key)("ショート")
            outputData(i + 1, 10) = dictSums(key)("ウエルド")
            outputData(i + 1, 11) = dictSums(key)("シワ")
            outputData(i + 1, 12) = dictSums(key)("異物")
            outputData(i + 1, 13) = dictSums(key)("シルバー")
            outputData(i + 1, 14) = dictSums(key)("フローマーク")
            outputData(i + 1, 15) = dictSums(key)("ゴミ付着")
            outputData(i + 1, 16) = dictSums(key)("GCカス")
            outputData(i + 1, 17) = dictSums(key)("キズ")
            outputData(i + 1, 18) = dictSums(key)("ヒケ")
            outputData(i + 1, 19) = dictSums(key)("糸引き")
            outputData(i + 1, 20) = dictSums(key)("型汚れ")
            outputData(i + 1, 21) = dictSums(key)("マクレ")
            outputData(i + 1, 22) = dictSums(key)("取出不良")
            outputData(i + 1, 23) = dictSums(key)("割れ剥化")
            outputData(i + 1, 24) = dictSums(key)("コアカス")
            outputData(i + 1, 25) = dictSums(key)("その他")
            outputData(i + 1, 26) = dictSums(key)("チョコ停打出調")
            outputData(i + 1, 27) = dictSums(key)("検査")
            outputData(i + 1, 28) = dictSums(key)("流出不良")
        Next i
        
        ' ★最適化：配列データを一括で書き込み
        destSheet.Range("A" & dataStartRow).Resize(sortIdx, 28).Value = outputData
        
        dataEndRow = dataStartRow + sortIdx - 1
    End If
    
    ' テーブルの作成
    Set tableRange = destSheet.Range("A5").Resize(dataEndRow - 4, 28)
    
    ' 既に同名のテーブルが存在する場合は削除
    On Error Resume Next
    If Not destSheet.ListObjects("_品番別aa") Is Nothing Then
        destSheet.ListObjects("_品番別aa").Delete
    End If
    On Error GoTo ErrorHandler
    
    Set destTable = destSheet.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    destTable.Name = "_品番別aa"
    destTable.ShowAutoFilter = False  ' フィルターボタン非表示
    
    ' ステータスバーを更新
    Application.StatusBar = "通称別直接集計: 書式設定中..."
    
    ' テーブル内の書式設定
    With destSheet.Range("A" & dataStartRow & ":AB" & dataEndRow)
        .ShrinkToFit = True  ' 縮小して全体を表示
    End With
    
    ' 不良率のフォーマット設定（%表示、小数点以下2桁）
    destSheet.Range("G" & dataStartRow & ":G" & dataEndRow).NumberFormat = "0.00%"
    
    ' サイクルのフォーマット設定（小数点以下1桁）
    destSheet.Range("D" & dataStartRow & ":D" & dataEndRow).NumberFormat = "0.0"
    
    ' 列幅の設定
    destSheet.Columns("A").ColumnWidth = 14  ' 通称列
    destSheet.Columns("B:G").ColumnWidth = 7  ' 型替～不良率列
    destSheet.Columns("H:AB").ColumnWidth = 3  ' 不良項目列
    
    ' 0の値を薄いグレーにする条件付き書式
    With destSheet.Range("B" & dataStartRow & ":AB" & dataEndRow)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="0"
        .FormatConditions(1).Font.Color = RGB(192, 192, 192)  ' 薄いグレー
    End With
    
    ' 処理完了
    Application.StatusBar = "通称別直接集計: 処理が完了しました。"
    
    ' 1秒待機してステータスバークリア
    Application.Wait Now + TimeValue("00:00:01")
    
Cleanup:
    ' ★最適化：設定を必ず元に戻す（エラー時も含む）
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
    Resume Cleanup
End Sub

Private Function 品番通称判定(品番 As String) As String
    ' 品番から通称を判定する関数
    ' 品番の文字列パターンから適切な通称を返す
    
    Select Case True
        Case InStr(品番, "TG") > 0
            品番通称判定 = "TG"
        Case InStr(品番, "62-28030") > 0 And InStr(品番, "Fr") > 0
            品番通称判定 = "62-28030Fr"
        Case InStr(品番, "62-28030") > 0 And InStr(品番, "Rr") > 0
            品番通称判定 = "62-28030Rr"
        Case InStr(品番, "62-58050") > 0 And InStr(品番, "Fr") > 0
            品番通称判定 = "62-58050Fr"
        Case InStr(品番, "62-58050") > 0 And InStr(品番, "Rr") > 0
            品番通称判定 = "62-58050Rr"
        Case Else
            品番通称判定 = "補給品"
    End Select
End Function