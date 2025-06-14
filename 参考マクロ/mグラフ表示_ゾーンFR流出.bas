Attribute VB_Name = "m期間集計_通称別b"
Sub 期間集計_通称別b()
    ' 「品番別」シートの「_品番別」テーブルから直接通称別集計を行い、
    ' 「品番別bb」シートの「_品番別bb」テーブルに出力するマクロ
    ' 品番を通称に変換してからグループ化し、不良項目は率として計算する
    
    ' ステータスバーに処理状況を表示
    Application.StatusBar = "通称別直接集計b: 処理を開始します..."
    
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
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' 入力元シートの取得
    On Error Resume Next
    Set srcSheet = ThisWorkbook.Worksheets("品番別")
    On Error GoTo 0
    
    If srcSheet Is Nothing Then
        Application.StatusBar = "通称別直接集計b: 「品番別」シートが見つかりません。"
        Exit Sub
    End If
    
    ' 出力先シートの設定
    On Error Resume Next
    Set destSheet = ThisWorkbook.Worksheets("品番別bb")
    On Error GoTo 0
    
    If destSheet Is Nothing Then
        ' シートが存在しない場合は新規作成
        Set destSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        destSheet.Name = "品番別bb"
    End If
    
    ' ステータスバーを更新
    Application.StatusBar = "通称別直接集計b: 日付条件を取得中..."
    
    ' 開始日と終了日を取得（品番別bbシートから）
    On Error Resume Next
    StartDate = CDbl(destSheet.Range("B1").Value)
    EndDate = CDbl(destSheet.Range("B2").Value)
    On Error GoTo 0
    
    ' 日付が設定されているかチェック
    useFilter = (StartDate > 0) And (EndDate > 0)
    
    ' テーブルの検索
    tableFound = False
    On Error Resume Next
    Set srcTable = srcSheet.ListObjects("_品番別")
    On Error GoTo 0
    
    If Not srcTable Is Nothing Then
        tableFound = True
    End If
    
    ' テーブルが見つからない場合は、データ範囲を検索して変換
    If Not tableFound Then
        ' データ範囲を特定
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
                On Error GoTo 0
            End If
        End If
    End If
    
    ' それでもテーブルが見つからない場合は処理中止
    If Not tableFound Then
        Application.StatusBar = "通称別直接集計b: テーブル「_品番別」が見つかりません。"
        Exit Sub
    End If
    
    ' ステータスバーを更新
    Application.StatusBar = "通称別直接集計b: データ取得中..."
    
    ' 元データの取得
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
            Case "稼働"
                kadoCol = i
            Case "サイクル"
                cycleCol = i
            Case "ショット数"
                shotCol = i
            Case "不良数"
                furyoCol = i
            Case "打出し"
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
            Case "ゴミ押し"
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
            Case "割れ白化"
                wareHakukaCol = i
            Case "コアカス"
                coreKasuCol = i
            Case "その他"
                sonotaCol = i
            Case "チョコ停打出し"
                chocoCol = i
            Case "検査"
                kensaCol = i
            Case "流出不良"
                ryushutuCol = i
        End Select
    Next i
    
    ' 必要な列が見つからない場合は処理中止
    If hinbanCol = 0 Or dateCol = 0 Then
        Application.StatusBar = "通称別直接集計b: 必要な列が見つかりません。"
        Exit Sub
    End If
    
    ' ステータスバーを更新
    Application.StatusBar = "通称別直接集計b: データ集計中..."
    
    ' Dictionaryオブジェクトを作成
    Set dictGroups = CreateObject("Scripting.Dictionary")
    Set dictSums = CreateObject("Scripting.Dictionary")
    Set dictCounts = CreateObject("Scripting.Dictionary")
    
    ' データをグループ化して集計
    For i = 1 To UBound(srcData, 1)
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
                dictSums(tsusho)("稼働") = 0
                dictSums(tsusho)("サイクル") = 0
                dictCounts(tsusho)("サイクル") = 0  ' サイクル平均計算用
                dictSums(tsusho)("ショット数") = 0
                dictSums(tsusho)("不良数") = 0
                dictSums(tsusho)("打出し") = 0
                dictSums(tsusho)("ショート") = 0
                dictSums(tsusho)("ウエルド") = 0
                dictSums(tsusho)("シワ") = 0
                dictSums(tsusho)("異物") = 0
                dictSums(tsusho)("シルバー") = 0
                dictSums(tsusho)("フローマーク") = 0
                dictSums(tsusho)("ゴミ押し") = 0
                dictSums(tsusho)("GCカス") = 0
                dictSums(tsusho)("キズ") = 0
                dictSums(tsusho)("ヒケ") = 0
                dictSums(tsusho)("糸引き") = 0
                dictSums(tsusho)("型汚れ") = 0
                dictSums(tsusho)("マクレ") = 0
                dictSums(tsusho)("取出不良") = 0
                dictSums(tsusho)("割れ白化") = 0
                dictSums(tsusho)("コアカス") = 0
                dictSums(tsusho)("その他") = 0
                dictSums(tsusho)("チョコ停打出し") = 0
                dictSums(tsusho)("検査") = 0
                dictSums(tsusho)("流出不良") = 0
            End If
            
            ' 各項目の合計値を更新
            If kataKaeCol > 0 And IsNumeric(srcData(i, kataKaeCol)) Then
                dictSums(tsusho)("型替") = dictSums(tsusho)("型替") + CDbl(srcData(i, kataKaeCol))
            End If
            
            If kadoCol > 0 And IsNumeric(srcData(i, kadoCol)) Then
                dictSums(tsusho)("稼働") = dictSums(tsusho)("稼働") + CDbl(srcData(i, kadoCol))
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
                dictSums(tsusho)("打出し") = dictSums(tsusho)("打出し") + CDbl(srcData(i, uchidashiCol))
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
                dictSums(tsusho)("ゴミ押し") = dictSums(tsusho)("ゴミ押し") + CDbl(srcData(i, gomiCol))
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
                dictSums(tsusho)("割れ白化") = dictSums(tsusho)("割れ白化") + CDbl(srcData(i, wareHakukaCol))
            End If
            
            If coreKasuCol > 0 And IsNumeric(srcData(i, coreKasuCol)) Then
                dictSums(tsusho)("コアカス") = dictSums(tsusho)("コアカス") + CDbl(srcData(i, coreKasuCol))
            End If
            
            If sonotaCol > 0 And IsNumeric(srcData(i, sonotaCol)) Then
                dictSums(tsusho)("その他") = dictSums(tsusho)("その他") + CDbl(srcData(i, sonotaCol))
            End If
            
            If chocoCol > 0 And IsNumeric(srcData(i, chocoCol)) Then
                dictSums(tsusho)("チョコ停打出し") = dictSums(tsusho)("チョコ停打出し") + CDbl(srcData(i, chocoCol))
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
    Application.StatusBar = "通称別直接集計b: データ出力準備中..."
    
    ' 出力先シートの4行目以降をクリア（1-3行目は残す）
    ' 出力先シートの4行目から31行目までをクリア
    destSheet.Range("A4:AB31").Clear
    
    ' 4行目以降の書式設定
    With destSheet.Range("A4:AB" & destSheet.Rows.Count)
        .Font.Name = "Yu Gothic UI"
        .Font.Size = 11
    End With
    
    ' タイトル行の作成（4行目）
    destRow = 4
    If useFilter Then
        destSheet.Range("A" & destRow).Value = "成形通称別不良集計：" & Format(StartDate, "yyyy/mm/dd") & "～" & Format(EndDate, "yyyy/mm/dd")
    Else
        destSheet.Range("A" & destRow).Value = "成形通称別不良集計：全期間"
    End If
    destSheet.Range("A" & destRow).Font.Bold = True
    
    ' ヘッダー行の作成（5行目）
    destRow = 5
    destSheet.Range("A" & destRow).Value = "通称"
    destSheet.Range("B" & destRow).Value = "型替"
    destSheet.Range("C" & destRow).Value = "稼働"
    destSheet.Range("D" & destRow).Value = "サイクル"
    destSheet.Range("E" & destRow).Value = "ショット数"
    destSheet.Range("F" & destRow).Value = "不良数"
    destSheet.Range("G" & destRow).Value = "不良率"
    destSheet.Range("H" & destRow).Value = "打出し"
    destSheet.Range("I" & destRow).Value = "ショート"
    destSheet.Range("J" & destRow).Value = "ウエルド"
    destSheet.Range("K" & destRow).Value = "シワ"
    destSheet.Range("L" & destRow).Value = "異物"
    destSheet.Range("M" & destRow).Value = "シルバー"
    destSheet.Range("N" & destRow).Value = "フローマーク"
    destSheet.Range("O" & destRow).Value = "ゴミ押し"
    destSheet.Range("P" & destRow).Value = "GCカス"
    destSheet.Range("Q" & destRow).Value = "キズ"
    destSheet.Range("R" & destRow).Value = "ヒケ"
    destSheet.Range("S" & destRow).Value = "糸引き"
    destSheet.Range("T" & destRow).Value = "型汚れ"
    destSheet.Range("U" & destRow).Value = "マクレ"
    destSheet.Range("V" & destRow).Value = "取出不良"
    destSheet.Range("W" & destRow).Value = "割れ白化"
    destSheet.Range("X" & destRow).Value = "コアカス"
    destSheet.Range("Y" & destRow).Value = "その他"
    destSheet.Range("Z" & destRow).Value = "チョコ停打出し"
    destSheet.Range("AA" & destRow).Value = "検査"
    destSheet.Range("AB" & destRow).Value = "流出不良"
    
    ' ヘッダー行の書式設定
    With destSheet.Range("A" & destRow & ":AB" & destRow)
        .HorizontalAlignment = xlCenter  ' 中央揃え
        .Font.Bold = True
        .ShrinkToFit = True  ' 縮小して全体を表示
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
        
        ' 配列をソート（特定の順序で）
        ' TG → 62-28030Fr → 62-28030Rr → 62-58050Fr → 62-58050Rr → 補給品
        Dim sortedArr() As Variant
        ReDim sortedArr(0 To UBound(tsushoArr))
        Dim sortIdx As Integer
        sortIdx = 0
        
        ' 順番に配列を作成
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
        Application.StatusBar = "通称別直接集計b: データ出力中..."
        
        ' データの書き込み
        For i = 0 To sortIdx - 1
            key = sortedArr(i)
            
            ' 基本データを書き込み
            destSheet.Cells(destRow, 1).Value = key
            destSheet.Cells(destRow, 2).Value = dictSums(key)("型替")
            destSheet.Cells(destRow, 3).Value = dictSums(key)("稼働")
            
            ' サイクルの平均値を計算
            If dictCounts(key)("サイクル") > 0 Then
                destSheet.Cells(destRow, 4).Value = dictSums(key)("サイクル") / dictCounts(key)("サイクル")
            Else
                destSheet.Cells(destRow, 4).Value = 0
            End If
            
            destSheet.Cells(destRow, 5).Value = dictSums(key)("ショット数")
            destSheet.Cells(destRow, 6).Value = dictSums(key)("不良数")
            
            ' 不良率の計算（不良数÷ショット数）
            If dictSums(key)("ショット数") > 0 Then
                destSheet.Cells(destRow, 7).Value = dictSums(key)("不良数") / dictSums(key)("ショット数")
            Else
                destSheet.Cells(destRow, 7).Value = 0
            End If
            
            ' 不良項目データを率として計算して書き込み
            ' すべて「項目値÷ショット数」で計算
            If dictSums(key)("ショット数") > 0 Then
                destSheet.Cells(destRow, 8).Value = dictSums(key)("打出し") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 9).Value = dictSums(key)("ショート") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 10).Value = dictSums(key)("ウエルド") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 11).Value = dictSums(key)("シワ") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 12).Value = dictSums(key)("異物") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 13).Value = dictSums(key)("シルバー") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 14).Value = dictSums(key)("フローマーク") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 15).Value = dictSums(key)("ゴミ押し") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 16).Value = dictSums(key)("GCカス") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 17).Value = dictSums(key)("キズ") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 18).Value = dictSums(key)("ヒケ") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 19).Value = dictSums(key)("糸引き") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 20).Value = dictSums(key)("型汚れ") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 21).Value = dictSums(key)("マクレ") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 22).Value = dictSums(key)("取出不良") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 23).Value = dictSums(key)("割れ白化") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 24).Value = dictSums(key)("コアカス") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 25).Value = dictSums(key)("その他") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 26).Value = dictSums(key)("チョコ停打出し") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 27).Value = dictSums(key)("検査") / dictSums(key)("ショット数")
                destSheet.Cells(destRow, 28).Value = dictSums(key)("流出不良") / dictSums(key)("ショット数")
            Else
                ' ショット数が0の場合はすべて0
                For j = 8 To 28
                    destSheet.Cells(destRow, j).Value = 0
                Next j
            End If
            
            destRow = destRow + 1
        Next i
        
        dataEndRow = destRow - 1
    End If
    
    ' テーブルの作成
    Set tableRange = destSheet.Range("A5").Resize(dataEndRow - 4, 28)
    
    ' すでに同名のテーブルが存在する場合は削除
    On Error Resume Next
    If Not destSheet.ListObjects("_品番別bb") Is Nothing Then
        destSheet.ListObjects("_品番別bb").Delete
    End If
    On Error GoTo 0
    
    Set destTable = destSheet.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    destTable.Name = "_品番別bb"
    destTable.ShowAutoFilter = False  ' フィルターボタンを非表示
    
    ' ステータスバーを更新
    Application.StatusBar = "通称別直接集計b: 書式設定中..."
    
    ' テーブル内の書式設定
    With destSheet.Range("A" & dataStartRow & ":AB" & dataEndRow)
        .ShrinkToFit = True  ' 縮小して全体を表示
    End With
    
    ' 不良率列のフォーマット設定（%表示、小数点以下2桁）
    destSheet.Range("G" & dataStartRow & ":G" & dataEndRow).NumberFormat = "0.00%"
    
    ' サイクル列のフォーマット設定（小数点以下1桁）
    destSheet.Range("D" & dataStartRow & ":D" & dataEndRow).NumberFormat = "0.0"
    
    ' 不良項目列のフォーマット設定（%表示、小数点以下2桁）
    destSheet.Range("H" & dataStartRow & ":AB" & dataEndRow).NumberFormat = "0.00%"
    
    ' 列幅の設定
    destSheet.Columns("A").ColumnWidth = 14  ' 通称列
    destSheet.Columns("B:G").ColumnWidth = 7  ' 型替～不良率列
    destSheet.Columns("H:AB").ColumnWidth = 3  ' 不良項目列（指定通り5に設定）
    
    ' 0の値を薄いグレーにする条件付き書式
    With destSheet.Range("B" & dataStartRow & ":AB" & dataEndRow)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="0"
        .FormatConditions(1).Font.Color = RGB(192, 192, 192)  ' 薄いグレー
    End With
    
    ' 処理完了
    Application.StatusBar = "通称別直接集計b: 処理が完了しました。"
    
    ' 1秒待機してステータスバークリア
    Application.Wait Now + TimeValue("00:00:01")
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    Application.StatusBar = False
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
End Sub
