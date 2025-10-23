Attribute VB_Name = "m廃棄_塗装_期間対応"
Option Explicit

' ========================================
' マクロ名: 転記_廃棄_塗装_期間対応
' 処理概要: 「_集計期間塗装」テーブルの各期間に基づく複数期間集計・転記
' ソーステーブル: シート「廃棄」テーブル「_廃棄」
' 期間テーブル: シート「塗装」テーブル「_集計期間塗装」
' ターゲットテーブル: シート「塗装」複数テーブル「_廃棄_塗装_{期間}」
' 参照テーブル: シート「塗装」テーブル「_手直し項目塗装」
' 処理方式: 各期間の開始日～終了日で日付フィルタ後、品番2列による9分類と不良内容による項目別集計
' 出力形式: データがある期間のみテーブルを出力（空白期間はスキップ）
' 改善点:
'   - 期間テーブルから動的に期間名を取得してテーブル名・タイトルを生成
'   - 手直しテーブルを動的検索して水平配置基準位置を決定
' ========================================

Sub 転記_廃棄_塗装_期間対応()
    ' 最適化設定の保存
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    Dim origEnableEvents As Boolean
    Dim origDisplayAlerts As Boolean
    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation
    origEnableEvents = Application.EnableEvents
    origDisplayAlerts = Application.DisplayAlerts

    ' 最適化設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' エラーハンドリング設定
    On Error GoTo ErrorHandler

    ' ステータスバー初期化
    Application.StatusBar = "廃棄期間対応転記処理（塗装）を開始します..."

    ' ============================================
    ' シートとテーブルの参照取得
    ' ============================================
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Set wsSource = ThisWorkbook.Worksheets("廃棄")
    Set wsTarget = ThisWorkbook.Worksheets("塗装")

    Dim tblSource As ListObject, tblItems As ListObject, tblPeriod As ListObject
    Set tblSource = wsSource.ListObjects("_廃棄")
    Set tblItems = wsTarget.ListObjects("_手直し項目塗装")
    Set tblPeriod = wsTarget.ListObjects("_集計期間塗装")

    ' ============================================
    ' 期間テーブルの読み込み
    ' ============================================
    Dim periodData As Range
    Set periodData = tblPeriod.DataBodyRange

    If periodData Is Nothing Then
        MsgBox "「_集計期間塗装」テーブルにデータがありません。", vbExclamation
        GoTo Cleanup
    End If

    ' 期間情報の配列作成（動的に全行対応）
    Dim periodCount As Long
    periodCount = periodData.Rows.Count

    Dim periodInfo() As Variant
    ReDim periodInfo(1 To periodCount, 1 To 3) ' 期間, 開始日, 終了日

    Dim p As Long
    For p = 1 To periodCount
        periodInfo(p, 1) = CStr(periodData.Cells(p, 1).Value) ' 期間
        periodInfo(p, 2) = periodData.Cells(p, 2).Value       ' 開始日
        periodInfo(p, 3) = periodData.Cells(p, 3).Value       ' 終了日
    Next p

    ' ============================================
    ' ソーステーブルの列インデックス取得
    ' ============================================
    Dim srcData As Range
    Set srcData = tblSource.DataBodyRange

    If srcData Is Nothing Then
        Application.StatusBar = "ソーステーブルにデータがありません"
        GoTo Cleanup
    End If

    Dim colHizuke As Long, colHinban2 As Long, colKoutei As Long, colFuryouNaiyou As Long, colKensuu As Long
    colHizuke = tblSource.ListColumns("日付").Index
    colHinban2 = tblSource.ListColumns("品番2").Index
    colKoutei = tblSource.ListColumns("工程").Index
    colFuryouNaiyou = tblSource.ListColumns("不良内容").Index
    colKensuu = tblSource.ListColumns("件数").Index

    ' ============================================
    ' 項目テーブルから項目リスト取得（動的）
    ' ============================================
    Dim itemsData As Range
    Set itemsData = tblItems.DataBodyRange

    Dim itemsList As Object
    Set itemsList = CreateObject("Scripting.Dictionary")

    If Not itemsData Is Nothing Then
        Dim i As Long
        For i = 1 To itemsData.Rows.Count
            Dim itemName As String
            itemName = CStr(itemsData.Cells(i, 1).Value)
            If Len(itemName) > 0 Then
                itemsList(itemName) = i ' 項目名と順序を記録
            End If
        Next i
    End If

    ' 「その他」項目を最後に追加（存在しない場合のみ）
    If Not itemsList.Exists("その他") Then
        itemsList("その他") = itemsList.Count + 1
    End If

    ' ============================================
    ' 品番分類リストの定義（9分類：LH→RHの順序＋補給品）
    ' ============================================
    Dim hinbanList As Object
    Set hinbanList = CreateObject("Scripting.Dictionary")
    hinbanList("58050FrLH") = 1
    hinbanList("58050FrRH") = 2
    hinbanList("58050RrLH") = 3
    hinbanList("58050RrRH") = 4
    hinbanList("28050FrLH") = 5
    hinbanList("28050FrRH") = 6
    hinbanList("28050RrLH") = 7
    hinbanList("28050RrRH") = 8
    hinbanList("補給品") = 9

    ' 項目リストの順序ソート
    Dim sortedItems() As String
    ReDim sortedItems(0 To itemsList.Count - 1)

    Dim itemKey As Variant
    For Each itemKey In itemsList.Keys
        sortedItems(itemsList(itemKey) - 1) = CStr(itemKey)
    Next itemKey

    ' ============================================
    ' 既存の出力テーブルとタイトル行の完全削除（K列以降）
    ' ============================================
    On Error Resume Next

    ' 既存の期間別テーブルを削除（"_廃棄_塗装_"で始まるテーブル全て）
    Dim tbl As ListObject
    Dim tblsToDelete As Collection
    Set tblsToDelete = New Collection

    For Each tbl In wsTarget.ListObjects
        If Left(tbl.Name, 7) = "_廃棄_塗装_" Then
            tblsToDelete.Add tbl
        End If
    Next tbl

    Dim j As Long
    For j = 1 To tblsToDelete.Count
        tblsToDelete(j).Range.EntireRow.Delete
    Next j

    ' 既存のタイトル行（「廃棄_塗装」）をK列以降で検索して削除
    Dim searchRange As Range
    Set searchRange = wsTarget.Range("K:XFD")
    If Not searchRange Is Nothing Then
        Dim foundCell As Range
        Set foundCell = searchRange.Find("廃棄_塗装", LookIn:=xlValues, LookAt:=xlPart)
        Do While Not foundCell Is Nothing
            foundCell.EntireRow.Delete
            Set foundCell = searchRange.Find("廃棄_塗装", LookIn:=xlValues, LookAt:=xlPart)
        Loop
    End If

    Err.Clear
    On Error GoTo ErrorHandler

    ' ============================================
    ' 手直しテーブルの位置を動的検索して開始位置を決定
    ' ============================================
    Dim firstHandaosiTable As ListObject
    Dim startRow As Long

    ' 最初の手直しテーブルを動的に探す（"_手直し塗装_"で始まる最初のテーブル）
    For Each tbl In wsTarget.ListObjects
        If Left(tbl.Name, 7) = "_手直し塗装_" Then
            Set firstHandaosiTable = tbl
            Exit For
        End If
    Next tbl

    ' 手直しテーブルのタイトル行位置を取得（テーブルの1行上）
    If Not firstHandaosiTable Is Nothing Then
        startRow = firstHandaosiTable.Range.Row - 1
    Else
        ' 手直しテーブルがない場合は、項目テーブルから3行空ける
        If Not itemsData Is Nothing Then
            startRow = tblItems.Range.Row + tblItems.Range.Rows.Count - 1 + 3
        Else
            startRow = tblItems.Range.Row + 3
        End If
    End If

    Dim currentRow As Long
    currentRow = startRow

    ' ============================================
    ' 各期間の処理ループ
    ' ============================================
    For p = 1 To periodCount
        Application.StatusBar = "期間 " & p & "/" & periodCount & " を処理中..."

        Dim periodName As String, startDate As Date, endDate As Date
        periodName = periodInfo(p, 1)
        startDate = CDate(periodInfo(p, 2))
        endDate = CDate(periodInfo(p, 3))

        ' ============================================
        ' 空白期間判定フラグ
        ' ============================================
        Dim hasData As Boolean
        hasData = False

        ' ============================================
        ' 集計用辞書の初期化
        ' ============================================
        Dim aggregateData As Object
        Set aggregateData = CreateObject("Scripting.Dictionary")

        ' 辞書キーの初期化（項目×品番のマトリックス）
        Dim hinbanKey As Variant
        For Each itemKey In itemsList.Keys
            For Each hinbanKey In hinbanList.Keys
                Dim dictKey As String
                dictKey = CStr(itemKey) & "|" & CStr(hinbanKey)
                aggregateData(dictKey) = 0
            Next hinbanKey
        Next itemKey

        ' ============================================
        ' ソーステーブルの集計処理（日付フィルタ付き）
        ' ============================================
        For i = 1 To srcData.Rows.Count
            ' 日付チェック
            Dim rowDate As Date
            If IsDate(srcData.Cells(i, colHizuke).Value) Then
                rowDate = CDate(srcData.Cells(i, colHizuke).Value)

                If rowDate >= startDate And rowDate <= endDate Then
                    ' 工程チェック：工程=塗装
                    Dim koutei As String
                    koutei = CStr(srcData.Cells(i, colKoutei).Value)

                    If koutei = "塗装" Then
                        ' 品番2チェック
                        Dim hinban2 As String
                        hinban2 = CStr(srcData.Cells(i, colHinban2).Value)

                        If hinbanList.Exists(hinban2) Then
                            ' 不良内容取得と項目マッピング
                            Dim furyouNaiyou As String
                            furyouNaiyou = CStr(srcData.Cells(i, colFuryouNaiyou).Value)

                            Dim targetItem As String
                            If itemsList.Exists(furyouNaiyou) Then
                                targetItem = furyouNaiyou
                            Else
                                targetItem = "その他"
                            End If

                            ' 件数取得と加算
                            Dim kensuu As Double
                            If IsNumeric(srcData.Cells(i, colKensuu).Value) Then
                                kensuu = CDbl(srcData.Cells(i, colKensuu).Value)

                                ' データありフラグ
                                If kensuu <> 0 Then
                                    hasData = True
                                End If

                                dictKey = targetItem & "|" & hinban2
                                aggregateData(dictKey) = aggregateData(dictKey) + kensuu
                            End If
                        End If
                    End If
                End If
            End If
        Next i

        ' ============================================
        ' 空白期間スキップ処理
        ' ============================================
        If Not hasData Then
            Application.StatusBar = "期間 " & p & " はデータ無しのためスキップします..."
            GoTo NextPeriod
        End If

        ' ============================================
        ' タイトル行の生成と出力（K列）
        ' ============================================
        Dim titleText As String
        titleText = "廃棄_塗装_" & periodName & "_" & Format(startDate, "m/d") & "～" & Format(endDate, "m/d")

        ' タイトルセルの書式設定
        With wsTarget.Cells(currentRow, 11) ' K列 = 11
            .Value = titleText
            .ShrinkToFit = False  ' 縮小して全体を表示しない
            .WrapText = False     ' 折り返しなし
            .Font.Bold = True     ' 太字
            .Font.Size = 12       ' フォントサイズ12
        End With

        ' テーブル開始位置（K列）
        Dim startCell As Range
        Set startCell = wsTarget.Cells(currentRow + 1, 11) ' K列 = 11

        ' ============================================
        ' データ配列の作成（+1は集計行用）
        ' ============================================
        Dim outputData() As Variant
        ReDim outputData(0 To itemsList.Count + 1, 0 To 9)

        ' ヘッダー行設定
        Dim headers() As String
        ReDim headers(0 To 9)
        headers(0) = "項目"
        headers(1) = "58050FrLH"
        headers(2) = "58050FrRH"
        headers(3) = "58050RrLH"
        headers(4) = "58050RrRH"
        headers(5) = "28050FrLH"
        headers(6) = "28050FrRH"
        headers(7) = "28050RrLH"
        headers(8) = "28050RrRH"
        headers(9) = "補給品"

        For i = 0 To 9
            outputData(0, i) = headers(i)
        Next i

        ' 各列の合計用（期間ごとに初期化）
        Dim colSums(1 To 9) As Double
        Dim csIdx As Long
        For csIdx = 1 To 9
            colSums(csIdx) = 0
        Next csIdx

        ' データ行設定
        For i = 0 To UBound(sortedItems)
            outputData(i + 1, 0) = sortedItems(i) ' 項目名

            ' 各品番の数量
            For Each hinbanKey In hinbanList.Keys
                dictKey = sortedItems(i) & "|" & CStr(hinbanKey)
                Dim colValue As Double
                colValue = aggregateData(dictKey)
                outputData(i + 1, hinbanList(hinbanKey)) = colValue
                colSums(hinbanList(hinbanKey)) = colSums(hinbanList(hinbanKey)) + colValue
            Next hinbanKey
        Next i

        ' ============================================
        ' 集計行の追加
        ' ============================================
        Dim totalRow As Long
        totalRow = itemsList.Count + 1
        outputData(totalRow, 0) = "合計"
        For i = 1 To 9
            outputData(totalRow, i) = colSums(i)
        Next i

        ' ============================================
        ' テーブル範囲への書き込み
        ' ============================================
        Dim tableRange As Range
        Set tableRange = startCell.Resize(UBound(outputData, 1) + 1, UBound(outputData, 2) + 1)
        tableRange.Value = outputData

        ' ============================================
        ' ListObjectとして設定
        ' ============================================
        Dim newTable As ListObject
        Set newTable = wsTarget.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        newTable.Name = "_廃棄_塗装_" & periodName
        newTable.ShowAutoFilter = False ' フィルターボタンを非表示
        newTable.TableStyle = "TableStyleLight17" ' テーブルスタイル設定

        ' ============================================
        ' 書式設定：フォント、サイズ、列幅、表示形式
        ' ============================================
        With newTable.Range
            .Font.Name = "游ゴシック"
            .Font.Size = 11
            .ShrinkToFit = True ' 縮小して全体を表示
        End With

        ' 列幅設定
        For i = 1 To newTable.Range.Columns.Count
            newTable.Range.Columns(i).ColumnWidth = 8
        Next i

        ' 次のテーブル位置を計算（現在のテーブル + 3行間隔）
        currentRow = startCell.Row + UBound(outputData, 1) + 3

NextPeriod:
    Next p

    ' 処理完了のステータスバー表示
    Application.StatusBar = "廃棄期間対応転記処理が完了しました（" & periodCount & "期間）"
    Application.Wait Now + TimeValue("00:00:01")

    GoTo Cleanup

ErrorHandler:
    ' エラー情報の詳細化
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear

    MsgBox "エラーが発生しました" & vbCrLf & _
           "エラー番号: " & errNum & vbCrLf & _
           "詳細: " & errDesc, vbCritical, "転記_廃棄_塗装_期間対応 エラー"

Cleanup:
    ' 設定を確実に復元
    Application.StatusBar = False
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.DisplayAlerts = origDisplayAlerts
End Sub
