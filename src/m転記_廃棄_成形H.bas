Attribute VB_Name = "m転記_廃棄_成形H"
Option Explicit

' ========================================
' マクロ名: 転記_廃棄_成形H
' 処理概要: 廃棄データを期間別に9分類で集計して成形Hシートに転記（合計列付き）
'
' 【処理の特徴】
' 1. 空白期間スキップ：集計期間テーブルに行があっても、該当期間内にデータがなければテーブルを作らない
' 2. 動的期間対応：集計期間テーブルの行数が変わっても自動的に対応（増減どちらもOK）
' 3. 高速化：配列処理による大量データの高速集計
' 4. ワースト順機能：項目テーブルの「ワースト」設定に応じて動的に出力順序を変更
' 5. 合計列追加：9品番の右に合計列を追加して各行の合計値を自動計算
'
' 【テーブル構成】
' 期間テーブル : シート「成形H」、テーブル「_廃棄項目成形H」
' ソーステーブル : シート「廃棄」、テーブル「_廃棄」；シート「ロット数量」、テーブル「_ロット数量」
' 項目テーブル : シート「成形H」、テーブル「_集計期間成形H」
' 出力テーブル : シート「成形H」、複数テーブル「_廃棄H_成形_{期間名}」
'
' 【処理フロー】
' 1. 既存出力テーブルとデータを完全削除
' 2. ワースト設定（全項目 or 数値N）を読み込み
' 3. 各期間ごとに日付フィルター + 品番による9分類集計
' 4. 集計結果を降順ソートしてワースト順出力
' 5. データがある期間のみテーブル出力（空白期間はスキップ）
' 6. 各行に9品番の合計列を追加
'
' 【出力形式】
' - 1行目：ショット数（「_ロット数量」テーブルで工程=加工、品番2で9分類照合）+ 合計
' - 2行目：不良数（「_廃棄」テーブルの「件数」列を品番2で集計）+ 合計
' - 3行目以降：ワースト順で項目別集計 + 合計
'   - 「全項目」設定：0でない項目を降順で全て出力
'   - 数値N設定：上位N件 + 「その他」行（N+1行、ただし0でない項目数<=Nなら「その他」なし）
'
' 【集計条件】
' - 「_廃棄」テーブル：工程=成形
' - 品番照合：完全一致
' ========================================

Sub 転記_廃棄_成形H()
    ' ============================================
    ' 最適化設定の保存と適用
    ' ============================================
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    Dim origEnableEvents As Boolean
    Dim origDisplayAlerts As Boolean

    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation
    origEnableEvents = Application.EnableEvents
    origDisplayAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    On Error GoTo ErrorHandler
    Application.StatusBar = "廃棄成形H転記処理を開始します..."

    ' ============================================
    ' シートとテーブルの参照取得
    ' ============================================
    Dim wsSource As Worksheet, wsLot As Worksheet, wsTarget As Worksheet
    Set wsSource = ThisWorkbook.Worksheets("廃棄")
    Set wsLot = ThisWorkbook.Worksheets("ロット数量")
    Set wsTarget = ThisWorkbook.Worksheets("成形H")

    Dim tblSource As ListObject, tblLot As ListObject, tblItems As ListObject, tblPeriod As ListObject
    On Error Resume Next
    Set tblSource = wsSource.ListObjects("_廃棄")
    Set tblLot = wsLot.ListObjects("_ロット数量")
    Set tblItems = wsTarget.ListObjects("_廃棄項目成形H")
    Set tblPeriod = wsTarget.ListObjects("_集計期間成形H")
    On Error GoTo ErrorHandler

    ' 必須テーブルチェック
    If tblSource Is Nothing Then
        MsgBox "シート「廃棄」にテーブル「_廃棄」が見つかりません。", vbCritical
        GoTo Cleanup
    End If

    If tblLot Is Nothing Then
        MsgBox "シート「ロット数量」にテーブル「_ロット数量」が見つかりません。", vbCritical
        GoTo Cleanup
    End If

    ' ============================================
    ' ワースト設定の読み込み
    ' ============================================
    Dim worstSetting As String
    Dim worstNum As Long
    Dim isAllItems As Boolean

    worstSetting = ""
    worstNum = 0
    isAllItems = False

    If Not tblItems Is Nothing Then
        If Not tblItems.DataBodyRange Is Nothing Then
            Dim colWorstIdx As Long
            On Error Resume Next
            colWorstIdx = tblItems.ListColumns("ワースト").Index
            On Error GoTo ErrorHandler

            If colWorstIdx > 0 Then
                worstSetting = Trim(CStr(tblItems.DataBodyRange.Cells(1, colWorstIdx).Value))

                If worstSetting = "全項目" Then
                    isAllItems = True
                ElseIf IsNumeric(worstSetting) Then
                    worstNum = CLng(worstSetting)
                    If worstNum <= 0 Then
                        MsgBox "「ワースト」列の数値は1以上を指定してください。", vbCritical
                        GoTo Cleanup
                    End If
                Else
                    MsgBox "「ワースト」列の値が不正です。「全項目」または数値を指定してください。", vbCritical
                    GoTo Cleanup
                End If
            End If
        End If
    End If

    ' ワースト設定が取得できない場合はデフォルト（全項目）
    If worstSetting = "" Then
        isAllItems = True
        Application.StatusBar = "ワースト設定が見つかりません。全項目モードで実行します..."
    End If

    ' ============================================
    ' 期間テーブルの読み込み
    ' ============================================
    Dim periodCount As Long
    periodCount = 0
    Dim periodInfo() As Variant

    If Not tblPeriod Is Nothing Then
        If Not tblPeriod.DataBodyRange Is Nothing Then
            periodCount = tblPeriod.DataBodyRange.Rows.Count
            If periodCount > 0 Then
                ReDim periodInfo(1 To periodCount, 1 To 3)
                Dim p As Long
                For p = 1 To periodCount
                    periodInfo(p, 1) = CStr(tblPeriod.DataBodyRange.Cells(p, 1).Value) ' 期間名
                    periodInfo(p, 2) = tblPeriod.DataBodyRange.Cells(p, 2).Value       ' 開始日
                    periodInfo(p, 3) = tblPeriod.DataBodyRange.Cells(p, 3).Value       ' 終了日
                Next p
            End If
        End If
    End If

    If periodCount = 0 Then
        MsgBox "「_集計期間成形H」に有効な集計期間がありません。処理を中止します。", vbExclamation
        GoTo Cleanup
    End If

    ' ============================================
    ' ソーステーブルのデータ範囲取得と列インデックス
    ' ============================================
    Dim srcData As Range
    Set srcData = tblSource.DataBodyRange
    If srcData Is Nothing Then
        Application.StatusBar = "ソーステーブルにデータがありません"
        GoTo Cleanup
    End If

    Dim colHizuke As Long, colHinban2 As Long, colKoutei As Long
    Dim colFuryouNaiyou As Long, colKensuu As Long
    colHizuke = tblSource.ListColumns("日付").Index
    colHinban2 = tblSource.ListColumns("品番2").Index
    colKoutei = tblSource.ListColumns("工程").Index
    colFuryouNaiyou = tblSource.ListColumns("不良内容").Index
    colKensuu = tblSource.ListColumns("件数").Index

    ' ============================================
    ' ロット数量テーブルのデータ範囲取得と列インデックス
    ' ============================================
    Dim lotData As Range
    Set lotData = tblLot.DataBodyRange

    Dim colLotHizuke As Long, colLotKoutei As Long, colLotHinban2 As Long, colLotSuuryou As Long
    If Not lotData Is Nothing Then
        colLotHizuke = tblLot.ListColumns("日付").Index
        colLotKoutei = tblLot.ListColumns("工程").Index
        colLotHinban2 = tblLot.ListColumns("品番2").Index
        colLotSuuryou = tblLot.ListColumns("ロット数量").Index
    End If

    ' ============================================
    ' 品番分類リストの定義（9分類）
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

    ' ============================================
    ' 既存の出力テーブルオブジェクトを削除
    ' ============================================
    Dim idxLO As Long
    For idxLO = wsTarget.ListObjects.Count To 1 Step -1
        Dim loTemp As ListObject
        Set loTemp = wsTarget.ListObjects(idxLO)
        If loTemp.Name Like "_廃棄H_成形_*" Then
            loTemp.Delete
        End If
    Next idxLO

    ' ============================================
    ' 既存出力範囲の行削除
    ' ============================================
    Dim itemsTableLastRow As Long, periodTableLastRow As Long
    itemsTableLastRow = 0
    If Not tblItems Is Nothing Then
        itemsTableLastRow = tblItems.Range.Row + tblItems.Range.Rows.Count - 1
    End If

    periodTableLastRow = 0
    If Not tblPeriod Is Nothing Then
        periodTableLastRow = tblPeriod.Range.Row + tblPeriod.Range.Rows.Count - 1
    End If

    Dim baseRow As Long
    If itemsTableLastRow > periodTableLastRow Then
        baseRow = itemsTableLastRow
    Else
        baseRow = periodTableLastRow
    End If
    If baseRow < 1 Then baseRow = 1

    Dim lastUsedRow As Long
    lastUsedRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
    If lastUsedRow >= baseRow + 1 Then
        wsTarget.Rows((baseRow + 1) & ":" & lastUsedRow).Delete
    End If

    ' ============================================
    ' 出力開始位置の決定
    ' ============================================
    Dim currentRow As Long
    currentRow = baseRow + 3

    ' ============================================
    ' 全グループ配列の定義
    ' ============================================
    Dim allGroups As Variant
    allGroups = Array("58050FrLH", "58050FrRH", "58050RrLH", "58050RrRH", _
                      "28050FrLH", "28050FrRH", "28050RrLH", "28050RrRH", "補給品")

    ' ============================================
    ' ソースデータを配列に取り込み
    ' ============================================
    Dim srcArr As Variant
    srcArr = srcData.Value

    Dim lotArr As Variant
    If Not lotData Is Nothing Then
        lotArr = lotData.Value
    End If

    ' ============================================
    ' 印刷範囲の記録用変数
    ' ============================================
    Dim printRangeStart As Long
    Dim printRangeEnd As Long
    printRangeStart = 0
    printRangeEnd = 0

    ' ============================================
    ' 各期間の処理ループ
    ' ============================================
    Dim periodIdx As Long
    For periodIdx = 1 To periodCount
        Application.StatusBar = "期間 " & periodIdx & "/" & periodCount & " を処理中..."

        Dim periodName As String, startDate As Date, endDate As Date
        periodName = CStr(periodInfo(periodIdx, 1))
        startDate = CDate(periodInfo(periodIdx, 2))
        endDate = CDate(periodInfo(periodIdx, 3))

        ' ============================================
        ' グループ別集計用Dictionaryの初期化
        ' ============================================
        Dim aggShot As Object, aggFuryo As Object, aggItems As Object
        Set aggShot = CreateObject("Scripting.Dictionary")
        Set aggFuryo = CreateObject("Scripting.Dictionary")
        Set aggItems = CreateObject("Scripting.Dictionary")

        Dim grp As Variant
        For Each grp In allGroups
            aggShot(CStr(grp)) = 0
            aggFuryo(CStr(grp)) = 0
            Set aggItems(CStr(grp)) = CreateObject("Scripting.Dictionary")
        Next grp

        ' ============================================
        ' ロット数量テーブルからショット数を集計
        ' ============================================
        If Not lotData Is Nothing Then
            Dim r As Long
            For r = 1 To UBound(lotArr, 1)
                Dim lotDate As Variant
                lotDate = lotArr(r, colLotHizuke)

                If IsDate(lotDate) Then
                    Dim dt As Date
                    dt = CDate(lotDate)

                    If dt >= startDate And dt <= endDate Then
                        Dim lotKoutei As String
                        lotKoutei = Trim(CStr(lotArr(r, colLotKoutei)))

                        If lotKoutei = "加工" Then
                            Dim hinban2Lot As String
                            hinban2Lot = Trim(CStr(lotArr(r, colLotHinban2)))

                            If hinbanList.Exists(hinban2Lot) Then
                                Dim lotQty As Double
                                If IsNumeric(lotArr(r, colLotSuuryou)) Then
                                    lotQty = CDbl(lotArr(r, colLotSuuryou))
                                    aggShot(hinban2Lot) = aggShot(hinban2Lot) + lotQty
                                End If
                            End If
                        End If
                    End If
                End If
            Next r
        End If

        ' ============================================
        ' 空白期間判定フラグ
        ' ============================================
        Dim hasData As Boolean
        hasData = False

        ' ============================================
        ' 廃棄テーブルから不良数と項目別集計
        ' ============================================
        For r = 1 To UBound(srcArr, 1)
            Dim srcDate As Variant
            srcDate = srcArr(r, colHizuke)

            If IsDate(srcDate) Then
                Dim srcDt As Date
                srcDt = CDate(srcDate)

                If srcDt >= startDate And srcDt <= endDate Then
                    ' 条件チェック：工程=成形
                    Dim koutei As String
                    koutei = Trim(CStr(srcArr(r, colKoutei)))

                    If koutei = "成形" Then
                        Dim hinban2 As String
                        hinban2 = Trim(CStr(srcArr(r, colHinban2)))

                        If hinbanList.Exists(hinban2) Then
                            Dim kensuu As Double
                            If IsNumeric(srcArr(r, colKensuu)) Then
                                kensuu = CDbl(srcArr(r, colKensuu))

                                ' データありフラグ
                                If kensuu <> 0 Then
                                    hasData = True
                                End If

                                ' 不良数に加算
                                aggFuryo(hinban2) = aggFuryo(hinban2) + kensuu

                                ' 不良内容による項目別集計
                                Dim furyouNaiyou As String
                                furyouNaiyou = Trim(CStr(srcArr(r, colFuryouNaiyou)))

                                ' 空欄の場合は「（空白）」として集計
                                If Len(furyouNaiyou) = 0 Then furyouNaiyou = "（空白）"
                                If Not aggItems(hinban2).Exists(furyouNaiyou) Then
                                        aggItems(hinban2)(furyouNaiyou) = 0
                                    End If
                                    aggItems(hinban2)(furyouNaiyou) = aggItems(hinban2)(furyouNaiyou) + kensuu
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            ' 進捗表示（200行ごと）
            If (r Mod 200) = 0 Then
                Application.StatusBar = "期間 " & periodIdx & "/" & periodCount & " - " & r & "/" & UBound(srcArr, 1) & " 行処理中..."
            End If
        Next r

        ' ============================================
        ' 空白期間スキップ処理
        ' ============================================
        If Not hasData Then
            Application.StatusBar = "期間 " & periodIdx & " はデータ無しのためスキップします..."
            GoTo NextPeriod
        End If

        ' ============================================
        ' 印刷範囲の開始位置を記録（最初のテーブルのみ）
        ' ============================================
        If printRangeStart = 0 Then
            printRangeStart = currentRow
        End If

        ' ============================================
        ' テーブル出力：タイトル行
        ' ============================================
        Dim titleText As String
        titleText = "廃棄H_成形_" & periodName & "_" & Format(startDate, "m/d") & "‾" & Format(endDate, "m/d")

        With wsTarget.Cells(currentRow, 1)
            .Value = titleText
            .ShrinkToFit = False
            .WrapText = False
            .Font.Bold = True
            .Font.Size = 12
        End With

        ' ============================================
        ' テーブル出力：ヘッダー行
        ' ============================================
        Dim outputStartRow As Long
        outputStartRow = currentRow + 1

        wsTarget.Cells(outputStartRow, 1).Value = "項目"

        Dim colOffset As Long
        colOffset = 2
        For Each grp In allGroups
            With wsTarget.Cells(outputStartRow, colOffset)
                .Value = CStr(grp)
                .ShrinkToFit = True
            End With
            colOffset = colOffset + 1
        Next grp

        ' 合計列のヘッダー追加
        With wsTarget.Cells(outputStartRow, colOffset)
            .Value = "合計"
            .ShrinkToFit = True
        End With

        ' ============================================
        ' テーブル出力：データ行
        ' ============================================
        Dim dataStartRow As Long
        dataStartRow = outputStartRow + 1
        Dim rowIdx As Long
        rowIdx = dataStartRow

        ' 1行目：ショット数
        With wsTarget.Cells(rowIdx, 1)
            .Value = "ショット数"
            .ShrinkToFit = True
        End With
        colOffset = 2
        Dim shotTotal As Double
        shotTotal = 0
        For Each grp In allGroups
            Dim shotVal As Double
            shotVal = aggShot(CStr(grp))
            wsTarget.Cells(rowIdx, colOffset).Value = shotVal
            shotTotal = shotTotal + shotVal
            colOffset = colOffset + 1
        Next grp
        ' 合計列
        wsTarget.Cells(rowIdx, colOffset).Value = shotTotal
        rowIdx = rowIdx + 1

        ' 2行目：不良数
        With wsTarget.Cells(rowIdx, 1)
            .Value = "不良数"
            .ShrinkToFit = True
        End With
        colOffset = 2
        Dim furyoTotal As Double
        furyoTotal = 0
        For Each grp In allGroups
            Dim furyoVal As Double
            furyoVal = aggFuryo(CStr(grp))
            wsTarget.Cells(rowIdx, colOffset).Value = furyoVal
            furyoTotal = furyoTotal + furyoVal
            colOffset = colOffset + 1
        Next grp
        ' 合計列
        wsTarget.Cells(rowIdx, colOffset).Value = furyoTotal
        rowIdx = rowIdx + 1

        ' ============================================
        ' 3行目以降：ワースト順で項目別集計
        ' ============================================

        ' 全グループの項目別合計を計算
        Dim totalItems As Object
        Set totalItems = CreateObject("Scripting.Dictionary")

        For Each grp In allGroups
            Dim itemKey As Variant
            For Each itemKey In aggItems(CStr(grp)).Keys
                If Not totalItems.Exists(CStr(itemKey)) Then
                    totalItems(CStr(itemKey)) = 0
                End If
                totalItems(CStr(itemKey)) = totalItems(CStr(itemKey)) + CDbl(aggItems(CStr(grp))(itemKey))
            Next itemKey
        Next grp

        ' 全グループ合計を配列化して降順ソート
        Dim totalArr() As Variant
        Dim totalCount As Long
        totalCount = totalItems.Count

        If totalCount > 0 Then
            ReDim totalArr(1 To totalCount, 1 To 2)
            Dim idx As Long
            idx = 1
            Dim totalKey As Variant
            For Each totalKey In totalItems.Keys
                totalArr(idx, 1) = CStr(totalKey)
                totalArr(idx, 2) = CDbl(totalItems(totalKey))
                idx = idx + 1
            Next totalKey

            ' 降順ソート
            Call BubbleSortDesc(totalArr)

            ' ============================================
            ' ワースト順出力の実行
            ' ============================================

            ' 出力する項目リストを作成
            Dim outputItemList() As String
            Dim outputItemCount As Long
            Dim hasSonotaRow As Boolean

            hasSonotaRow = False
            outputItemCount = 0

            ' 0でない項目だけをフィルタリング
            Dim nonZeroCount As Long
            nonZeroCount = 0
            Dim i2 As Long
            For i2 = 1 To UBound(totalArr, 1)
                If CDbl(totalArr(i2, 2)) <> 0 Then
                    nonZeroCount = nonZeroCount + 1
                End If
            Next i2

            ' ワースト設定に応じて出力項目を決定
            If isAllItems Then
                ' 「全項目」モード
                outputItemCount = nonZeroCount
                If outputItemCount > 0 Then
                    ReDim outputItemList(1 To outputItemCount)
                    Dim outIdx As Long
                    outIdx = 1
                    For i2 = 1 To UBound(totalArr, 1)
                        If CDbl(totalArr(i2, 2)) <> 0 Then
                            outputItemList(outIdx) = CStr(totalArr(i2, 1))
                            outIdx = outIdx + 1
                        End If
                    Next i2
                End If
            Else
                ' 数値Nモード
                If nonZeroCount > worstNum Then
                    outputItemCount = worstNum
                    ReDim outputItemList(1 To outputItemCount)
                    For i2 = 1 To worstNum
                        outputItemList(i2) = CStr(totalArr(i2, 1))
                    Next i2
                    hasSonotaRow = True
                Else
                    outputItemCount = nonZeroCount
                    If outputItemCount > 0 Then
                        ReDim outputItemList(1 To outputItemCount)
                        outIdx = 1
                        For i2 = 1 To UBound(totalArr, 1)
                            If CDbl(totalArr(i2, 2)) <> 0 Then
                                outputItemList(outIdx) = CStr(totalArr(i2, 1))
                                outIdx = outIdx + 1
                            End If
                        Next i2
                    End If
                End If
            End If

            ' ============================================
            ' 項目行の出力（ワースト順）
            ' ============================================
            Dim outItem As Long
            For outItem = 1 To outputItemCount
                Dim currentItemName As String
                currentItemName = outputItemList(outItem)

                With wsTarget.Cells(rowIdx, 1)
                    .Value = currentItemName
                    .ShrinkToFit = True
                End With

                colOffset = 2
                Dim itemRowTotal As Double
                itemRowTotal = 0
                For Each grp In allGroups
                    Dim itemValue As Double
                    itemValue = 0

                    If aggItems(CStr(grp)).Exists(currentItemName) Then
                        itemValue = CDbl(aggItems(CStr(grp))(currentItemName))
                    End If

                    wsTarget.Cells(rowIdx, colOffset).Value = itemValue
                    itemRowTotal = itemRowTotal + itemValue
                    colOffset = colOffset + 1
                Next grp

                ' 合計列
                wsTarget.Cells(rowIdx, colOffset).Value = itemRowTotal

                rowIdx = rowIdx + 1
            Next outItem

            ' ============================================
            ' 「その他」行の出力（必要な場合のみ）
            ' ============================================
            If hasSonotaRow Then
                With wsTarget.Cells(rowIdx, 1)
                    .Value = "その他"
                    .ShrinkToFit = True
                End With

                colOffset = 2
                Dim sonotaRowTotal As Double
                sonotaRowTotal = 0
                For Each grp In allGroups
                    Dim sonotaSum As Double
                    sonotaSum = 0

                    Dim k As Long
                    For k = worstNum + 1 To UBound(totalArr, 1)
                        Dim sonotaItemName As String
                        sonotaItemName = CStr(totalArr(k, 1))

                        If aggItems(CStr(grp)).Exists(sonotaItemName) Then
                            sonotaSum = sonotaSum + CDbl(aggItems(CStr(grp))(sonotaItemName))
                        End If
                    Next k

                    wsTarget.Cells(rowIdx, colOffset).Value = sonotaSum
                    sonotaRowTotal = sonotaRowTotal + sonotaSum
                    colOffset = colOffset + 1
                Next grp

                ' 合計列
                wsTarget.Cells(rowIdx, colOffset).Value = sonotaRowTotal

                rowIdx = rowIdx + 1
            End If
        End If

        ' ============================================
        ' テーブル化
        ' ============================================
        Dim lastCol As Long
        lastCol = UBound(allGroups) + 3  ' 項目列 + 9品番 + 合計列

        Dim tableRange As Range
        On Error Resume Next
        Set tableRange = wsTarget.Range(wsTarget.Cells(outputStartRow, 1), wsTarget.Cells(rowIdx - 1, lastCol))
        On Error GoTo ErrorHandler

        If Not tableRange Is Nothing Then
            Dim baseName As String, tryName As String, tryIdx As Long
            baseName = "_廃棄H_成形_" & Replace(periodName, " ", "_")
            tryName = baseName
            tryIdx = 1
            Do While TableExists(wsTarget, tryName)
                tryIdx = tryIdx + 1
                tryName = baseName & "_" & tryIdx
            Loop

            Dim newTable As ListObject
            Set newTable = wsTarget.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
            newTable.Name = tryName

            On Error Resume Next
            newTable.TableStyle = "TableStyleLight21"
            newTable.ShowAutoFilter = False
            On Error GoTo ErrorHandler

            Dim cIdx As Long
            For cIdx = 1 To newTable.Range.Columns.Count
                newTable.Range.Columns(cIdx).ColumnWidth = 8
            Next cIdx
        End If

        ' ============================================
        ' 印刷範囲の終了位置を更新
        ' ============================================
        printRangeEnd = rowIdx - 1

        ' 次のテーブルの開始位置（2行空ける）
        currentRow = rowIdx + 2

NextPeriod:
    Next periodIdx

    ' ============================================
    ' 印刷範囲の設定
    ' ============================================
    If printRangeStart > 0 And printRangeEnd > 0 Then
        Dim printLastCol As Long
        printLastCol = UBound(allGroups) + 3  ' 項目列 + 9品番 + 合計列

        On Error Resume Next
        wsTarget.PageSetup.PrintArea = wsTarget.Range( _
            wsTarget.Cells(printRangeStart, 1), _
            wsTarget.Cells(printRangeEnd, printLastCol)).Address
        On Error GoTo ErrorHandler

        Application.StatusBar = "印刷範囲を設定しました"
    End If

    GoTo Cleanup

ErrorHandler:
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.DisplayAlerts = origDisplayAlerts
    Application.StatusBar = False

    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "転記_廃棄_成形H"
    Exit Sub

Cleanup:
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.DisplayAlerts = origDisplayAlerts
    Application.StatusBar = False
End Sub

' ============================================
' Private関数: TableExists
' ============================================
Private Function TableExists(ws As Worksheet, tblName As String) As Boolean
    Dim lo As ListObject
    TableExists = False

    If ws Is Nothing Then Exit Function

    For Each lo In ws.ListObjects
        If lo.Name = tblName Then
            TableExists = True
            Exit Function
        End If
    Next lo
End Function

' ============================================
' Private関数: BubbleSortDesc
' 目的：2次元配列を2列目（値）の降順でソート
' ============================================
Private Sub BubbleSortDesc(ByRef arr() As Variant)
    Dim i As Long, j As Long
    Dim tempName As String, tempValue As Double
    Dim n As Long

    n = UBound(arr, 1)

    For i = 1 To n - 1
        For j = i + 1 To n
            If CDbl(arr(i, 2)) < CDbl(arr(j, 2)) Then
                tempName = arr(i, 1)
                tempValue = arr(i, 2)

                arr(i, 1) = arr(j, 1)
                arr(i, 2) = arr(j, 2)

                arr(j, 1) = tempName
                arr(j, 2) = tempValue
            End If
        Next j
    Next i
End Sub
