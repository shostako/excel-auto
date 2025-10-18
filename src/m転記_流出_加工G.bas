Attribute VB_Name = "m転記_流出_加工G"
Option Explicit

' ========================================
' マクロ名: 転記_流出_加工G
' 処理概要: 手直しと廃棄データを統合し期間別に9分類で集計して加工Gシートに転記（合計列付き）
' 集計条件: 手直し(発生=加工、品番3→品番2マッピング) + 廃棄(工程=加工、品番2)
' ========================================

Sub 転記_流出_加工G()
    Dim origScreenUpdating As Boolean, origCalculation As XlCalculation
    Dim origEnableEvents As Boolean, origDisplayAlerts As Boolean

    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation
    origEnableEvents = Application.EnableEvents
    origDisplayAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    On Error GoTo ErrorHandler
    Application.StatusBar = "流出加工G転記処理を開始します..."

    Dim wsHandaosi As Worksheet, wsHaiki As Worksheet, wsLot As Worksheet, wsTarget As Worksheet
    Set wsHandaosi = ThisWorkbook.Worksheets("手直し")
    Set wsHaiki = ThisWorkbook.Worksheets("廃棄")
    Set wsLot = ThisWorkbook.Worksheets("ロット数量")
    Set wsTarget = ThisWorkbook.Worksheets("加工G")

    Dim tblHandaosi As ListObject, tblHaiki As ListObject, tblLot As ListObject
    Dim tblItems As ListObject, tblPeriod As ListObject
    On Error Resume Next
    Set tblHandaosi = wsHandaosi.ListObjects("_手直し")
    Set tblHaiki = wsHaiki.ListObjects("_廃棄")
    Set tblLot = wsLot.ListObjects("_ロット数量")
    Set tblItems = wsTarget.ListObjects("_流出項目加工G")
    Set tblPeriod = wsTarget.ListObjects("_集計期間加工G")
    On Error GoTo ErrorHandler

    If tblHandaosi Is Nothing Then
        MsgBox "シート「手直し」にテーブル「_手直し」が見つかりません。", vbCritical
        GoTo Cleanup
    End If

    If tblHaiki Is Nothing Then
        MsgBox "シート「廃棄」にテーブル「_廃棄」が見つかりません。", vbCritical
        GoTo Cleanup
    End If

    If tblLot Is Nothing Then
        MsgBox "シート「ロット数量」にテーブル「_ロット数量」が見つかりません。", vbCritical
        GoTo Cleanup
    End If

    Dim worstSetting As String, worstNum As Long, isAllItems As Boolean
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

    If worstSetting = "" Then
        isAllItems = True
        Application.StatusBar = "ワースト設定が見つかりません。全項目モードで実行します..."
    End If

    Dim periodCount As Long, periodInfo() As Variant
    periodCount = 0

    If Not tblPeriod Is Nothing Then
        If Not tblPeriod.DataBodyRange Is Nothing Then
            periodCount = tblPeriod.DataBodyRange.Rows.Count
            If periodCount > 0 Then
                ReDim periodInfo(1 To periodCount, 1 To 3)
                Dim p As Long
                For p = 1 To periodCount
                    periodInfo(p, 1) = CStr(tblPeriod.DataBodyRange.Cells(p, 1).Value)
                    periodInfo(p, 2) = tblPeriod.DataBodyRange.Cells(p, 2).Value
                    periodInfo(p, 3) = tblPeriod.DataBodyRange.Cells(p, 3).Value
                Next p
            End If
        End If
    End If

    If periodCount = 0 Then
        MsgBox "「_集計期間加工G」に有効な集計期間がありません。処理を中止します。", vbExclamation
        GoTo Cleanup
    End If

    ' 手直しテーブルの列インデックス
    Dim handaosiData As Range
    Set handaosiData = tblHandaosi.DataBodyRange

    Dim colHdHizuke As Long, colHdHinban3 As Long, colHdHassei As Long
    Dim colHdMode2 As Long, colHdSuuryou As Long
    If Not handaosiData Is Nothing Then
        colHdHizuke = tblHandaosi.ListColumns("日付").Index
        colHdHinban3 = tblHandaosi.ListColumns("品番3").Index
        colHdHassei = tblHandaosi.ListColumns("発生").Index
        colHdMode2 = tblHandaosi.ListColumns("モード2").Index
        colHdSuuryou = tblHandaosi.ListColumns("数量").Index
    End If

    ' 廃棄テーブルの列インデックス
    Dim haikiData As Range
    Set haikiData = tblHaiki.DataBodyRange

    Dim colHkHizuke As Long, colHkHinban2 As Long, colHkKoutei As Long
    Dim colHkFuryou As Long, colHkKensuu As Long
    If Not haikiData Is Nothing Then
        colHkHizuke = tblHaiki.ListColumns("日付").Index
        colHkHinban2 = tblHaiki.ListColumns("品番2").Index
        colHkKoutei = tblHaiki.ListColumns("工程").Index
        colHkFuryou = tblHaiki.ListColumns("不良内容").Index
        colHkKensuu = tblHaiki.ListColumns("件数").Index
    End If

    ' ロット数量テーブルの列インデックス
    Dim lotData As Range
    Set lotData = tblLot.DataBodyRange

    Dim colLotHizuke As Long, colLotKoutei As Long, colLotHinban2 As Long, colLotSuuryou As Long
    If Not lotData Is Nothing Then
        colLotHizuke = tblLot.ListColumns("日付").Index
        colLotKoutei = tblLot.ListColumns("工程").Index
        colLotHinban2 = tblLot.ListColumns("品番2").Index
        colLotSuuryou = tblLot.ListColumns("ロット数量").Index
    End If

    ' 品番3→品番2マッピング（手直し用）
    Dim hinban3To2 As Object
    Set hinban3To2 = CreateObject("Scripting.Dictionary")
    hinban3To2("58050FrLH") = "58050FrLH"
    hinban3To2("58050FrRH") = "58050FrRH"
    hinban3To2("58050RrLH") = "58050RrLH"
    hinban3To2("58050RrRH") = "58050RrRH"
    hinban3To2("28050FrLH") = "28050FrLH"
    hinban3To2("28050FrRH") = "28050FrRH"
    hinban3To2("28050RrLH") = "28050RrLH"
    hinban3To2("28050RrRH") = "28050RrRH"

    ' 9品番リスト（品番2ベース）
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

    Dim idxLO As Long
    For idxLO = wsTarget.ListObjects.Count To 1 Step -1
        Dim loTemp As ListObject
        Set loTemp = wsTarget.ListObjects(idxLO)
        If loTemp.Name Like "_流出G_加工_*" Then
            loTemp.Delete
        End If
    Next idxLO

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

    Dim currentRow As Long
    currentRow = baseRow + 3

    Dim allGroups As Variant
    allGroups = Array("58050FrLH", "58050FrRH", "58050RrLH", "58050RrRH", _
                      "28050FrLH", "28050FrRH", "28050RrLH", "28050RrRH", "補給品")

    Dim handaosiArr As Variant, haikiArr As Variant, lotArr As Variant
    If Not handaosiData Is Nothing Then
        handaosiArr = handaosiData.Value
    End If

    If Not haikiData Is Nothing Then
        haikiArr = haikiData.Value
    End If

    If Not lotData Is Nothing Then
        lotArr = lotData.Value
    End If

    Dim printRangeStart As Long, printRangeEnd As Long
    printRangeStart = 0
    printRangeEnd = 0

    Dim periodIdx As Long
    For periodIdx = 1 To periodCount
        Application.StatusBar = "期間 " & periodIdx & "/" & periodCount & " を処理中..."

        Dim periodName As String, startDate As Date, endDate As Date
        periodName = CStr(periodInfo(periodIdx, 1))
        startDate = CDate(periodInfo(periodIdx, 2))
        endDate = CDate(periodInfo(periodIdx, 3))

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

        ' ロット数量からショット数集計
        If Not lotData Is Nothing Then
            Dim r As Long
            For r = 1 To UBound(lotArr, 1)
                Dim lotDate As Variant
                lotDate = lotArr(r, colLotHizuke)

                If IsDate(lotDate) Then
                    Dim dt As Date
                    dt = CDate(lotDate)

                    If dt >= startDate And dt <= endDate Then
                        Dim koutei As String
                        koutei = Trim(CStr(lotArr(r, colLotKoutei)))

                        If koutei = "加工" Then
                            Dim hinban2 As String
                            hinban2 = Trim(CStr(lotArr(r, colLotHinban2)))

                            If hinbanList.Exists(hinban2) Then
                                Dim lotQty As Double
                                If IsNumeric(lotArr(r, colLotSuuryou)) Then
                                    lotQty = CDbl(lotArr(r, colLotSuuryou))
                                    aggShot(hinban2) = aggShot(hinban2) + lotQty
                                End If
                            End If
                        End If
                    End If
                End If
            Next r
        End If

        Dim hasData As Boolean
        hasData = False

        ' 手直しテーブルから不良数と項目別集計（品番3→品番2マッピング）
        If Not handaosiData Is Nothing Then
            For r = 1 To UBound(handaosiArr, 1)
                Dim hdDate As Variant
                hdDate = handaosiArr(r, colHdHizuke)

                If IsDate(hdDate) Then
                    Dim hdDt As Date
                    hdDt = CDate(hdDate)

                    If hdDt >= startDate And hdDt <= endDate Then
                        Dim hdHassei As String
                        hdHassei = Trim(CStr(handaosiArr(r, colHdHassei)))

                        If hdHassei = "加工" Then
                            Dim hdHinban3 As String
                            hdHinban3 = Trim(CStr(handaosiArr(r, colHdHinban3)))

                            If hinban3To2.Exists(hdHinban3) Then
                                Dim hdHinban2 As String
                                hdHinban2 = hinban3To2(hdHinban3)

                                Dim hdSuuryou As Double
                                If IsNumeric(handaosiArr(r, colHdSuuryou)) Then
                                    hdSuuryou = CDbl(handaosiArr(r, colHdSuuryou))

                                    If hdSuuryou <> 0 Then
                                        hasData = True
                                    End If

                                    aggFuryo(hdHinban2) = aggFuryo(hdHinban2) + hdSuuryou

                                    Dim hdMode2 As String
                                    hdMode2 = Trim(CStr(handaosiArr(r, colHdMode2)))

                                    If Len(hdMode2) > 0 Then
                                        If Not aggItems(hdHinban2).Exists(hdMode2) Then
                                            aggItems(hdHinban2)(hdMode2) = 0
                                        End If
                                        aggItems(hdHinban2)(hdMode2) = aggItems(hdHinban2)(hdMode2) + hdSuuryou
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                If (r Mod 200) = 0 Then
                    Application.StatusBar = "期間 " & periodIdx & "/" & periodCount & " - 手直し " & r & "/" & UBound(handaosiArr, 1) & " 行処理中..."
                End If
            Next r
        End If

        ' 廃棄テーブルから不良数と項目別集計
        If Not haikiData Is Nothing Then
            For r = 1 To UBound(haikiArr, 1)
                Dim hkDate As Variant
                hkDate = haikiArr(r, colHkHizuke)

                If IsDate(hkDate) Then
                    Dim hkDt As Date
                    hkDt = CDate(hkDate)

                    If hkDt >= startDate And hkDt <= endDate Then
                        Dim hkKoutei As String
                        hkKoutei = Trim(CStr(haikiArr(r, colHkKoutei)))

                        If hkKoutei = "加工" Then
                            Dim hkHinban2 As String
                            hkHinban2 = Trim(CStr(haikiArr(r, colHkHinban2)))

                            If hinbanList.Exists(hkHinban2) Then
                                Dim hkKensuu As Double
                                If IsNumeric(haikiArr(r, colHkKensuu)) Then
                                    hkKensuu = CDbl(haikiArr(r, colHkKensuu))

                                    If hkKensuu <> 0 Then
                                        hasData = True
                                    End If

                                    aggFuryo(hkHinban2) = aggFuryo(hkHinban2) + hkKensuu

                                    Dim hkFuryou As String
                                    hkFuryou = Trim(CStr(haikiArr(r, colHkFuryou)))

                                    If Len(hkFuryou) > 0 Then
                                        If Not aggItems(hkHinban2).Exists(hkFuryou) Then
                                            aggItems(hkHinban2)(hkFuryou) = 0
                                        End If
                                        aggItems(hkHinban2)(hkFuryou) = aggItems(hkHinban2)(hkFuryou) + hkKensuu
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

                If (r Mod 200) = 0 Then
                    Application.StatusBar = "期間 " & periodIdx & "/" & periodCount & " - 廃棄 " & r & "/" & UBound(haikiArr, 1) & " 行処理中..."
                End If
            Next r
        End If

        If Not hasData Then
            Application.StatusBar = "期間 " & periodIdx & " はデータなしのためスキップします..."
            GoTo NextPeriod
        End If

        If printRangeStart = 0 Then
            printRangeStart = currentRow
        End If

        Dim titleText As String
        titleText = "流出G_加工_" & periodName & "_" & Format(startDate, "m/d") & "~" & Format(endDate, "m/d")

        With wsTarget.Cells(currentRow, 1)
            .Value = titleText
            .ShrinkToFit = False
            .WrapText = False
            .Font.Bold = True
            .Font.Size = 12
        End With

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

        With wsTarget.Cells(outputStartRow, colOffset)
            .Value = "合計"
            .ShrinkToFit = True
        End With

        Dim dataStartRow As Long, rowIdx As Long
        dataStartRow = outputStartRow + 1
        rowIdx = dataStartRow

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
        wsTarget.Cells(rowIdx, colOffset).Value = shotTotal
        rowIdx = rowIdx + 1

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
        wsTarget.Cells(rowIdx, colOffset).Value = furyoTotal
        rowIdx = rowIdx + 1

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

        Dim totalArr() As Variant, totalCount As Long
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

            Call QuickSortDesc(totalArr, 1, totalCount)

            Dim outputItemList() As String, outputItemCount As Long, hasSonotaRow As Boolean
            hasSonotaRow = False
            outputItemCount = 0

            Dim nonZeroCount As Long, i2 As Long
            nonZeroCount = 0
            For i2 = 1 To UBound(totalArr, 1)
                If CDbl(totalArr(i2, 2)) <> 0 Then
                    nonZeroCount = nonZeroCount + 1
                End If
            Next i2

            If isAllItems Then
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

            Dim outItem As Long
            For outItem = 1 To outputItemCount
                Dim currentItemName As String
                currentItemName = outputItemList(outItem)

                With wsTarget.Cells(rowIdx, 1)
                    .Value = currentItemName
                    .ShrinkToFit = True
                End With

                colOffset = 2
                Dim itemTotal As Double
                itemTotal = 0
                For Each grp In allGroups
                    Dim itemValue As Double
                    itemValue = 0

                    If aggItems(CStr(grp)).Exists(currentItemName) Then
                        itemValue = CDbl(aggItems(CStr(grp))(currentItemName))
                    End If

                    wsTarget.Cells(rowIdx, colOffset).Value = itemValue
                    itemTotal = itemTotal + itemValue
                    colOffset = colOffset + 1
                Next grp
                wsTarget.Cells(rowIdx, colOffset).Value = itemTotal
                rowIdx = rowIdx + 1
            Next outItem

            If hasSonotaRow Then
                With wsTarget.Cells(rowIdx, 1)
                    .Value = "その他"
                    .ShrinkToFit = True
                End With

                colOffset = 2
                Dim sonotaTotal As Double
                sonotaTotal = 0
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
                    sonotaTotal = sonotaTotal + sonotaSum
                    colOffset = colOffset + 1
                Next grp
                wsTarget.Cells(rowIdx, colOffset).Value = sonotaTotal
                rowIdx = rowIdx + 1
            End If
        End If

        Dim lastCol As Long
        lastCol = UBound(allGroups) + 3

        Dim tableRange As Range
        On Error Resume Next
        Set tableRange = wsTarget.Range(wsTarget.Cells(outputStartRow, 1), wsTarget.Cells(rowIdx - 1, lastCol))
        On Error GoTo ErrorHandler

        If Not tableRange Is Nothing Then
            Dim baseName As String, tryName As String, tryIdx As Long
            baseName = "_流出G_加工_" & Replace(periodName, " ", "_")
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
            newTable.TableStyle = "TableStyleLight16"
            newTable.ShowAutoFilter = False
            On Error GoTo ErrorHandler

            Dim cIdx As Long
            For cIdx = 1 To newTable.Range.Columns.Count
                newTable.Range.Columns(cIdx).ColumnWidth = 8
            Next cIdx
        End If

        printRangeEnd = rowIdx - 1
        currentRow = rowIdx + 2

NextPeriod:
    Next periodIdx

    If printRangeStart > 0 And printRangeEnd > 0 Then
        Dim printLastCol As Long
        printLastCol = UBound(allGroups) + 3

        On Error Resume Next
        wsTarget.PageSetup.PrintArea = wsTarget.Range( _
            wsTarget.Cells(printRangeStart, 1), _
            wsTarget.Cells(printRangeEnd, printLastCol)).Address
        On Error GoTo ErrorHandler

        Application.StatusBar = "印刷範囲を設定しました"
    End If

Cleanup:
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.DisplayAlerts = origDisplayAlerts
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.DisplayAlerts = origDisplayAlerts
    Application.StatusBar = False

    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "転記_流出_加工G"
End Sub

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

Private Sub QuickSortDesc(ByRef arr() As Variant, ByVal left As Long, ByVal right As Long)
    Dim i As Long, j As Long, pivot As Double
    Dim tempName As String, tempValue As Double

    If left >= right Then Exit Sub

    i = left
    j = right
    pivot = CDbl(arr((left + right) \ 2, 2))

    Do While i <= j
        Do While CDbl(arr(i, 2)) > pivot
            i = i + 1
        Loop
        Do While CDbl(arr(j, 2)) < pivot
            j = j - 1
        Loop
        If i <= j Then
            tempName = arr(i, 1)
            tempValue = arr(i, 2)
            arr(i, 1) = arr(j, 1)
            arr(i, 2) = arr(j, 2)
            arr(j, 1) = tempName
            arr(j, 2) = tempValue
            i = i + 1
            j = j - 1
        End If
    Loop

    If left < j Then Call QuickSortDesc(arr, left, j)
    If i < right Then Call QuickSortDesc(arr, i, right)
End Sub
