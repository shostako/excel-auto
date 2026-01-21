Attribute VB_Name = "m転記_廃棄_加工H"
Option Explicit

' ========================================
' マクロ名: 転記_廃棄_加工H
' 処理概要: 廃棄データを期間別に9分類で集計して加工Hシートに転記（合計列付き）
'
' 【処理の特徴】
' 1. 空白期間対応：集計期間テーブルの空白行（日付シリアル値0）をスキップ
' 2. 動的期間対応：集計期間テーブルの行数が変わっても自動的に対応（増減どちらもOK）
' 3. 高速化：配列処理による大量データの高速集計
' 4. ワースト順機能：項目テーブルの「ワースト」設定に応じて動的に出力順序を変更
' 5. 合計列追加：9品番の右に合計列を追加して各行の合計値を自動計算
'
' 【テーブル構成】
' 項目テーブル : シート「加工H」、テーブル「_廃棄項目加工H」
' 期間テーブル : シート「加工H」、テーブル「_集計期間加工H」
' ソーステーブル : シート「廃棄」、テーブル「_廃棄」
' 出力テーブル : シート「加工H」、複数テーブル「_廃棄H_加工_{期間名}」
'
' 【処理フロー】
' 1. 既存出力テーブルとデータを完全削除
' 2. ワースト設定（全項目 or 数値N）を読み込み
' 3. 各期間ごとに日付フィルター + 品番による9分類集計
' 4. 集計結果を降順ソートしてワースト順出力
' 5. 全期間でテーブル出力（データがなくても構造は出力）
' 6. 各行に9品番の合計列を追加
'
' 【出力形式】
' - 1行目：不良数（「_廃棄」テーブルの「件数」列を品番2で集計）+ 合計
' - 2行目以降：ワースト順で項目別集計 + 合計
'   - 「全項目」設定：0でない項目を降順で全て出力
'   - 数値N設定：上位N件 + 「その他」行（N+1行、ただし0でない項目数<=Nなら「その他」なし）
'
' 【集計条件】
' - 「_廃棄」テーブル：工程=加工
' - 品番照合：完全一致
' ========================================

Sub 転記_廃棄_加工H()
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
    Application.StatusBar = "廃棄加工H転記処理を開始します..."

    Dim wsSource As Worksheet, wsTarget As Worksheet
    Set wsSource = ThisWorkbook.Worksheets("廃棄")
    Set wsTarget = ThisWorkbook.Worksheets("加工H")

    Dim tblSource As ListObject, tblItems As ListObject, tblPeriod As ListObject
    On Error Resume Next
    Set tblSource = wsSource.ListObjects("_廃棄")
    Set tblItems = wsTarget.ListObjects("_廃棄項目加工H")
    Set tblPeriod = wsTarget.ListObjects("_集計期間加工H")
    On Error GoTo ErrorHandler

    If tblSource Is Nothing Then
        MsgBox "シート「廃棄」にテーブル「_廃棄」が見つかりません。", vbCritical
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

    ' 期間テーブルの読み込み（空白行スキップ対応）
    Dim periodCount As Long, periodInfo() As Variant
    periodCount = 0

    If Not tblPeriod Is Nothing Then
        If Not tblPeriod.DataBodyRange Is Nothing Then
            Dim totalRows As Long
            totalRows = tblPeriod.DataBodyRange.Rows.Count

            periodCount = 0
            Dim p As Long
            For p = 1 To totalRows
                Dim startVal As Variant, endVal As Variant
                startVal = tblPeriod.DataBodyRange.Cells(p, 2).Value
                endVal = tblPeriod.DataBodyRange.Cells(p, 3).Value

                If Not IsEmpty(startVal) And Not IsEmpty(endVal) Then
                    If IsDate(startVal) And IsDate(endVal) Then
                        If CDbl(startVal) <> 0 And CDbl(endVal) <> 0 Then
                            periodCount = periodCount + 1
                        End If
                    End If
                End If
            Next p

            If periodCount > 0 Then
                ReDim periodInfo(1 To periodCount, 1 To 3)
                Dim pIdx As Long
                pIdx = 0
                For p = 1 To totalRows
                    startVal = tblPeriod.DataBodyRange.Cells(p, 2).Value
                    endVal = tblPeriod.DataBodyRange.Cells(p, 3).Value

                    If Not IsEmpty(startVal) And Not IsEmpty(endVal) Then
                        If IsDate(startVal) And IsDate(endVal) Then
                            If CDbl(startVal) <> 0 And CDbl(endVal) <> 0 Then
                                pIdx = pIdx + 1
                                periodInfo(pIdx, 1) = CStr(tblPeriod.DataBodyRange.Cells(p, 1).Value)
                                periodInfo(pIdx, 2) = startVal
                                periodInfo(pIdx, 3) = endVal
                            End If
                        End If
                    End If
                Next p
            End If
        End If
    End If

    If periodCount = 0 Then
        MsgBox "「_集計期間加工H」に有効な集計期間がありません。処理を中止します。", vbExclamation
        GoTo Cleanup
    End If

    Dim srcData As Range
    Set srcData = tblSource.DataBodyRange
    If srcData Is Nothing Then
        Application.StatusBar = "ソーステーブルにデータがありません"
        GoTo Cleanup
    End If

    Dim colHizuke As Long, colHinban2 As Long, colKoutei As Long
    Dim colFuryou As Long, colKensuu As Long
    colHizuke = tblSource.ListColumns("日付").Index
    colHinban2 = tblSource.ListColumns("品番2").Index
    colKoutei = tblSource.ListColumns("工程").Index
    colFuryou = tblSource.ListColumns("不良内容").Index
    colKensuu = tblSource.ListColumns("件数").Index

    ' 9品番定義
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
        If loTemp.Name Like "_廃棄H_加工_*" Then
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

    Dim srcArr As Variant
    srcArr = srcData.Value

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

        Dim aggFuryo As Object, aggItems As Object
        Set aggFuryo = CreateObject("Scripting.Dictionary")
        Set aggItems = CreateObject("Scripting.Dictionary")

        Dim grp As Variant
        For Each grp In allGroups
            aggFuryo(CStr(grp)) = 0
            Set aggItems(CStr(grp)) = CreateObject("Scripting.Dictionary")
        Next grp

        Dim r As Long
        For r = 1 To UBound(srcArr, 1)
            Dim srcDate As Variant
            srcDate = srcArr(r, colHizuke)

            If IsDate(srcDate) Then
                Dim srcDt As Date
                srcDt = CDate(srcDate)

                If srcDt >= startDate And srcDt <= endDate Then
                    Dim srcKoutei As String
                    srcKoutei = Trim(CStr(srcArr(r, colKoutei)))

                    If srcKoutei = "加工" Then
                        Dim srcHinban2 As String
                        srcHinban2 = Trim(CStr(srcArr(r, colHinban2)))

                        If hinbanList.Exists(srcHinban2) Then
                            Dim kensuu As Double
                            If IsNumeric(srcArr(r, colKensuu)) Then
                                kensuu = CDbl(srcArr(r, colKensuu))

                                aggFuryo(srcHinban2) = aggFuryo(srcHinban2) + kensuu

                                Dim furyou As String
                                furyou = Trim(CStr(srcArr(r, colFuryou)))

                                If Len(furyou) = 0 Then furyou = "（空白）"
                                If Not aggItems(srcHinban2).Exists(furyou) Then
                                    aggItems(srcHinban2)(furyou) = 0
                                End If
                                aggItems(srcHinban2)(furyou) = aggItems(srcHinban2)(furyou) + kensuu
                            End If
                        End If
                    End If
                End If
            End If

            If (r Mod 200) = 0 Then
                Application.StatusBar = "期間 " & periodIdx & "/" & periodCount & " - " & r & "/" & UBound(srcArr, 1) & " 行処理中..."
            End If
        Next r

        If printRangeStart = 0 Then
            printRangeStart = currentRow
        End If

        Dim titleText As String
        titleText = "加工_廃棄のみ_" & Format(startDate, "m/d") & "‾" & Format(endDate, "m/d")

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

        ' 1行目：不良数
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
            baseName = "_廃棄H_加工_" & Replace(periodName, " ", "_")
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
           "エラー番号: " & Err.Number, vbCritical, "転記_廃棄_加工H"
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
