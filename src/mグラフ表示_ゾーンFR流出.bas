Attribute VB_Name = "mグラフ表示_ゾーンFR流出"
Option Explicit

Sub グラフ表示_ゾーンFR流出()
    ' ========================================
    ' マクロ名: グラフ表示_ゾーンFR流出
    ' 処理概要: ピボットテーブルのフィルタ設定を行い、ゾーンFR流出グラフの表示/非表示を制御
    ' 対象シート: ゾーンFrRr流出
    ' 最適化: PivotFilters.Addで日付フィルタを高速化（PivotItemループ廃止）
    ' ========================================

    Dim ws As Worksheet
    Dim pt1 As PivotTable, pt2 As PivotTable, pt3 As PivotTable, pt4 As PivotTable, pt5 As PivotTable
    Dim dtStart As Date, dtEnd As Date
    Dim occurrenceValue As String
    Dim discovery2Value As String
    Dim discovery2Dict As Object
    Dim arrDiscovery2 As Variant
    Dim isProcessing As Boolean
    Dim isMould As Boolean
    Dim isDiscovery2Empty As Boolean
    Dim commentText As String
    Dim i As Long

    On Error GoTo ErrorHandler

    ' ============================================
    ' 高速化設定
    ' ============================================
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "処理を開始しています..."

    ' ============================================
    ' ワークシート・ピボットテーブル取得
    ' ============================================
    Set ws = ThisWorkbook.Worksheets("ゾーンFrRr流出")
    If ws Is Nothing Then
        MsgBox "指定されたワークシート 'ゾーンFrRr流出' が見つかりません。", vbExclamation
        GoTo Cleanup
    End If

    Application.StatusBar = "ピボットテーブルを確認しています..."
    Set pt1 = ws.PivotTables("ピボットテーブル31")
    Set pt2 = ws.PivotTables("ピボットテーブル32")
    Set pt3 = ws.PivotTables("ピボットテーブル33")
    Set pt4 = ws.PivotTables("ピボットテーブル34")
    Set pt5 = ws.PivotTables("ピボットテーブル35")

    ' ============================================
    ' パラメータ取得・検証
    ' ============================================
    If IsDate(ws.Range("E1").Value) And IsDate(ws.Range("E2").Value) Then
        dtStart = ws.Range("E1").Value
        dtEnd = ws.Range("E2").Value
    Else
        MsgBox "日付範囲が正しく設定されていません。セルE1とE2を確認してください。", vbExclamation
        GoTo Cleanup
    End If

    occurrenceValue = Trim(CStr(ws.Range("E3").Value))
    discovery2Value = Trim(CStr(ws.Range("E4").Value))

    If occurrenceValue = "" Then
        MsgBox "発生の値が設定されていません。セルE3を確認してください。", vbExclamation
        GoTo Cleanup
    End If

    isDiscovery2Empty = (discovery2Value = "")
    isProcessing = (occurrenceValue = "加工")
    isMould = (occurrenceValue = "モール")

    ' 発見2値をDictionaryで高速化
    Set discovery2Dict = CreateObject("Scripting.Dictionary")
    If Not isDiscovery2Empty Then
        arrDiscovery2 = Split(discovery2Value, ",")
        For i = LBound(arrDiscovery2) To UBound(arrDiscovery2)
            discovery2Dict(Trim(arrDiscovery2(i))) = True
        Next i
    End If

    ' ============================================
    ' モード2フィルタリセット
    ' ============================================
    Application.StatusBar = "モード2フィルタをリセット中..."
    Call ResetMode2Filters(Array(pt1, pt2, pt3, pt4, pt5))

    ' ============================================
    ' ピボットテーブルフィルタ設定（高速版）
    ' ============================================
    Application.StatusBar = "アルヴェル Fr ピボットテーブルを設定中..."
    Call FilterPivotTableFast(pt1, dtStart, dtEnd, "アルヴェル", "Fr", occurrenceValue, discovery2Dict, isDiscovery2Empty)

    Application.StatusBar = "アルヴェル Rr ピボットテーブルを設定中..."
    Call FilterPivotTableFast(pt2, dtStart, dtEnd, "アルヴェル", "Rr", occurrenceValue, discovery2Dict, isDiscovery2Empty)

    Application.StatusBar = "ノアヴォク Fr ピボットテーブルを設定中..."
    Call FilterPivotTableFast(pt3, dtStart, dtEnd, "ノアヴォク", "Fr", occurrenceValue, discovery2Dict, isDiscovery2Empty)

    Application.StatusBar = "ノアヴォク Rr ピボットテーブルを設定中..."
    Call FilterPivotTableFast(pt4, dtStart, dtEnd, "ノアヴォク", "Rr", occurrenceValue, discovery2Dict, isDiscovery2Empty)

    Application.StatusBar = "モード抽出用ピボットテーブルを設定中..."
    Call FilterPivotTableForModeFast(pt5, dtStart, dtEnd, occurrenceValue, discovery2Dict, isDiscovery2Empty)

    ' ============================================
    ' グラフ表示設定
    ' ============================================
    Application.StatusBar = "グラフ表示設定を適用中..."
    Dim showGraph1 As Boolean, showGraph2 As Boolean
    Dim showGraph3 As Boolean, showGraph4 As Boolean
    Dim startDateStr As String, endDateStr As String

    Select Case True
        Case isProcessing
            showGraph1 = False
            showGraph2 = False
            showGraph3 = False
            showGraph4 = False
            commentText = "発生が「加工」のため、グラフは表示されません。"

        Case isMould
            showGraph1 = True
            showGraph2 = True
            showGraph3 = False
            showGraph4 = False
            startDateStr = Format(dtStart, "m/d")
            endDateStr = Format(dtEnd, "m/d")
            commentText = occurrenceValue & " 流出不良集計 " & startDateStr & " 〜 " & endDateStr

        Case Else
            showGraph1 = True
            showGraph2 = True
            showGraph3 = True
            showGraph4 = True
            startDateStr = Format(dtStart, "m/d")
            endDateStr = Format(dtEnd, "m/d")
            commentText = occurrenceValue & " 流出不良集計 " & startDateStr & " 〜 " & endDateStr
    End Select

    Call SetChartVisibilityBatch(ws, Array("グラフ1", "グラフ2", "グラフ3", "グラフ4"), _
                                    Array(showGraph1, showGraph2, showGraph3, showGraph4))

    ' ============================================
    ' グラフ軸の動的調整
    ' ============================================
    Application.StatusBar = "グラフ軸を調整中..."
    Dim maxValues() As Double
    ReDim maxValues(1 To 4)

    maxValues(1) = GetPivotTableMaxValueFast(pt1)
    maxValues(2) = GetPivotTableMaxValueFast(pt2)
    maxValues(3) = GetPivotTableMaxValueFast(pt3)
    maxValues(4) = GetPivotTableMaxValueFast(pt4)

    Dim overallMax As Double
    overallMax = Application.WorksheetFunction.Max(maxValues)

    Dim axisMax As Double
    axisMax = GetNiceMaxValue(overallMax)

    Dim tickInterval As Double
    tickInterval = GetNiceTickInterval(axisMax)

    If showGraph1 Then SetChartAxisSettings ws, "グラフ1", axisMax, tickInterval
    If showGraph2 Then SetChartAxisSettings ws, "グラフ2", axisMax, tickInterval
    If showGraph3 Then SetChartAxisSettings ws, "グラフ3", axisMax, tickInterval
    If showGraph4 Then SetChartAxisSettings ws, "グラフ4", axisMax, tickInterval

    ' ============================================
    ' コメント設定・モード入力規則
    ' ============================================
    Application.StatusBar = "最終設定を適用中..."
    With ws.Range("D6")
        .Value = commentText
        .Font.Name = "Yu Gothic UI"
        .Font.Size = 11
        .Font.Bold = True
    End With

    Call SetupModeValidation(ws, pt5)

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

    Set ws = Nothing
    Set pt1 = Nothing
    Set pt2 = Nothing
    Set pt3 = Nothing
    Set pt4 = Nothing
    Set pt5 = Nothing
    Set discovery2Dict = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical, "マクロエラー"
    Resume Cleanup
End Sub

Private Sub FilterPivotTableFast(ByVal pt As PivotTable, _
                                 ByVal startDate As Date, _
                                 ByVal endDate As Date, _
                                 ByVal alNoahFilter As String, _
                                 ByVal frRrFilter As String, _
                                 ByVal occurrenceFilter As String, _
                                 ByVal discovery2Dict As Object, _
                                 ByVal isDiscovery2Empty As Boolean)
    ' ============================================
    ' ピボットテーブルフィルタ設定
    ' xlCaptionBetween でテキストベースの日付フィルタ
    ' ============================================

    Dim pf As PivotField
    Dim pi As PivotItem
    Dim d As Date

    On Error Resume Next

    ' ページフィールド（単一選択）
    pt.PivotFields("アル/ノア").CurrentPage = alNoahFilter
    pt.PivotFields("Fr/Rr").CurrentPage = frRrFilter
    pt.PivotFields("発生").CurrentPage = occurrenceFilter

    ' 日付フィルタ（ページフィールド用、複数選択）
    Set pf = pt.PivotFields("日付")

    pt.ManualUpdate = True  ' 再計算抑制

    pf.ClearAllFilters
    pf.EnableMultiplePageItems = True

    ' パス1: 範囲内を表示
    For Each pi In pf.PivotItems
        If IsDate(pi.Name) Then
            d = CDate(pi.Name)
            If d >= startDate And d <= endDate Then
                pi.Visible = True
            End If
        End If
    Next pi

    ' パス2: 範囲外を非表示
    For Each pi In pf.PivotItems
        If IsDate(pi.Name) Then
            d = CDate(pi.Name)
            If d < startDate Or d > endDate Then
                pi.Visible = False
            End If
        Else
            pi.Visible = False
        End If
    Next pi

    pt.ManualUpdate = False  ' 一括更新

    ' 発見2フィルタ
    If Not isDiscovery2Empty Then
        With pt.PivotFields("発見2")
            .ClearAllFilters
            For Each pi In .PivotItems
                pi.Visible = discovery2Dict.Exists(pi.Name)
            Next pi
        End With
    Else
        pt.PivotFields("発見2").ClearAllFilters
    End If

    On Error GoTo 0
End Sub

Private Sub FilterPivotTableForModeFast(ByVal pt As PivotTable, _
                                        ByVal startDate As Date, _
                                        ByVal endDate As Date, _
                                        ByVal occurrenceFilter As String, _
                                        ByVal discovery2Dict As Object, _
                                        ByVal isDiscovery2Empty As Boolean)
    ' ============================================
    ' ピボットテーブル35（モード抽出用）フィルタ設定
    ' ページフィールド用の複数選択方式
    ' ============================================

    Dim pf As PivotField
    Dim pi As PivotItem
    Dim d As Date

    On Error Resume Next

    ' 日付フィルタ（ページフィールド用、複数選択）
    Set pf = pt.PivotFields("日付")

    pt.ManualUpdate = True  ' 再計算抑制

    pf.ClearAllFilters
    pf.EnableMultiplePageItems = True

    ' パス1: 範囲内を表示
    For Each pi In pf.PivotItems
        If IsDate(pi.Name) Then
            d = CDate(pi.Name)
            If d >= startDate And d <= endDate Then
                pi.Visible = True
            End If
        End If
    Next pi

    ' パス2: 範囲外を非表示
    For Each pi In pf.PivotItems
        If IsDate(pi.Name) Then
            d = CDate(pi.Name)
            If d < startDate Or d > endDate Then
                pi.Visible = False
            End If
        Else
            pi.Visible = False
        End If
    Next pi

    pt.ManualUpdate = False  ' 一括更新

    ' アル/ノア・Fr/Rr：全て表示
    pt.PivotFields("アル/ノア").ClearAllFilters
    pt.PivotFields("Fr/Rr").ClearAllFilters

    ' 発生フィールド
    pt.PivotFields("発生").CurrentPage = occurrenceFilter

    ' 発見2フィルタ
    If Not isDiscovery2Empty Then
        With pt.PivotFields("発見2")
            .ClearAllFilters
            For Each pi In .PivotItems
                pi.Visible = discovery2Dict.Exists(pi.Name)
            Next pi
        End With
    Else
        pt.PivotFields("発見2").ClearAllFilters
    End If

    On Error GoTo 0
End Sub

Private Sub SetChartVisibilityBatch(ByVal ws As Worksheet, ByVal chartNames As Variant, ByVal visibilities As Variant)
    Dim i As Long
    Dim chObj As ChartObject

    On Error Resume Next
    For i = LBound(chartNames) To UBound(chartNames)
        Set chObj = ws.ChartObjects(chartNames(i))
        If Not chObj Is Nothing Then
            chObj.Visible = visibilities(i)
        End If
        Set chObj = Nothing
    Next i
    On Error GoTo 0
End Sub

Private Function GetPivotTableMaxValueFast(ByVal pt As PivotTable) As Double
    Dim dataRange As Range

    On Error Resume Next
    Set dataRange = pt.DataBodyRange

    If dataRange Is Nothing Then
        GetPivotTableMaxValueFast = 0
        Exit Function
    End If

    GetPivotTableMaxValueFast = Application.WorksheetFunction.Max(dataRange)
    On Error GoTo 0
End Function

Private Function GetNiceMaxValue(ByVal maxValue As Double) As Double
    If maxValue <= 0 Then
        GetNiceMaxValue = 10
        Exit Function
    End If

    Dim minTarget As Double, maxTarget As Double
    minTarget = maxValue * 1.1
    maxTarget = maxValue * 1.2

    Dim magnitude As Long
    magnitude = Int(Log(maxTarget) / Log(10))
    Dim base As Double
    base = 10 ^ magnitude

    Dim candidates As Variant
    candidates = Array(1, 1.2, 1.5, 2, 2.5, 3, 4, 5, 6, 7, 8, 9, 10)

    Dim i As Long
    Dim niceValue As Double

    For i = LBound(candidates) To UBound(candidates)
        niceValue = candidates(i) * base
        If niceValue >= minTarget Then
            GetNiceMaxValue = niceValue
            Exit Function
        End If
    Next i

    GetNiceMaxValue = maxTarget
End Function

Private Function GetNiceTickInterval(ByVal maxValue As Double) As Double
    Dim targetTicks As Long
    targetTicks = 6
    Dim roughInterval As Double
    roughInterval = maxValue / targetTicks

    Dim magnitude As Long
    magnitude = Int(Log(roughInterval) / Log(10))
    Dim base As Double
    base = 10 ^ magnitude

    Select Case roughInterval / base
        Case Is <= 1: GetNiceTickInterval = base
        Case Is <= 2: GetNiceTickInterval = 2 * base
        Case Is <= 5: GetNiceTickInterval = 5 * base
        Case Else: GetNiceTickInterval = 10 * base
    End Select
End Function

Private Sub SetChartAxisSettings(ByVal ws As Worksheet, ByVal chartName As String, ByVal maxValue As Double, ByVal tickInterval As Double)
    Dim chObj As ChartObject
    Dim ch As Chart

    On Error Resume Next
    Set chObj = ws.ChartObjects(chartName)

    If Not chObj Is Nothing Then
        Set ch = chObj.Chart
        With ch.Axes(xlValue)
            .MaximumScaleIsAuto = False
            .MaximumScale = maxValue
            .MinimumScaleIsAuto = False
            .MinimumScale = 0
            .MajorUnitIsAuto = False
            .MajorUnit = tickInterval
            .MinorUnitIsAuto = False
            .MinorUnit = tickInterval / 2
        End With
    End If

    Set chObj = Nothing
    Set ch = Nothing
    On Error GoTo 0
End Sub

Private Sub ResetMode2Filters(ByVal pivotTables As Variant)
    Dim pt As PivotTable
    Dim i As Long

    On Error Resume Next
    For i = LBound(pivotTables) To UBound(pivotTables)
        Set pt = pivotTables(i)
        With pt.PivotFields("モード2")
            .ClearAllFilters
            .CurrentPage = "(すべて)"
        End With
    Next i
    On Error GoTo 0
End Sub

Private Sub SetupModeValidation(ByVal ws As Worksheet, ByVal pt5 As PivotTable)
    Dim modeItems As Object
    Dim pi As PivotItem
    Dim cellValue As String
    Dim modeList As String

    Set modeItems = CreateObject("Scripting.Dictionary")

    Dim excludeDict As Object
    Set excludeDict = CreateObject("Scripting.Dictionary")
    excludeDict.Add "A", True
    excludeDict.Add "B", True
    excludeDict.Add "C", True
    excludeDict.Add "D", True
    excludeDict.Add "E", True
    excludeDict.Add "Fr RH", True

    On Error Resume Next
    For Each pi In pt5.PivotFields("モード").PivotItems
        cellValue = Trim(pi.Name)
        If cellValue <> "" And Not excludeDict.Exists(cellValue) And Not modeItems.Exists(cellValue) Then
            modeItems.Add cellValue, cellValue
        End If
    Next pi
    On Error GoTo 0

    If modeItems.Count > 0 Then
        modeList = Join(modeItems.Keys, ",")

        With ws.Range("T3")
            .Validation.Delete
            .Value = ""
            .Validation.Add Type:=xlValidateList, _
                           AlertStyle:=xlValidAlertStop, _
                           Formula1:=modeList
            .Value = ""
        End With
    Else
        ws.Range("T3").Validation.Delete
        ws.Range("T3").Value = "モード項目なし"
    End If

    Set modeItems = Nothing
    Set excludeDict = Nothing
End Sub
