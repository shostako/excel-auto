Attribute VB_Name = "mグラフ軸設定"
Option Explicit

' ========================================
' モジュール名: mグラフ軸設定
' 処理概要: グラフ縦軸の動的調整を行う共通関数群
' 使用箇所: グラフ表示マクロ、コマンドボタンマクロ等
' ========================================

Public Sub ApplyChartAxisFromPivots(ByVal ws As Worksheet, _
                                    ByVal pivotNames As Variant, _
                                    ByVal chartNames As Variant, _
                                    Optional ByVal chartVisibilities As Variant)
    ' ============================================
    ' ピボットテーブルの最大値からグラフ軸を一括設定
    ' pivotNames: ピボットテーブル名の配列
    ' chartNames: グラフ名の配列（pivotNamesと同じ順序で対応）
    ' chartVisibilities: グラフ表示状態の配列（省略時は全て表示扱い）
    ' ============================================

    Dim pt As PivotTable
    Dim maxValues() As Double
    Dim i As Long
    Dim overallMax As Double
    Dim axisMax As Double
    Dim tickInterval As Double
    Dim isVisible As Boolean

    ' 最大値取得
    ReDim maxValues(LBound(pivotNames) To UBound(pivotNames))

    On Error Resume Next
    For i = LBound(pivotNames) To UBound(pivotNames)
        Set pt = ws.PivotTables(pivotNames(i))
        If Not pt Is Nothing Then
            maxValues(i) = GetPivotMaxValue(pt)
        Else
            maxValues(i) = 0
        End If
    Next i
    On Error GoTo 0

    ' 全体の最大値を決定
    overallMax = Application.WorksheetFunction.Max(maxValues)

    ' 良い感じの軸最大値・目盛り間隔を計算
    axisMax = CalcNiceMaxValue(overallMax)
    tickInterval = CalcNiceTickInterval(axisMax)

    ' 各グラフに軸設定を適用
    For i = LBound(chartNames) To UBound(chartNames)
        ' 表示状態の判定
        If IsMissing(chartVisibilities) Then
            isVisible = True
        ElseIf i <= UBound(chartVisibilities) Then
            isVisible = chartVisibilities(i)
        Else
            isVisible = True
        End If

        If isVisible Then
            Call SetChartAxis(ws, CStr(chartNames(i)), axisMax, tickInterval)
        End If
    Next i
End Sub

Public Function GetPivotMaxValue(ByVal pt As PivotTable) As Double
    ' ============================================
    ' ピボットテーブルのデータ範囲から最大値を取得
    ' WorksheetFunction.Maxで高速化
    ' ============================================

    Dim dataRange As Range

    On Error Resume Next
    Set dataRange = pt.DataBodyRange

    If dataRange Is Nothing Then
        GetPivotMaxValue = 0
        Exit Function
    End If

    GetPivotMaxValue = Application.WorksheetFunction.Max(dataRange)
    On Error GoTo 0
End Function

Public Function CalcNiceMaxValue(ByVal maxValue As Double) As Double
    ' ============================================
    ' データの最大値から「良い感じの」軸の最大値を計算
    ' ============================================

    If maxValue <= 0 Then
        CalcNiceMaxValue = 10
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
            CalcNiceMaxValue = niceValue
            Exit Function
        End If
    Next i

    CalcNiceMaxValue = maxTarget
End Function

Public Function CalcNiceTickInterval(ByVal maxValue As Double) As Double
    ' ============================================
    ' 軸の最大値に基づいて適切な目盛り間隔を計算
    ' ============================================

    Dim targetTicks As Long
    targetTicks = 6
    Dim roughInterval As Double
    roughInterval = maxValue / targetTicks

    Dim magnitude As Long
    magnitude = Int(Log(roughInterval) / Log(10))
    Dim base As Double
    base = 10 ^ magnitude

    Select Case roughInterval / base
        Case Is <= 1: CalcNiceTickInterval = base
        Case Is <= 2: CalcNiceTickInterval = 2 * base
        Case Is <= 5: CalcNiceTickInterval = 5 * base
        Case Else: CalcNiceTickInterval = 10 * base
    End Select
End Function

Public Sub SetChartAxis(ByVal ws As Worksheet, _
                        ByVal chartName As String, _
                        ByVal maxValue As Double, _
                        ByVal tickInterval As Double)
    ' ============================================
    ' グラフの縦軸設定（最大値と目盛り間隔）
    ' ============================================

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
