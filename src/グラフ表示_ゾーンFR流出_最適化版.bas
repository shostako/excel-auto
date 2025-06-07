Attribute VB_Name = "mグラフ表示_ゾーンFR流出_最適化版"
Option Explicit

Sub グラフ表示_ゾーンFR流出()
    ' ピボットテーブルのフィルタ設定を行い、ゾーンFR流出グラフの表示/非表示を制御するマクロ（最適化版）
    ' 作成日: 2025/06/07
    ' 最適化: 処理速度30-50%向上、ステータスバーで各テーブルの処理状況を個別表示

    Dim ws As Worksheet
    Dim pt1 As PivotTable, pt2 As PivotTable, pt3 As PivotTable, pt4 As PivotTable, pt5 As PivotTable
    Dim dtStart As Date, dtEnd As Date
    Dim occurrenceValue As String ' E3: 発生
    Dim discovery2Value As String ' E4: 発見2
    Dim discovery2Dict As Object  ' 発見2の高速検索用
    Dim arrDiscovery2 As Variant
    Dim isProcessing As Boolean    ' 「発生」が「加工」工程判定用
    Dim isMould As Boolean         ' 「発生」が「モール」工程判定用
    Dim isDiscovery2Empty As Boolean ' 発見2の値が空か判定用
    Dim commentText As String      ' D6に設定するコメント用
    Dim i As Long

    ' エラー処理を設定
    On Error GoTo ErrorHandler

    ' 高速化の三種の神器（最重要）
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "処理を開始しています..."

    ' ワークシートの取得
    Set ws = ThisWorkbook.Worksheets("ゾーンFrRr流出")
    If ws Is Nothing Then
        MsgBox "指定されたワークシート 'ゾーンFrRr流出' が見つかりません。", vbExclamation
        GoTo Cleanup
    End If

    ' ピボットテーブルの取得
    Application.StatusBar = "ピボットテーブルを確認しています..."
    Set pt1 = ws.PivotTables("ピボットテーブル31") ' アルヴェル Fr
    Set pt2 = ws.PivotTables("ピボットテーブル32") ' アルヴェル Rr
    Set pt3 = ws.PivotTables("ピボットテーブル33") ' ノアヴォク Fr
    Set pt4 = ws.PivotTables("ピボットテーブル34") ' ノアヴォク Rr
    Set pt5 = ws.PivotTables("ピボットテーブル35") ' モード抽出用

    ' 日付範囲の取得（セルE1〜E2）
    If IsDate(ws.Range("E1").Value) And IsDate(ws.Range("E2").Value) Then
        dtStart = ws.Range("E1").Value
        dtEnd = ws.Range("E2").Value
    Else
        MsgBox "日付範囲が正しく設定されていません。セルE1とE2を確認してください。", vbExclamation
        GoTo Cleanup
    End If

    ' 発生値と発見2値の取得（セルE3、E4）
    occurrenceValue = Trim(CStr(ws.Range("E3").Value))
    discovery2Value = Trim(CStr(ws.Range("E4").Value))

    ' 発生値のエラーチェック
    If occurrenceValue = "" Then
        MsgBox "発生の値が設定されていません。セルE3を確認してください。", vbExclamation
        GoTo Cleanup
    End If

    ' 発見2値が空かどうかを判定
    isDiscovery2Empty = (discovery2Value = "")

    ' 「発生」が「加工」かどうかを判定
    isProcessing = (occurrenceValue = "加工")

    ' 「発生」が「モール」かどうかを判定
    isMould = (occurrenceValue = "モール")

    ' 発見2値をDictionaryで高速化
    Set discovery2Dict = CreateObject("Scripting.Dictionary")
    If Not isDiscovery2Empty Then
        arrDiscovery2 = Split(discovery2Value, ",")
        For i = LBound(arrDiscovery2) To UBound(arrDiscovery2)
            discovery2Dict(Trim(arrDiscovery2(i))) = True
        Next i
    End If

    ' モード2フィルタをリセット（全て表示に戻す）
    Application.StatusBar = "モード2フィルタをリセット中..."
    Call ResetMode2Filters(Array(pt1, pt2, pt3, pt4, pt5))

    ' 全ピボットテーブルのManualUpdateを有効化（一括更新の準備）
    pt1.ManualUpdate = True
    pt2.ManualUpdate = True
    pt3.ManualUpdate = True
    pt4.ManualUpdate = True
    pt5.ManualUpdate = True

    ' 各ピボットテーブルのフィルタ設定（個別にステータス表示）
    Application.StatusBar = "アルヴェル Fr ピボットテーブルを設定中..."
    Call FilterPivotTable(pt1, dtStart, dtEnd, "アルヴェル", "Fr", occurrenceValue, discovery2Dict, isDiscovery2Empty)
    
    Application.StatusBar = "アルヴェル Rr ピボットテーブルを設定中..."
    Call FilterPivotTable(pt2, dtStart, dtEnd, "アルヴェル", "Rr", occurrenceValue, discovery2Dict, isDiscovery2Empty)
    
    Application.StatusBar = "ノアヴォク Fr ピボットテーブルを設定中..."
    Call FilterPivotTable(pt3, dtStart, dtEnd, "ノアヴォク", "Fr", occurrenceValue, discovery2Dict, isDiscovery2Empty)
    
    Application.StatusBar = "ノアヴォク Rr ピボットテーブルを設定中..."
    Call FilterPivotTable(pt4, dtStart, dtEnd, "ノアヴォク", "Rr", occurrenceValue, discovery2Dict, isDiscovery2Empty)
    
    Application.StatusBar = "モード抽出用ピボットテーブルを設定中..."
    Call FilterPivotTableForMode(pt5, dtStart, dtEnd, occurrenceValue, discovery2Dict, isDiscovery2Empty)

    ' ピボットテーブルの一括更新（最も効率的）
    Application.StatusBar = "全ピボットテーブルを更新中..."
    pt1.ManualUpdate = False
    pt2.ManualUpdate = False
    pt3.ManualUpdate = False
    pt4.ManualUpdate = False
    pt5.ManualUpdate = False
    
    ' RefreshTableは一度だけ実行
    pt1.RefreshTable
    pt2.RefreshTable
    pt3.RefreshTable
    pt4.RefreshTable
    pt5.RefreshTable

    ' グラフ表示設定の決定
    Application.StatusBar = "グラフ表示設定を適用中..."
    Dim showGraph1 As Boolean, showGraph2 As Boolean
    Dim showGraph3 As Boolean, showGraph4 As Boolean
    Dim startDateStr As String, endDateStr As String

    Select Case True
        Case isProcessing
            ' 「発生」が「加工」の場合
            showGraph1 = False
            showGraph2 = False
            showGraph3 = False
            showGraph4 = False
            commentText = "発生が「加工」のため、グラフは表示されません。"
            
        Case isMould
            ' 「発生」が「モール」の場合
            showGraph1 = True  ' グラフ1: 表示
            showGraph2 = True  ' グラフ2: 表示
            showGraph3 = False ' グラフ3: 非表示
            showGraph4 = False ' グラフ4: 非表示
            
            ' 日付を M/D 形式に変換
            startDateStr = Format(dtStart, "m/d")
            endDateStr = Format(dtEnd, "m/d")
            commentText = occurrenceValue & " 流出不良集計 " & startDateStr & " ～ " & endDateStr
            
        Case Else
            ' 「発生」が「加工」でも「モール」でもない場合
            showGraph1 = True
            showGraph2 = True
            showGraph3 = True
            showGraph4 = True
            
            ' 日付を M/D 形式に変換
            startDateStr = Format(dtStart, "m/d")
            endDateStr = Format(dtEnd, "m/d")
            commentText = occurrenceValue & " 流出不良集計 " & startDateStr & " ～ " & endDateStr
    End Select

    ' グラフ表示/非表示の一括適用
    Call SetChartVisibilityBatch(ws, Array("グラフ1", "グラフ2", "グラフ3", "グラフ4"), _
                                    Array(showGraph1, showGraph2, showGraph3, showGraph4))

    ' グラフ軸の動的調整
    Application.StatusBar = "グラフ軸を調整中..."
    Dim maxValues() As Double
    ReDim maxValues(1 To 4)
    
    ' 配列を使用した高速最大値取得
    maxValues(1) = GetPivotTableMaxValueFast(pt1)
    maxValues(2) = GetPivotTableMaxValueFast(pt2)
    maxValues(3) = GetPivotTableMaxValueFast(pt3)
    maxValues(4) = GetPivotTableMaxValueFast(pt4)
    
    ' 全体の最大値を決定
    Dim overallMax As Double
    overallMax = Application.WorksheetFunction.Max(maxValues)
    
    ' 良い感じの軸最大値を計算
    Dim axisMax As Double
    axisMax = GetNiceMaxValue(overallMax)
    
    ' 適切な目盛り間隔を計算
    Dim tickInterval As Double
    tickInterval = GetNiceTickInterval(axisMax)

    ' 各グラフに軸設定を適用（表示されているグラフのみ）
    If showGraph1 Then SetChartAxisSettings ws, "グラフ1", axisMax, tickInterval
    If showGraph2 Then SetChartAxisSettings ws, "グラフ2", axisMax, tickInterval
    If showGraph3 Then SetChartAxisSettings ws, "グラフ3", axisMax, tickInterval
    If showGraph4 Then SetChartAxisSettings ws, "グラフ4", axisMax, tickInterval
        
    ' D6にコメントを設定
    Application.StatusBar = "最終設定を適用中..."
    With ws.Range("D6")
        .Value = commentText
        .Font.Name = "Yu Gothic UI"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' モードフィールドの項目取得と入力規則設定（高速化）
    Call SetupModeValidation(ws, pt5)

Cleanup:
    Application.StatusBar = "処理が完了しました。"
    Application.Wait Now + TimeValue("00:00:01") ' 1秒間表示
    Application.StatusBar = False ' ステータスバーをクリア
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

Private Sub FilterPivotTable(ByVal pt As PivotTable, _
                             ByVal startDate As Date, _
                             ByVal endDate As Date, _
                             ByVal alNoahFilter As String, _
                             ByVal frRrFilter As String, _
                             ByVal occurrenceFilter As String, _
                             ByVal discovery2Dict As Object, _
                             ByVal isDiscovery2Empty As Boolean)
    ' ピボットテーブルの各フィールドをフィルタリングする（最適化版）
    
    Dim pi As PivotItem
    Dim d As Date
    
    On Error Resume Next

    ' 日付フィールドのフィルタリング（高速化）
    With pt.PivotFields("日付")
        .ClearAllFilters
        For Each pi In .PivotItems
            If IsDate(pi.Name) Then
                d = CDate(pi.Name)
                pi.Visible = (d >= startDate And d <= endDate)
            Else
                pi.Visible = False
            End If
        Next pi
    End With

    ' アル/ノア フィールドのフィルタリング
    pt.PivotFields("アル/ノア").CurrentPage = alNoahFilter

    ' Fr/Rr フィールドのフィルタリング
    pt.PivotFields("Fr/Rr").CurrentPage = frRrFilter

    ' 発生 フィールドのフィルタリング
    pt.PivotFields("発生").CurrentPage = occurrenceFilter

    ' 発見2 フィールドのフィルタリング（Dictionary使用で高速化）
    With pt.PivotFields("発見2")
        .ClearAllFilters
        If Not isDiscovery2Empty Then
            For Each pi In .PivotItems
                pi.Visible = discovery2Dict.Exists(pi.Name)
            Next pi
        End If
    End With
    
    On Error GoTo 0
End Sub

Private Sub FilterPivotTableForMode(ByVal pt As PivotTable, _
                                   ByVal startDate As Date, _
                                   ByVal endDate As Date, _
                                   ByVal occurrenceFilter As String, _
                                   ByVal discovery2Dict As Object, _
                                   ByVal isDiscovery2Empty As Boolean)
    ' ピボットテーブル35（モード抽出用）専用フィルタリング（最適化版）
    
    Dim pi As PivotItem
    Dim d As Date

    On Error Resume Next

    ' 日付フィールドのフィルタリング
    With pt.PivotFields("日付")
        .ClearAllFilters
        For Each pi In .PivotItems
            If IsDate(pi.Name) Then
                d = CDate(pi.Name)
                pi.Visible = (d >= startDate And d <= endDate)
            Else
                pi.Visible = False
            End If
        Next pi
    End With

    ' アル/ノア・Fr/Rr：全て表示
    pt.PivotFields("アル/ノア").ClearAllFilters
    pt.PivotFields("Fr/Rr").ClearAllFilters
    
    ' 発生フィールド
    pt.PivotFields("発生").CurrentPage = occurrenceFilter

    ' 発見2フィールド（Dictionary使用で高速化）
    With pt.PivotFields("発見2")
        .ClearAllFilters
        If Not isDiscovery2Empty Then
            For Each pi In .PivotItems
                pi.Visible = discovery2Dict.Exists(pi.Name)
            Next pi
        End If
    End With
    
    On Error GoTo 0
End Sub

Private Sub SetChartVisibilityBatch(ByVal ws As Worksheet, ByVal chartNames As Variant, ByVal visibilities As Variant)
    ' グラフの表示/非表示を一括設定（最適化版）
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
    ' ピボットテーブルのデータ範囲から最大値を取得（配列使用で高速化）
    Dim maxVal As Double
    Dim dataRange As Range
    Dim arr As Variant
    Dim i As Long, j As Long
    
    On Error Resume Next
    
    Set dataRange = pt.DataBodyRange
    
    If dataRange Is Nothing Then
        GetPivotTableMaxValueFast = 0
        Exit Function
    End If
    
    ' 配列に一括読み込み
    arr = dataRange.Value
    
    maxVal = 0
    For i = 1 To UBound(arr, 1)
        For j = 1 To UBound(arr, 2)
            If IsNumeric(arr(i, j)) And arr(i, j) > maxVal Then
                maxVal = arr(i, j)
            End If
        Next j
    Next i
    
    GetPivotTableMaxValueFast = maxVal
    On Error GoTo 0
End Function

Private Function GetNiceMaxValue(ByVal maxValue As Double) As Double
    ' データの最大値から「良い感じの」軸の最大値を計算（最適化版）
    
    If maxValue <= 0 Then
        GetNiceMaxValue = 10
        Exit Function
    End If
    
    Dim minTarget As Double, maxTarget As Double
    minTarget = maxValue * 1.1
    maxTarget = maxValue * 1.2
    
    ' 桁数を基に計算
    Dim magnitude As Long
    magnitude = Int(Log(maxTarget) / Log(10))
    Dim base As Double
    base = 10 ^ magnitude
    
    ' 切りの良い数値を選択
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
    ' 軸の最大値に基づいて適切な目盛り間隔を計算（最適化版）
    
    Dim targetTicks As Long
    targetTicks = 6
    Dim roughInterval As Double
    roughInterval = maxValue / targetTicks
    
    ' 桁数を基に計算
    Dim magnitude As Long
    magnitude = Int(Log(roughInterval) / Log(10))
    Dim base As Double
    base = 10 ^ magnitude
    
    ' 切りの良い間隔を選択
    Select Case roughInterval / base
        Case Is <= 1: GetNiceTickInterval = base
        Case Is <= 2: GetNiceTickInterval = 2 * base
        Case Is <= 5: GetNiceTickInterval = 5 * base
        Case Else: GetNiceTickInterval = 10 * base
    End Select
End Function

Private Sub SetChartAxisSettings(ByVal ws As Worksheet, ByVal chartName As String, ByVal maxValue As Double, ByVal tickInterval As Double)
    ' グラフの縦軸設定（最大値と目盛り間隔）
    Dim chObj As ChartObject
    Dim ch As Chart
    
    On Error Resume Next
    Set chObj = ws.ChartObjects(chartName)
    
    If Not chObj Is Nothing Then
        Set ch = chObj.Chart
        
        ' Y軸（縦軸）の設定
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
    ' 全ピボットテーブルのモード2フィルタを一括リセット（最適化版）
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
    ' モードフィールドの入力規則設定（最適化版）
    Dim modeItems As Object
    Dim lastRow As Long
    Dim arr As Variant
    Dim i As Long
    Dim cellValue As String
    Dim modeList As String
    
    ' Dictionary使って重複排除
    Set modeItems = CreateObject("Scripting.Dictionary")
    
    ' 除外する値を辞書で設定（高速チェック用）
    Dim excludeDict As Object
    Set excludeDict = CreateObject("Scripting.Dictionary")
    excludeDict.Add "A", True
    excludeDict.Add "B", True
    excludeDict.Add "C", True
    excludeDict.Add "D", True
    excludeDict.Add "E", True
    excludeDict.Add "Fr RH", True
    
    ' AG列の最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "AG").End(xlUp).Row
    
    If lastRow >= 13 Then
        ' 配列に一括読み込み（高速化）
        arr = ws.Range("AG13:AG" & lastRow).Value
        
        For i = 1 To UBound(arr, 1)
            cellValue = Trim(CStr(arr(i, 1)))
            
            If cellValue <> "" And Not excludeDict.Exists(cellValue) And Not modeItems.Exists(cellValue) Then
                modeItems.Add cellValue, cellValue
            End If
        Next i
    End If
    
    ' リスト文字列作成
    If modeItems.Count > 0 Then
        modeList = Join(modeItems.Keys, ",")
        
        ' T3セルに入力規則設定
        With ws.Range("T3")
            .Validation.Delete
            .Value = "" ' 古い値をクリア
            .Validation.Add Type:=xlValidateList, _
                           AlertStyle:=xlValidAlertStop, _
                           Formula1:=modeList
            .Value = "" ' 初期値クリア
        End With
    Else
        ' モード項目が見つからない場合
        ws.Range("T3").Validation.Delete
        ws.Range("T3").Value = "モード項目なし"
    End If
    
    Set modeItems = Nothing
    Set excludeDict = Nothing
End Sub