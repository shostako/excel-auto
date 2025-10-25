Attribute VB_Name = "m富士山型ウォーターフォール"
Option Explicit

'==========================================
' 富士山型ウォーターフォール生成マクロ
'
' 入力:
'   - _期間A (工程/流出/廃棄/数量)
'   - _期間B (工程/成形/塗装/数量)
'
' 出力:
'   - シート「富士山_変換」
'   - テーブル「富士山_変換」
'   - 富士山型ウォーターフォールグラフ
'
' グラフ構成:
'   - 左側: 期間A（流出/廃棄の積み上げ）
'   - 中央: 加工流出総数292（単色棒）
'   - 右側: 期間B（成形/塗装の積み上げ）
'==========================================

Public Sub Build富士山型WF()
    On Error GoTo ErrorHandler
    Application.StatusBar = "富士山型ウォーターフォール生成を開始します..."

    ' 最適化設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim wb As Workbook: Set wb = ThisWorkbook

    ' テーブル取得
    Dim loA As ListObject: Set loA = GetListObjectByName(wb, "_期間A")
    Dim loB As ListObject: Set loB = GetListObjectByName(wb, "_期間B")

    If loA Is Nothing Then Err.Raise 5, , "テーブル _期間A が見つかりません"
    If loB Is Nothing Then Err.Raise 5, , "テーブル _期間B が見つかりません"

    ' 必須列チェック
    RequireColumn loA, "工程": RequireColumn loA, "流出": RequireColumn loA, "廃棄"
    RequireColumn loB, "工程": RequireColumn loB, "成形": RequireColumn loB, "塗装"

    Application.StatusBar = "変換テーブルを作成中..."

    ' 出力シート作成
    Dim wsOut As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next: wb.Worksheets("富士山_変換").Delete: On Error GoTo ErrorHandler
    Application.DisplayAlerts = True

    Set wsOut = wb.Worksheets.Add(After:=loA.Parent)
    wsOut.Name = "富士山_変換"

    ' 出力テーブル作成
    Dim loOut As ListObject
    With wsOut
        .Range("A1:H1").Value = Array("工程", "Base", "流出", "廃棄", "成形", "塗装", "単色", "累積")
        Set loOut = .ListObjects.Add(xlSrcRange, .Range("A1:H2"), , xlYes)
        loOut.Name = "富士山_変換"
        If Not loOut.DataBodyRange Is Nothing Then loOut.DataBodyRange.Delete
    End With

    ' 変換処理
    Dim cum As Double: cum = 0
    Dim r As Long, rowOut As ListRow
    Dim nm As String, leak As Double, scrap As Double, molding As Double, painting As Double
    Dim delta As Double, nextCum As Double, baseVal As Double

    ' 期間A処理（成形、次工程発見、塗装）
    Application.StatusBar = "期間Aデータを処理中..."
    For r = 1 To loA.DataBodyRange.Rows.Count
        nm = CStr(loA.ListColumns("工程").DataBodyRange(r, 1).Value)
        leak = ToDbl(loA.ListColumns("流出").DataBodyRange(r, 1).Value)
        scrap = ToDbl(loA.ListColumns("廃棄").DataBodyRange(r, 1).Value)

        ' 数量列があればそれを使用、なければ流出+廃棄
        delta = GetOptionalDbl(loA, "数量", r, leak + scrap)
        nextCum = cum + delta
        baseVal = Min2(cum, nextCum)

        Set rowOut = loOut.ListRows.Add
        rowOut.Range(1, 1).Value = nm           ' 工程
        rowOut.Range(1, 2).Value = baseVal      ' Base
        rowOut.Range(1, 3).Value = Abs(leak)    ' 流出
        rowOut.Range(1, 4).Value = Abs(scrap)   ' 廃棄
        rowOut.Range(1, 5).Value = 0            ' 成形（期間Aでは0）
        rowOut.Range(1, 6).Value = 0            ' 塗装（期間Aでは0）
        rowOut.Range(1, 7).Value = 0            ' 単色（期間Aでは0）
        rowOut.Range(1, 8).Value = nextCum      ' 累積

        cum = nextCum

        ' 加工流出総数に到達したら中央棒を追加
        If InStr(1, nm, "加工流出総数", vbTextCompare) > 0 Then
            Set rowOut = loOut.ListRows.Add
            rowOut.Range(1, 1).Value = "加工流出総数"
            rowOut.Range(1, 2).Value = 0            ' Base=0（地面から）
            rowOut.Range(1, 3).Value = 0            ' 流出
            rowOut.Range(1, 4).Value = 0            ' 廃棄
            rowOut.Range(1, 5).Value = 0            ' 成形
            rowOut.Range(1, 6).Value = 0            ' 塗装
            rowOut.Range(1, 7).Value = Abs(cum)     ' 単色=292
            rowOut.Range(1, 8).Value = Abs(cum)     ' 累積（絶対値）

            ' 期間Bの開始点として累積をリセット
            cum = Abs(cum)
            Exit For
        End If
    Next r

    ' 期間B処理（加工手直し、差戻し、廃棄）
    Application.StatusBar = "期間Bデータを処理中..."
    For r = 1 To loB.DataBodyRange.Rows.Count
        nm = CStr(loB.ListColumns("工程").DataBodyRange(r, 1).Value)

        ' 加工流出総数はスキップ（既に追加済み）
        If InStr(1, nm, "加工流出総数", vbTextCompare) > 0 Then GoTo NextRowB

        molding = ToDbl(loB.ListColumns("成形").DataBodyRange(r, 1).Value)
        painting = ToDbl(loB.ListColumns("塗装").DataBodyRange(r, 1).Value)

        ' 数量列があればそれを使用、なければ成形+塗装
        delta = GetOptionalDbl(loB, "数量", r, molding + painting)
        nextCum = cum + delta
        baseVal = Min2(cum, nextCum)

        Set rowOut = loOut.ListRows.Add
        rowOut.Range(1, 1).Value = nm               ' 工程
        rowOut.Range(1, 2).Value = baseVal          ' Base
        rowOut.Range(1, 3).Value = 0                ' 流出（期間Bでは0）
        rowOut.Range(1, 4).Value = 0                ' 廃棄（期間Bでは0）
        rowOut.Range(1, 5).Value = Abs(molding)     ' 成形
        rowOut.Range(1, 6).Value = Abs(painting)    ' 塗装
        rowOut.Range(1, 7).Value = 0                ' 単色（期間Bでは0）
        rowOut.Range(1, 8).Value = nextCum          ' 累積

        cum = nextCum

NextRowB:
    Next r

    ' グラフ生成
    Application.StatusBar = "グラフを生成中..."
    CreateFujisanChart wsOut, loOut

    ' 最適化設定を戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
End Sub

'==========================================
' グラフ生成処理
'==========================================
Private Sub CreateFujisanChart(ByVal ws As Worksheet, ByVal lo As ListObject)
    Dim chObj As ChartObject, ch As Chart
    Set chObj = ws.ChartObjects.Add( _
        Left:=ws.Range("J2").Left, _
        Top:=ws.Range("J2").Top, _
        Width:=720, _
        Height:=420 _
    )
    Set ch = chObj.Chart
    ch.ChartType = xlColumnStacked

    ' データ範囲設定（Base + 流出 + 廃棄 + 成形 + 塗装 + 単色）
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ch.SetSourceData ws.Range("B1:G" & lastR)
    ch.Axes(xlCategory).CategoryNames = ws.Range("A2:A" & lastR)
    ch.SetElement (msoElementLegendBottom)

    ' Base系列を透明化
    With ch.FullSeriesCollection(1)
        .Name = "Base"
        .Format.Fill.Visible = msoFalse
        .Format.Line.Visible = msoFalse
    End With

    ' タイトル設定
    ch.HasTitle = True
    ch.ChartTitle.Text = "富士山型ウォーターフォール（期間A → 加工流出総数 → 期間B）"
    ch.ChartGroups(1).GapWidth = 50
    ch.Axes(xlValue).HasMajorGridlines = True

    ' 配色設定
    ColorFujisanSeries ch
End Sub

'==========================================
' 配色処理
'==========================================
Private Sub ColorFujisanSeries(ByVal ch As Chart)
    ' 色定義（Tailwind風）
    Dim blueLeakDark As Long: blueLeakDark = RGB(37, 99, 235)      ' 流出（濃青）
    Dim blueScrapLight As Long: blueScrapLight = RGB(147, 197, 253) ' 廃棄（淡青）
    Dim greenMoldDark As Long: greenMoldDark = RGB(34, 197, 94)     ' 成形（濃緑）
    Dim greenPaintLight As Long: greenPaintLight = RGB(134, 239, 172) ' 塗装（淡緑）
    Dim graySingle As Long: graySingle = RGB(156, 163, 175)         ' 単色（グレー）

    ' 系列2: 流出（濃青）
    With ch.FullSeriesCollection(2)
        .Name = "流出"
        .Format.Fill.Visible = msoTrue
        .Format.Fill.Solid
        .Format.Fill.ForeColor.RGB = blueLeakDark
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(255, 255, 255)
        .Format.Line.Weight = 0.75
    End With

    ' 系列3: 廃棄（淡青）
    With ch.FullSeriesCollection(3)
        .Name = "廃棄"
        .Format.Fill.Visible = msoTrue
        .Format.Fill.Solid
        .Format.Fill.ForeColor.RGB = blueScrapLight
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(255, 255, 255)
        .Format.Line.Weight = 0.75
    End With

    ' 系列4: 成形（濃緑）
    With ch.FullSeriesCollection(4)
        .Name = "成形"
        .Format.Fill.Visible = msoTrue
        .Format.Fill.Solid
        .Format.Fill.ForeColor.RGB = greenMoldDark
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(255, 255, 255)
        .Format.Line.Weight = 0.75
    End With

    ' 系列5: 塗装（淡緑）
    With ch.FullSeriesCollection(5)
        .Name = "塗装"
        .Format.Fill.Visible = msoTrue
        .Format.Fill.Solid
        .Format.Fill.ForeColor.RGB = greenPaintLight
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(255, 255, 255)
        .Format.Line.Weight = 0.75
    End With

    ' 系列6: 単色（グレー）
    With ch.FullSeriesCollection(6)
        .Name = "加工流出総数"
        .Format.Fill.Visible = msoTrue
        .Format.Fill.Solid
        .Format.Fill.ForeColor.RGB = graySingle
        .Format.Line.Visible = msoTrue
        .Format.Line.ForeColor.RGB = RGB(255, 255, 255)
        .Format.Line.Weight = 0.75
    End With
End Sub

'==========================================
' ヘルパー関数群
'==========================================

' ワークブック内のListObjectを名前で取得
Private Function GetListObjectByName(ByVal wb As Workbook, ByVal loName As String) As ListObject
    Dim ws As Worksheet, t As ListObject
    For Each ws In wb.Worksheets
        For Each t In ws.ListObjects
            If StrComp(t.Name, loName, vbTextCompare) = 0 Then
                Set GetListObjectByName = t
                Exit Function
            End If
        Next t
    Next ws
End Function

' 必須列チェック
Private Sub RequireColumn(ByVal lo As ListObject, ByVal colName As String)
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then Exit Sub
    Next lc
    Err.Raise 5, , "テーブル「" & lo.Name & "」に列「" & colName & "」がありません"
End Sub

' 任意列の値取得（数値）
Private Function GetOptionalDbl(ByVal lo As ListObject, ByVal colName As String, _
                                ByVal r As Long, Optional ByVal defaultVal As Double = 0) As Double
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then
            GetOptionalDbl = ToDbl(lc.DataBodyRange(r, 1).Value, defaultVal)
            Exit Function
        End If
    Next lc
    GetOptionalDbl = defaultVal
End Function

' 文字列数値をDoubleにパース
Private Function ToDbl(ByVal v As Variant, Optional ByVal d As Double = 0) As Double
    Dim s As String
    If IsError(v) Or IsNull(v) Or VarType(v) = vbEmpty Then
        ToDbl = d
        Exit Function
    End If

    If IsNumeric(v) Then
        ToDbl = CDbl(v)
        Exit Function
    End If

    s = CStr(v): s = Trim$(s)
    s = Replace(s, "−", "-")      ' U+2212
    s = Replace(s, "▲", "-")
    s = Replace(s, "△", "-")
    s = Replace(s, "(", "-")
    s = Replace(s, ")", "")
    s = Replace(s, ",", "")
    s = StrConv(s, vbNarrow)

    If s = "" Or s = "-" Then
        ToDbl = d
        Exit Function
    End If

    If IsNumeric(s) Then
        ToDbl = CDbl(Val(s))
    Else
        ToDbl = d
    End If
End Function

' 2値の最小値
Private Function Min2(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then Min2 = a Else Min2 = b
End Function
