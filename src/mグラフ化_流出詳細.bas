Attribute VB_Name = "mグラフ化_流出詳細"
Option Explicit

' ========================================
' マクロ名: グラフ化_流出詳細
' 処理概要: 期間Aと期間Bのデータを統合し、中央292をピークとする富士山型ウォーターフォールグラフを生成
' ソーステーブル: シート「10/1～10/17 詳細」テーブル「_期間A」（工程/流出/廃棄/数量）
'                シート「10/1～10/17 詳細」テーブル「_期間B」（工程/成形/塗装/数量）
' ターゲットテーブル: シート「期間AB_変換」テーブル「期間AB_変換」
' 処理方式: 累積値ベースのウォーターフォール（Base = Min(前累積, 次累積)）
' グラフ構成: 左側（流出/廃棄）→ 中央（単色292）→ 右側（成形/塗装）
' ========================================

Sub グラフ化_流出詳細()
    ' ============================================
    ' 最適化設定の保存と有効化
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

    Application.StatusBar = "処理を開始します..."

    ' ============================================
    ' ワークブックとテーブルの取得
    ' ============================================
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim loA As ListObject, loB As ListObject
    Set loA = GetListObjectByName(wb, "_期間A")
    Set loB = GetListObjectByName(wb, "_期間B")

    If loA Is Nothing Then Err.Raise 5, , "テーブル「_期間A」が見つかりません"
    If loB Is Nothing Then Err.Raise 5, , "テーブル「_期間B」が見つかりません"

    ' ============================================
    ' 出力シート・テーブルの準備（完全削除→新規作成）
    ' ============================================
    Dim wsOut As Worksheet
    On Error Resume Next
    Set wsOut = wb.Worksheets("期間AB_変換")
    If Not wsOut Is Nothing Then
        Application.DisplayAlerts = False
        wsOut.Delete
        Application.DisplayAlerts = True
    End If
    Err.Clear
    On Error GoTo ErrorHandler

    Set wsOut = wb.Worksheets.Add(After:=loA.Parent)
    wsOut.Name = "期間AB_変換"

    ' ヘッダー作成
    With wsOut
        .Range("A1:H1").Value = Array("工程", "Base", "流出", "廃棄", "成形", "塗装", "単色", "増減符号")
    End With

    Dim loOut As ListObject
    Set loOut = wsOut.ListObjects.Add(xlSrcRange, wsOut.Range("A1:H2"), , xlYes)
    loOut.Name = "期間AB_変換"

    If Not loOut.DataBodyRange Is Nothing Then loOut.DataBodyRange.Delete

    Application.StatusBar = "データ変換中..."

    ' ============================================
    ' 期間Aのデータ処理（左側：流出/廃棄）
    ' ============================================
    Dim cum As Double
    cum = 0

    Dim r As Long, n As Long
    n = loA.DataBodyRange.Rows.Count

    Dim nm As String, leak As Double, scrap As Double, qty As Double
    Dim delta As Double, nextCum As Double, baseVal As Double
    Dim hasQty As Boolean

    For r = 1 To n
        nm = CStr(loA.ListColumns("工程").DataBodyRange(r, 1).Value)
        leak = ToDbl(loA.ListColumns("流出").DataBodyRange(r, 1).Value)
        scrap = ToDbl(loA.ListColumns("廃棄").DataBodyRange(r, 1).Value)

        ' 数量列がある場合はその値を使用、なければ内訳合算
        hasQty = TryToDbl(GetOptional(loA, "数量", r), qty)
        If hasQty Then
            delta = qty
        Else
            delta = leak + scrap
        End If

        nextCum = cum + delta
        baseVal = Min2(cum, nextCum)

        ' 変換テーブルに行追加
        Dim rowOut As ListRow
        Set rowOut = loOut.ListRows.Add
        rowOut.Range(1, 1).Value = nm                      ' 工程
        rowOut.Range(1, 2).Value = baseVal                 ' Base
        rowOut.Range(1, 3).Value = Abs(leak)               ' 流出（常に正）
        rowOut.Range(1, 4).Value = Abs(scrap)              ' 廃棄（常に正）
        rowOut.Range(1, 5).Value = 0                       ' 成形（期間Aでは0）
        rowOut.Range(1, 6).Value = 0                       ' 塗装（期間Aでは0）
        rowOut.Range(1, 7).Value = 0                       ' 単色（期間Aでは0）
        rowOut.Range(1, 8).Value = IIf(delta < 0, -1, 1)   ' 増減符号

        cum = nextCum
    Next r

    ' ============================================
    ' 中央の単色棒（加工流出総数）を追加
    ' ============================================
    Set rowOut = loOut.ListRows.Add
    rowOut.Range(1, 1).Value = "加工流出総数"
    rowOut.Range(1, 2).Value = 0               ' Base = 0（地面から292まで）
    rowOut.Range(1, 3).Value = 0               ' 流出
    rowOut.Range(1, 4).Value = 0               ' 廃棄
    rowOut.Range(1, 5).Value = 0               ' 成形
    rowOut.Range(1, 6).Value = 0               ' 塗装
    rowOut.Range(1, 7).Value = Abs(cum)        ' 単色（292）
    rowOut.Range(1, 8).Value = 0               ' 増減符号（中立）

    ' ============================================
    ' 期間Bのデータ処理（右側：成形/塗装）
    ' ============================================
    ' cumは292からスタート（期間Aの最終累積値）

    n = loB.DataBodyRange.Rows.Count
    Dim molding As Double, painting As Double

    For r = 1 To n
        nm = CStr(loB.ListColumns("工程").DataBodyRange(r, 1).Value)

        ' 「加工流出総数」行はスキップ（中央の単色棒で既に表示済み）
        If InStr(1, nm, "加工流出総数", vbTextCompare) > 0 Then
            GoTo NextRowB
        End If

        molding = ToDbl(loB.ListColumns("成形").DataBodyRange(r, 1).Value)
        painting = ToDbl(loB.ListColumns("塗装").DataBodyRange(r, 1).Value)

        ' 数量列がある場合はその値を使用、なければ内訳合算
        hasQty = TryToDbl(GetOptional(loB, "数量", r), qty)
        If hasQty Then
            delta = qty
        Else
            delta = molding + painting
        End If

        nextCum = cum + delta
        baseVal = Min2(cum, nextCum)

        ' 変換テーブルに行追加
        Set rowOut = loOut.ListRows.Add
        rowOut.Range(1, 1).Value = nm                      ' 工程
        rowOut.Range(1, 2).Value = baseVal                 ' Base
        rowOut.Range(1, 3).Value = 0                       ' 流出（期間Bでは0）
        rowOut.Range(1, 4).Value = 0                       ' 廃棄（期間Bでは0）
        rowOut.Range(1, 5).Value = Abs(molding)            ' 成形（常に正）
        rowOut.Range(1, 6).Value = Abs(painting)           ' 塗装（常に正）
        rowOut.Range(1, 7).Value = 0                       ' 単色（期間Bでは0）
        rowOut.Range(1, 8).Value = IIf(delta < 0, -1, 1)   ' 増減符号

        cum = nextCum

NextRowB:
    Next r

    Application.StatusBar = "グラフ生成中..."

    ' ============================================
    ' 積み上げ縦棒グラフの作成
    ' ============================================
    Dim chObj As ChartObject, ch As Chart
    Set chObj = wsOut.ChartObjects.Add(Left:=wsOut.Range("J2").Left, _
                                       Top:=wsOut.Range("J2").Top, _
                                       Width:=720, _
                                       Height:=420)
    Set ch = chObj.Chart
    ch.ChartType = xlColumnStacked

    Dim lastR As Long
    lastR = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row

    ' データ範囲設定（Base + 流出 + 廃棄 + 成形 + 塗装 + 単色）
    ch.SetSourceData wsOut.Range("B1:G" & lastR)
    ch.Axes(xlCategory).CategoryNames = wsOut.Range("A2:A" & lastR)
    ch.SetElement (msoElementLegendBottom)

    ' ============================================
    ' Base系列を透明化
    ' ============================================
    With ch.FullSeriesCollection(1)
        .Name = "Base"
        .Format.Fill.Visible = msoFalse
        .Format.Line.Visible = msoFalse
    End With

    ' ============================================
    ' グラフの基本設定
    ' ============================================
    ch.HasTitle = True
    ch.ChartTitle.Text = "富士山型ウォーターフォール（期間A/B統合）"
    ch.ChartGroups(1).GapWidth = 50
    ch.Axes(xlValue).HasMajorGridlines = True

    ' 系列名設定
    ch.FullSeriesCollection(2).Name = "流出"
    ch.FullSeriesCollection(3).Name = "廃棄"
    ch.FullSeriesCollection(4).Name = "成形"
    ch.FullSeriesCollection(5).Name = "塗装"
    ch.FullSeriesCollection(6).Name = "総数"

    ' ============================================
    ' 軸範囲の推定と設定
    ' ============================================
    Dim minV As Double, maxV As Double
    EstimateMinMaxMinBase wsOut, 2, 7, minV, maxV

    With ch.Axes(xlValue)
        .MinimumScale = WorksheetFunction.Floor(minV * 1.1, 1)
        .MaximumScale = WorksheetFunction.Ceiling(maxV * 1.1, 1)
    End With

    Application.StatusBar = "配色設定中..."

    ' ============================================
    ' 系列ごとの配色（増減符号に基づく）
    ' ============================================
    ColorSeriesBySign ch, wsOut.Range("H2:H" & lastR)

    Application.StatusBar = "処理が完了しました"
    Application.Wait Now + TimeValue("00:00:01")

    GoTo Cleanup

ErrorHandler:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear

    MsgBox "エラーが発生しました" & vbCrLf & _
           "エラー番号: " & errNum & vbCrLf & _
           "詳細: " & errDesc, vbCritical, "エラー"

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.DisplayAlerts = origDisplayAlerts
End Sub


' ================ ヘルパー関数群 ================

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

' 任意列のr行目の値（無ければEmpty）
Private Function GetOptional(ByVal lo As ListObject, ByVal colName As String, ByVal r As Long) As Variant
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, colName, vbTextCompare) = 0 Then
            GetOptional = lc.DataBodyRange(r, 1).Value
            Exit Function
        End If
    Next lc
    GetOptional = Empty
End Function

' 文字列数値をDoubleにパース（全角/カンマ/会計表記/Unicodeマイナス対応）
Private Function TryToDbl(ByVal v As Variant, ByRef outVal As Double) As Boolean
    Dim s As String
    If IsError(v) Or IsNull(v) Or VarType(v) = vbEmpty Then Exit Function
    s = CStr(v): s = Trim$(s)
    s = Replace(s, "−", "-")      ' U+2212
    s = Replace(s, "▲", "-")      ' ▲15 → -15
    s = Replace(s, "△", "-")
    s = Replace(s, "(", "-")      ' (15) → -15
    s = Replace(s, ")", "")
    s = Replace(s, ",", "")
    s = StrConv(s, vbNarrow)      ' 全角→半角
    If s = "" Or s = "-" Then Exit Function
    If Not IsNumeric(s) Then Exit Function
    outVal = CDbl(Val(s))
    TryToDbl = True
End Function

Private Function ToDbl(ByVal v As Variant, Optional ByVal d As Double = 0) As Double
    Dim t As Double
    If TryToDbl(v, t) Then ToDbl = t Else ToDbl = d
End Function

' 2値の最小
Private Function Min2(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then Min2 = a Else Min2 = b
End Function

' 軸推定（各列の到達点：Base..Base+ABS(内訳合計)）
Private Sub EstimateMinMaxMinBase(ws As Worksheet, firstDataCol As Long, lastDataCol As Long, _
                                  ByRef minV As Double, ByRef maxV As Double)
    Dim r As Long, lastR As Long, baseV As Double, s As Double, up As Double, dn As Double, c As Long
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    minV = 0: maxV = 0
    For r = 2 To lastR
        baseV = ToDbl(ws.Cells(r, 2).Value)
        s = 0
        For c = firstDataCol + 1 To lastDataCol
            s = s + Abs(ToDbl(ws.Cells(r, c).Value))
        Next c
        up = baseV + s
        dn = baseV
        If up > maxV Then maxV = up
        If dn < minV Then minV = dn
    Next r
End Sub

' 系列ごとの配色（増減符号に基づく）
Private Sub ColorSeriesBySign(ByVal ch As Chart, ByVal rngSign As Range)
    Dim n As Long: n = rngSign.Rows.Count

    ' 系列取得
    Dim sLeak As Series, sScrap As Series
    Dim sMolding As Series, sPainting As Series
    Dim sTotal As Series

    Set sLeak = ch.FullSeriesCollection(2)      ' 流出
    Set sScrap = ch.FullSeriesCollection(3)     ' 廃棄
    Set sMolding = ch.FullSeriesCollection(4)   ' 成形
    Set sPainting = ch.FullSeriesCollection(5)  ' 塗装
    Set sTotal = ch.FullSeriesCollection(6)     ' 単色（総数）

    ' 配色定義（Tailwind相当の階調）
    ' 期間A（青系）
    Dim aDark As Long, aLight As Long
    aDark = RGB(37, 99, 235)     ' 青・濃
    aLight = RGB(147, 197, 253)  ' 青・淡

    ' 期間B（緑系）
    Dim bDark As Long, bLight As Long
    bDark = RGB(34, 197, 94)     ' 緑・濃
    bLight = RGB(134, 239, 172)  ' 緑・淡

    ' マイナス色（赤系）
    Dim negDark As Long, negLight As Long
    negDark = RGB(220, 38, 38)    ' 赤・濃
    negLight = RGB(252, 165, 165) ' 赤・淡

    ' 中央総数（グレー）
    Dim totalGray As Long
    totalGray = RGB(107, 114, 128)

    Dim i As Long, signVal As Long
    For i = 1 To n
        signVal = CInt(ToDbl(rngSign.Cells(i, 1).Value))

        ' 流出（期間A・濃色）
        If sLeak.Points(i).HasDataLabel = False Then
            With sLeak.Points(i).Format.Fill
                .Visible = msoTrue
                .Solid
                .ForeColor.RGB = IIf(signVal < 0, negDark, aDark)
            End With
            With sLeak.Points(i).Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 255, 255)
                .Weight = 0.75
            End With
        End If

        ' 廃棄（期間A・淡色）
        If sScrap.Points(i).HasDataLabel = False Then
            With sScrap.Points(i).Format.Fill
                .Visible = msoTrue
                .Solid
                .ForeColor.RGB = IIf(signVal < 0, negLight, aLight)
            End With
            With sScrap.Points(i).Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 255, 255)
                .Weight = 0.75
            End With
        End If

        ' 成形（期間B・濃色）
        If sMolding.Points(i).HasDataLabel = False Then
            With sMolding.Points(i).Format.Fill
                .Visible = msoTrue
                .Solid
                .ForeColor.RGB = IIf(signVal < 0, negDark, bDark)
            End With
            With sMolding.Points(i).Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 255, 255)
                .Weight = 0.75
            End With
        End If

        ' 塗装（期間B・淡色）
        If sPainting.Points(i).HasDataLabel = False Then
            With sPainting.Points(i).Format.Fill
                .Visible = msoTrue
                .Solid
                .ForeColor.RGB = IIf(signVal < 0, negLight, bLight)
            End With
            With sPainting.Points(i).Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 255, 255)
                .Weight = 0.75
            End With
        End If

        ' 単色（中央総数）
        If sTotal.Points(i).HasDataLabel = False Then
            With sTotal.Points(i).Format.Fill
                .Visible = msoTrue
                .Solid
                .ForeColor.RGB = totalGray
            End With
            With sTotal.Points(i).Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(255, 255, 255)
                .Weight = 0.75
            End With
        End If
    Next i
End Sub
