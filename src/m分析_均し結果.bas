Attribute VB_Name = "m分析_均し結果"
Option Explicit

' ==========================================
' 均し結果分析マクロ
' ==========================================
' 目的: 均しマクロ実行後の日毎生産数量を分析
' 出力: 日次合計、週次合計、ばらつき指標
' ==========================================

Sub 均し結果を分析()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.StatusBar = "均し結果を分析中..."

    ' シート参照
    Dim ws均し As Worksheet
    Set ws均し = ThisWorkbook.Sheets("均し")

    Dim tbl均し As ListObject
    Set tbl均し = ws均し.ListObjects("_成形展開均し")

    ' データ行チェック
    If tbl均し.DataBodyRange Is Nothing Then
        MsgBox "エラー: テーブル「_成形展開均し」にデータ行がありません", vbCritical
        Application.StatusBar = False
        Exit Sub
    End If

    ' 対象年月取得
    Dim ws展開 As Worksheet
    Set ws展開 = ThisWorkbook.Sheets("展開")

    Dim 対象年月 As Date
    対象年月 = CDate(ws展開.Range("A3").Value)

    Dim 対象年 As Long, 対象月 As Long
    対象年 = Year(対象年月)
    対象月 = Month(対象年月)

    Dim maxDay As Long
    maxDay = Day(DateSerial(対象年, 対象月 + 1, 0))

    Debug.Print "========================================="
    Debug.Print "均し結果分析: " & 対象年 & "/" & 対象月
    Debug.Print "========================================="

    ' ==========================================
    ' 1. 稼働日情報読み込み
    ' ==========================================
    Dim ws品番 As Worksheet
    Set ws品番 = ThisWorkbook.Sheets("品番")

    Dim tblHoliday As ListObject
    Set tblHoliday = ws品番.ListObjects("_休日")

    Dim holidays As Object
    Set holidays = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To tblHoliday.DataBodyRange.Rows.Count
        Dim holidayDate As Date
        holidayDate = CDate(tblHoliday.DataBodyRange(i, 1).Value)

        Dim holidayKey As Long
        holidayKey = Year(holidayDate) * 10000 + Month(holidayDate) * 100 + Day(holidayDate)
        holidays(holidayKey) = True
    Next i

    Dim workDays As Object
    Set workDays = CreateObject("Scripting.Dictionary")

    Dim d As Long, dt As Date, wd As Long, dateKey As Long
    For d = 1 To maxDay
        dt = DateSerial(対象年, 対象月, d)
        wd = Weekday(dt)
        dateKey = Year(dt) * 10000 + Month(dt) * 100 + Day(dt)

        If wd <> 1 And wd <> 7 And Not holidays.Exists(dateKey) Then
            workDays.Add workDays.Count + 1, d
        End If
    Next d

    Dim 稼働日数 As Long
    稼働日数 = workDays.Count

    Debug.Print "稼働日数: " & 稼働日数
    Debug.Print ""

    ' ==========================================
    ' 2. 列インデックス取得
    ' ==========================================
    Dim 開始列 As Long
    開始列 = GetColumnIndex(tbl均し, "1")

    ' ==========================================
    ' 3. 日次合計集計
    ' ==========================================
    Dim 日次合計 As Object
    Set 日次合計 = CreateObject("Scripting.Dictionary")

    Dim arr均し As Variant
    arr均し = tbl均し.DataBodyRange.Value

    Dim r As Long, 数量 As Long
    For d = 1 To maxDay
        日次合計(d) = 0
    Next d

    For r = 1 To UBound(arr均し, 1)
        For d = 1 To maxDay
            If 開始列 + d - 1 <= UBound(arr均し, 2) Then
                数量 = 0
                On Error Resume Next
                数量 = CLng(arr均し(r, 開始列 + d - 1))
                On Error GoTo ErrorHandler

                日次合計(d) = CLng(日次合計(d)) + 数量
            End If
        Next d
    Next r

    ' ==========================================
    ' 4. 統計値算出
    ' ==========================================
    Dim 全体合計 As Long, 稼働日合計 As Long
    全体合計 = 0
    稼働日合計 = 0

    For d = 1 To maxDay
        全体合計 = 全体合計 + CLng(日次合計(d))
    Next d

    Dim wdIdx As Long, 稼働日 As Long
    For wdIdx = 1 To workDays.Count
        稼働日 = CLng(workDays(wdIdx))
        稼働日合計 = 稼働日合計 + CLng(日次合計(稼働日))
    Next wdIdx

    Dim 日次平均 As Double
    日次平均 = 稼働日合計 / 稼働日数

    ' 標準偏差算出
    Dim 偏差平方和 As Double, 稼働日数量 As Long
    偏差平方和 = 0

    For wdIdx = 1 To workDays.Count
        稼働日 = CLng(workDays(wdIdx))
        稼働日数量 = CLng(日次合計(稼働日))
        偏差平方和 = 偏差平方和 + (稼働日数量 - 日次平均) ^ 2
    Next wdIdx

    Dim 標準偏差 As Double, 変動係数 As Double
    標準偏差 = Sqr(偏差平方和 / 稼働日数)
    変動係数 = 標準偏差 / 日次平均 * 100

    ' 最大・最小
    Dim 最大値 As Long, 最小値 As Long, 最大日 As Long, 最小日 As Long
    最大値 = 0
    最小値 = 999999

    For wdIdx = 1 To workDays.Count
        稼働日 = CLng(workDays(wdIdx))
        稼働日数量 = CLng(日次合計(稼働日))

        If 稼働日数量 > 最大値 Then
            最大値 = 稼働日数量
            最大日 = 稼働日
        End If

        If 稼働日数量 < 最小値 Then
            最小値 = 稼働日数量
            最小日 = 稼働日
        End If
    Next wdIdx

    Debug.Print "--- 全体統計 ---"
    Debug.Print "月間合計: " & Format(全体合計, "#,##0") & "個"
    Debug.Print "稼働日合計: " & Format(稼働日合計, "#,##0") & "個"
    Debug.Print "日次平均: " & Format(日次平均, "#,##0.0") & "個"
    Debug.Print "標準偏差: " & Format(標準偏差, "#,##0.0") & "個"
    Debug.Print "変動係数: " & Format(変動係数, "0.0") & "%"
    Debug.Print "最大値: " & Format(最大値, "#,##0") & "個 (" & 最大日 & "日)"
    Debug.Print "最小値: " & Format(最小値, "#,##0") & "個 (" & 最小日 & "日)"
    Debug.Print "レンジ: " & Format(最大値 - 最小値, "#,##0") & "個 (最大-最小)"
    Debug.Print ""

    ' ==========================================
    ' 5. 日次詳細出力
    ' ==========================================
    Debug.Print "--- 日次詳細 ---"
    Debug.Print "日付, 曜日, 数量, 平均比, 累計"

    Dim 累計 As Long
    累計 = 0

    For d = 1 To maxDay
        dt = DateSerial(対象年, 対象月, d)
        wd = Weekday(dt)
        dateKey = Year(dt) * 10000 + Month(dt) * 100 + Day(dt)

        Dim 曜日名 As String
        Select Case wd
            Case 1: 曜日名 = "日"
            Case 2: 曜日名 = "月"
            Case 3: 曜日名 = "火"
            Case 4: 曜日名 = "水"
            Case 5: 曜日名 = "木"
            Case 6: 曜日名 = "金"
            Case 7: 曜日名 = "土"
        End Select

        数量 = CLng(日次合計(d))
        累計 = 累計 + 数量

        Dim 平均比 As String
        If wd = 1 Or wd = 7 Or holidays.Exists(dateKey) Then
            ' 休日
            平均比 = "-"
            Debug.Print Format(d, "00") & "日, " & 曜日名 & ", " & _
                       Format(数量, "#,##0") & ", " & 平均比 & ", " & _
                       Format(累計, "#,##0") & " [休日]"
        Else
            ' 稼働日
            Dim 平均比率 As Double
            平均比率 = 数量 / 日次平均 * 100
            平均比 = Format(平均比率, "0.0") & "%"

            Dim 偏差表示 As String
            If 平均比率 > 120 Then
                偏差表示 = " ⚠ 過剰"
            ElseIf 平均比率 < 80 Then
                偏差表示 = " ⚠ 過少"
            Else
                偏差表示 = ""
            End If

            Debug.Print Format(d, "00") & "日, " & 曜日名 & ", " & _
                       Format(数量, "#,##0") & ", " & 平均比 & ", " & _
                       Format(累計, "#,##0") & 偏差表示
        End If
    Next d

    Debug.Print ""

    ' ==========================================
    ' 6. 週次集計
    ' ==========================================
    Debug.Print "--- 週次集計 ---"

    Dim 週番号 As Long, 週開始日 As Long, 週終了日 As Long
    Dim 週合計 As Long, 週稼働日数 As Long

    週番号 = 1
    週開始日 = 1

    Do While 週開始日 <= maxDay
        週終了日 = 週開始日 + 6
        If 週終了日 > maxDay Then 週終了日 = maxDay

        週合計 = 0
        週稼働日数 = 0

        For d = 週開始日 To 週終了日
            dt = DateSerial(対象年, 対象月, d)
            wd = Weekday(dt)
            dateKey = Year(dt) * 10000 + Month(dt) * 100 + Day(dt)

            If wd <> 1 And wd <> 7 And Not holidays.Exists(dateKey) Then
                週稼働日数 = 週稼働日数 + 1
            End If

            週合計 = 週合計 + CLng(日次合計(d))
        Next d

        Dim 週平均 As Double
        If 週稼働日数 > 0 Then
            週平均 = 週合計 / 週稼働日数
        Else
            週平均 = 0
        End If

        Debug.Print "第" & 週番号 & "週 (" & 週開始日 & "-" & 週終了日 & "日): " & _
                   Format(週合計, "#,##0") & "個, " & _
                   "稼働" & 週稼働日数 & "日, " & _
                   "平均" & Format(週平均, "#,##0.0") & "個/日"

        週番号 = 週番号 + 1
        週開始日 = 週終了日 + 1
    Loop

    Debug.Print ""

    ' ==========================================
    ' 7. 改善提案生成
    ' ==========================================
    Debug.Print "--- 改善提案 ---"

    ' 問題日の検出
    Dim 極端過剰日 As Object, 極端過少日 As Object
    Dim 中程度過剰日 As Object, 中程度過少日 As Object
    Set 極端過剰日 = CreateObject("Scripting.Dictionary")
    Set 極端過少日 = CreateObject("Scripting.Dictionary")
    Set 中程度過剰日 = CreateObject("Scripting.Dictionary")
    Set 中程度過少日 = CreateObject("Scripting.Dictionary")

    For wdIdx = 1 To workDays.Count
        稼働日 = CLng(workDays(wdIdx))
        稼働日数量 = CLng(日次合計(稼働日))
        平均比率 = 稼働日数量 / 日次平均 * 100

        If 平均比率 >= 150 Then
            極端過剰日(稼働日) = 稼働日数量
        ElseIf 平均比率 > 120 Then
            中程度過剰日(稼働日) = 稼働日数量
        ElseIf 平均比率 <= 50 Then
            極端過少日(稼働日) = 稼働日数量
        ElseIf 平均比率 < 80 Then
            中程度過少日(稼働日) = 稼働日数量
        End If
    Next wdIdx

    ' 改善提案メッセージ生成
    Dim 改善必要 As Boolean
    改善必要 = False

    If 極端過剰日.Count > 0 Then
        改善必要 = True
        Debug.Print ""
        Debug.Print "【フェーズ1: 極端な過剰日の調整】"
        Debug.Print "平均の150%超の日が " & 極端過剰日.Count & " 日あります："

        Dim key As Variant
        For Each key In 極端過剰日.Keys
            稼働日 = CLng(key)
            稼働日数量 = CLng(極端過剰日(稼働日))
            平均比率 = 稼働日数量 / 日次平均 * 100
            Dim 超過量 As Long
            超過量 = 稼働日数量 - CLng(日次平均 * 1.2)

            Debug.Print "  " & 稼働日 & "日: " & Format(稼働日数量, "#,##0") & "個 (" & _
                       Format(平均比率, "0.0") & "%) → 約" & Format(超過量, "#,##0") & "個を他日に移動"
        Next key
    End If

    If 極端過少日.Count > 0 Then
        改善必要 = True
        Debug.Print ""
        Debug.Print "【フェーズ2: 極端な過少日の調整】"
        Debug.Print "平均の50%未満の日が " & 極端過少日.Count & " 日あります："

        For Each key In 極端過少日.Keys
            稼働日 = CLng(key)
            稼働日数量 = CLng(極端過少日(稼働日))
            平均比率 = 稼働日数量 / 日次平均 * 100
            Dim 不足量 As Long
            不足量 = CLng(日次平均 * 0.8) - 稼働日数量

            Debug.Print "  " & 稼働日 & "日: " & Format(稼働日数量, "#,##0") & "個 (" & _
                       Format(平均比率, "0.0") & "%) → 約" & Format(不足量, "#,##0") & "個を追加"
        Next key
    End If

    If 中程度過剰日.Count > 0 Then
        改善必要 = True
        Debug.Print ""
        Debug.Print "【フェーズ3: 中程度の過剰日の調整】"
        Debug.Print "平均の120%～150%の日が " & 中程度過剰日.Count & " 日あります："

        For Each key In 中程度過剰日.Keys
            稼働日 = CLng(key)
            稼働日数量 = CLng(中程度過剰日(稼働日))
            平均比率 = 稼働日数量 / 日次平均 * 100
            超過量 = 稼働日数量 - CLng(日次平均)

            Debug.Print "  " & 稼働日 & "日: " & Format(稼働日数量, "#,##0") & "個 (" & _
                       Format(平均比率, "0.0") & "%) → 約" & Format(超過量, "#,##0") & "個を他日に移動"
        Next key
    End If

    If 中程度過少日.Count > 0 Then
        改善必要 = True
        Debug.Print ""
        Debug.Print "【フェーズ4: 中程度の過少日の調整】"
        Debug.Print "平均の50%～80%の日が " & 中程度過少日.Count & " 日あります："

        For Each key In 中程度過少日.Keys
            稼働日 = CLng(key)
            稼働日数量 = CLng(中程度過少日(稼働日))
            平均比率 = 稼働日数量 / 日次平均 * 100
            不足量 = CLng(日次平均) - 稼働日数量

            Debug.Print "  " & 稼働日 & "日: " & Format(稼働日数量, "#,##0") & "個 (" & _
                       Format(平均比率, "0.0") & "%) → 約" & Format(不足量, "#,##0") & "個を追加"
        Next key
    End If

    If Not 改善必要 Then
        Debug.Print ""
        Debug.Print "すべての稼働日が平均±20%以内です。"
        Debug.Print "改善の必要はありません。"
    Else
        Debug.Print ""
        Debug.Print "※ 自動調整マクロ「m調整_自動均し」で段階的に改善できます"
    End If

    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "分析完了"
    Debug.Print "========================================="

    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
End Sub

' ==========================================
' テーブル列インデックス取得
' ==========================================
Private Function GetColumnIndex(ByRef tbl As ListObject, ByVal colName As String) As Long
    Dim i As Long
    For i = 1 To tbl.ListColumns.Count
        If tbl.ListColumns(i).Name = colName Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 1, "GetColumnIndex", "列[" & colName & "]が見つかりません"
End Function
