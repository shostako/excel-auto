Attribute VB_Name = "m転記_負荷均し"
Option Explicit

' ==========================================
' 負荷均しマクロ（制約追加版）
' ==========================================
' 月間の成形品番生産数を稼働日に均等配分
' ソース: テーブル「_成形展開」
' ターゲット: テーブル「_成形展開均し」
' マスタ: テーブル「_品番」「_休日」「_パラメータ」
'
' 【追加制約】
' 1. 補給品優先配置: モール品補給品→非モール補給品の順で優先配置
' 2. 系列別処理: モール×アルヴェル系 → ノアヴォク系 → その他の順で均し
' 3. 号口単品分散: 号口かつ単品は全て異なる日に分散配置
' 4. 補給品×号口単品同日禁止: 補給品と号口単品は同じ日に配置しない
' ==========================================

Sub 転記_負荷均し()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.StatusBar = "負荷均し処理を開始します..."

    ' 各シートの参照
    Dim ws品番 As Worksheet, ws展開 As Worksheet, ws均し As Worksheet
    Set ws品番 = ThisWorkbook.Sheets("品番")
    Set ws展開 = ThisWorkbook.Sheets("展開")
    Set ws均し = ThisWorkbook.Sheets("均し")

    ' ==========================================
    ' 1. パラメータ読み込み
    ' ==========================================
    Application.StatusBar = "パラメータを読み込み中..."

    Dim tblParam As ListObject
    Set tblParam = ws品番.ListObjects("_パラメータ")

    Dim paramDict As Object
    Set paramDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To tblParam.DataBodyRange.Rows.Count
        paramDict(CStr(tblParam.DataBodyRange(i, 1).Value)) = tblParam.DataBodyRange(i, 2).Value
    Next i

    Dim 誤差許容率 As Double
    Dim グループ制約モード As String, 月末処理モード As String

    誤差許容率 = CDbl(paramDict("日次目標誤差許容率(%)"))
    グループ制約モード = CStr(paramDict("グループ制約モード"))
    月末処理モード = CStr(paramDict("月末残数処理モード"))

    ' 対象年月（「展開」シートのセルA3から取得）
    Dim 対象年 As Long, 対象月 As Long
    Dim 対象年月 As Date
    対象年月 = CDate(ws展開.Range("A3").Value)
    対象年 = Year(対象年月)
    対象月 = Month(対象年月)

    Debug.Print "=== 負荷均し処理開始（制約追加版） ==="
    Debug.Print "対象年月: " & 対象年 & "/" & 対象月
    Debug.Print "誤差許容率: " & 誤差許容率 & "%"
    Debug.Print "グループ制約: " & グループ制約モード
    Debug.Print "月末処理: " & 月末処理モード

    ' ==========================================
    ' 2. 稼働日算出
    ' ==========================================
    Application.StatusBar = "稼働日を算出中..."

    Dim tblHoliday As ListObject
    Set tblHoliday = ws品番.ListObjects("_休日")

    Dim holidays As Object
    Set holidays = CreateObject("Scripting.Dictionary")

    ' 休日テーブル読み込み（整数値化: YYYYMMDD）
    For i = 1 To tblHoliday.DataBodyRange.Rows.Count
        Dim holidayDate As Date
        holidayDate = CDate(tblHoliday.DataBodyRange(i, 1).Value)

        Dim holidayKey As Long
        holidayKey = Year(holidayDate) * 10000 + Month(holidayDate) * 100 + Day(holidayDate)
        holidays(holidayKey) = True

        Debug.Print "休日登録: " & Format(holidayDate, "yyyy/mm/dd") & " → " & holidayKey
    Next i

    Dim workDays As Object
    Set workDays = CreateObject("Scripting.Dictionary")

    Dim maxDay As Long
    maxDay = Day(DateSerial(対象年, 対象月 + 1, 0))

    Dim d As Long, dt As Date, wd As Long, dateKey As Long
    For d = 1 To maxDay
        dt = DateSerial(対象年, 対象月, d)
        wd = Weekday(dt)

        ' 日付を整数値化（YYYYMMDD）
        dateKey = Year(dt) * 10000 + Month(dt) * 100 + Day(dt)

        If wd = 1 Or wd = 7 Then
            ' 土日
            Debug.Print d & "日: 土日除外"
        ElseIf holidays.Exists(dateKey) Then
            ' 休日
            Debug.Print d & "日: 休日除外 (dateKey=" & dateKey & ")"
        Else
            workDays.Add workDays.Count + 1, d
            Debug.Print d & "日: 稼働日に追加 (dateKey=" & dateKey & ")"
        End If
    Next d

    Dim 稼働日数 As Long
    稼働日数 = workDays.Count

    Debug.Print "稼働日数: " & 稼働日数
    Debug.Print "稼働日: " & Join(DictValuesToArray(workDays), ", ")

