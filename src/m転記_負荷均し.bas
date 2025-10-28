Attribute VB_Name = "m転記_負荷均し"
Option Explicit

' ==========================================
' 負荷均しマクロ
' ==========================================
' 月間の成形品番生産数を稼働日に均等配分
' ソース: テーブル「_成形展開」
' ターゲット: テーブル「_成形展開均し」
' マスタ: テーブル「_品番」「_休日」「_パラメータ」
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

    Dim 誤差許容率 As Double, 誤差許容個数 As Long
    Dim グループ制約モード As String, 月末処理モード As String

    誤差許容率 = CDbl(paramDict("日次目標誤差許容率(%)"))
    誤差許容個数 = CLng(paramDict("日次目標誤差許容個数(個)"))
    グループ制約モード = CStr(paramDict("グループ制約モード"))
    月末処理モード = CStr(paramDict("月末残数処理モード"))

    ' 対象年月（「展開」シートのセルA3から取得）
    Dim 対象年 As Long, 対象月 As Long
    Dim 対象年月 As Date
    対象年月 = CDate(ws展開.Range("A3").Value)
    対象年 = Year(対象年月)
    対象月 = Month(対象年月)

    Debug.Print "=== 負荷均し処理開始 ==="
    Debug.Print "対象年月: " & 対象年 & "/" & 対象月
    Debug.Print "誤差許容率: " & 誤差許容率 & "%"
    Debug.Print "誤差許容個数: " & 誤差許容個数 & "個"
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

    ' ==========================================
    ' 3. 月間総需要集計
    ' ==========================================
    Application.StatusBar = "月間総需要を集計中..."

    Dim tbl成形展開 As ListObject
    Set tbl成形展開 = ws展開.ListObjects("_成形展開")

    ' データ行チェック
    If tbl成形展開.DataBodyRange Is Nothing Then
        MsgBox "エラー: テーブル「_成形展開」にデータ行がありません", vbCritical
        Application.StatusBar = False
        Exit Sub
    End If

    Debug.Print "テーブル「_成形展開」行数: " & tbl成形展開.DataBodyRange.Rows.Count
    Debug.Print "テーブル「_成形展開」列数: " & tbl成形展開.ListColumns.Count

    Dim arr展開 As Variant
    arr展開 = tbl成形展開.DataBodyRange.Value

    ' 列インデックス取得
    Dim 成形品番列 As Long
    On Error GoTo ColumnError
    成形品番列 = GetColumnIndex(tbl成形展開, "成形品番")
    Debug.Print "成形品番列: " & 成形品番列

    Dim 開始列 As Long
    開始列 = GetColumnIndex(tbl成形展開, "1")
    Debug.Print "開始列: " & 開始列
    On Error GoTo ErrorHandler

    Dim 月間総需要 As Object, 月間残数 As Object
    Set 月間総需要 = CreateObject("Scripting.Dictionary")
    Set 月間残数 = CreateObject("Scripting.Dictionary")

    Dim 全品番総需要 As Long
    全品番総需要 = 0

    Dim r As Long, 品番 As String, 合計 As Long
    For r = 1 To UBound(arr展開, 1)
        品番 = CStr(arr展開(r, 成形品番列))
        合計 = 0

        For d = 1 To maxDay
            If 開始列 + d - 1 <= UBound(arr展開, 2) Then
                合計 = 合計 + CLng(arr展開(r, 開始列 + d - 1))
            End If
        Next d

        月間総需要(品番) = 合計
        月間残数(品番) = 合計
        全品番総需要 = 全品番総需要 + 合計
    Next r

    Debug.Print "全品番総需要: " & 全品番総需要

    ' ==========================================
    ' 4. 日次目標算出
    ' ==========================================
    Dim 日次目標 As Double
    日次目標 = 全品番総需要 / 稼働日数

    Debug.Print "日次目標: " & Format(日次目標, "0.0")

    ' ==========================================
    ' 5. 品番マスタ読み込み
    ' ==========================================
    Application.StatusBar = "品番マスタを読み込み中..."

    Dim tbl品番 As ListObject
    Set tbl品番 = ws品番.ListObjects("_品番")

    Dim arr品番 As Variant
    arr品番 = tbl品番.DataBodyRange.Value

    Dim 品番マスタ As Object
    Set 品番マスタ = CreateObject("Scripting.Dictionary")

    ' 列インデックス取得
    Dim 品番_成形品番列 As Long, 品番_単位列 As Long, 品番_上限列 As Long
    Dim 品番_優先度列 As Long, 品番_グループ列 As Long
    Dim 品番_型番列 As Long, 品番_セット列 As Long

    品番_成形品番列 = GetColumnIndex(tbl品番, "成形品番")
    品番_単位列 = GetColumnIndex(tbl品番, "単位")
    品番_上限列 = GetColumnIndex(tbl品番, "上限")
    品番_優先度列 = GetColumnIndex(tbl品番, "優先度")
    品番_グループ列 = GetColumnIndex(tbl品番, "グループ")
    品番_型番列 = GetColumnIndex(tbl品番, "型番")
    品番_セット列 = GetColumnIndex(tbl品番, "セット")

    For r = 1 To UBound(arr品番, 1)
        On Error Resume Next
        Dim info As Object
        Set info = CreateObject("Scripting.Dictionary")

        Dim 成形品番値 As Variant, 単位値 As Variant, 上限値 As Variant, 優先度値 As Variant
        成形品番値 = arr品番(r, 品番_成形品番列)
        単位値 = arr品番(r, 品番_単位列)
        上限値 = arr品番(r, 品番_上限列)
        優先度値 = arr品番(r, 品番_優先度列)

        ' 空行チェック
        If IsEmpty(成形品番値) Or 成形品番値 = "" Then
            Debug.Print "行" & r & ": 成形品番が空 - スキップ"
            GoTo NextRow
        End If

        ' 数値型チェック
        If Not IsNumeric(単位値) Or IsEmpty(単位値) Then
            Debug.Print "行" & r & ": 単位が数値でない - デフォルト1"
            単位値 = 1
        End If
        If Not IsNumeric(上限値) Or IsEmpty(上限値) Then
            Debug.Print "行" & r & ": 上限が数値でない - デフォルト9999"
            上限値 = 9999
        End If
        If Not IsNumeric(優先度値) Or IsEmpty(優先度値) Then
            Debug.Print "行" & r & ": 優先度が数値でない - デフォルト3"
            優先度値 = 3
        End If

        info("単位") = CLng(単位値)
        info("上限") = CLng(上限値)
        info("優先度") = CLng(優先度値)

        Dim grpVal As Variant, 型番Val As Variant, セットVal As Variant
        grpVal = arr品番(r, 品番_グループ列)
        info("グループ") = IIf(IsEmpty(grpVal) Or IsNull(grpVal) Or grpVal = "", "", CStr(grpVal))

        型番Val = arr品番(r, 品番_型番列)
        info("型番") = IIf(IsEmpty(型番Val) Or IsNull(型番Val) Or 型番Val = "", "", CStr(型番Val))

        セットVal = arr品番(r, 品番_セット列)
        info("セット") = IIf(IsEmpty(セットVal) Or IsNull(セットVal) Or セットVal = "", "", CStr(セットVal))

        Set 品番マスタ(CStr(成形品番値)) = info

        If Err.Number <> 0 Then
            Debug.Print "行" & r & " エラー: " & Err.Description
            Err.Clear
        End If
NextRow:
        On Error GoTo ErrorHandler
    Next r

    Debug.Print "品番マスタ件数: " & 品番マスタ.Count

    ' ==========================================
    ' 5.5. セットペアマスタ構築
    ' ==========================================
    Application.StatusBar = "セットペアマスタを構築中..."

    Dim セットペアマスタ As Object
    Set セットペアマスタ = CreateObject("Scripting.Dictionary")

    Dim key As Variant, セットベース As String
    For Each key In 品番マスタ.Keys
        品番 = CStr(key)

        If 品番マスタ(品番)("セット") = "SET" Then
            セットベース = GetSetBase(品番)

            If Not セットペアマスタ.Exists(セットベース) Then
                Set セットペアマスタ(セットベース) = CreateObject("Scripting.Dictionary")
            End If

            セットペアマスタ(セットベース)(品番) = True
        End If
    Next key

    Debug.Print "セットペアマスタ件数: " & セットペアマスタ.Count

    ' ペア整合性チェック（F/R両方存在するか）
    For Each key In セットペアマスタ.Keys
        セットベース = CStr(key)
        Dim ペアF As String, ペアR As String
        ペアF = セットベース & "F"
        ペアR = セットベース & "R"

        Dim ペアF存在 As Boolean, ペアR存在 As Boolean
        ペアF存在 = セットペアマスタ(セットベース).Exists(ペアF)
        ペアR存在 = セットペアマスタ(セットベース).Exists(ペアR)

        If Not ペアF存在 Or Not ペアR存在 Then
            Debug.Print "警告: セットペア不整合 [" & セットベース & "] F=" & ペアF存在 & " R=" & ペアR存在
        Else
            Debug.Print "セットペア: [" & セットベース & "] → " & ペアF & ", " & ペアR
        End If
    Next key

    ' ==========================================
    ' 6. 転記先テーブル初期化
    ' ==========================================
    Application.StatusBar = "転記先を初期化中..."

    Dim tbl均し As ListObject
    Set tbl均し = ws均し.ListObjects("_成形展開均し")

    ' 転記データ（品番_日付 → 数量）
    Dim 転記データ As Object
    Set 転記データ = CreateObject("Scripting.Dictionary")

    ' 当日割り当て数記録（日付 → 累積数量）
    Dim 当日割り当て As Object
    Set 当日割り当て = CreateObject("Scripting.Dictionary")

    ' 型番ごとの当日割り当て記録（型番_日付 → 累積数量）
    Dim 当日型番割り当て As Object
    Set 当日型番割り当て = CreateObject("Scripting.Dictionary")

    ' グループ初回割り当て日記録（グループID → 日付）
    Dim グループ初回日 As Object
    Set グループ初回日 = CreateObject("Scripting.Dictionary")

    ' ==========================================
    ' 7. 日次割り当てループ
    ' ==========================================
    Dim 優先度 As Long, 稼働日 As Long, 割り当て As Long
    Dim グループID As String, 転記キー As String
    Dim wdIdx As Long

    For 優先度 = 1 To 3
        Application.StatusBar = "優先度" & 優先度 & "を処理中..."
        Debug.Print "--- 優先度" & 優先度 & " 処理開始 ---"

        ' ==========================================
        ' 7.1. 非セット品番の処理（従来通り）
        ' ==========================================
        Debug.Print "=== 非セット品番処理 ==="
        Dim key As Variant
        For Each key In 品番マスタ.Keys
            品番 = CStr(key)

            ' 優先度フィルタ
            If 品番マスタ(品番)("優先度") <> 優先度 Then GoTo NextItem1

            ' セット品はスキップ（後でペア単位で処理）
            If 品番マスタ(品番)("セット") = "SET" Then GoTo NextItem1

            ' 残数チェック
            If 月間残数(品番) = 0 Then GoTo NextItem1

            ' 稼働日ループ
            For wdIdx = 1 To workDays.Count
                稼働日 = CLng(workDays(wdIdx))

                ' グループ制約チェック（初回割り当て日に追従）
                グループID = 品番マスタ(品番)("グループ")
                Dim 対象稼働日 As Long
                対象稼働日 = 稼働日

                If グループID <> "" And グループ初回日.Exists(グループID) Then
                    Dim 初回日 As Long
                    初回日 = CLng(グループ初回日(グループID))

                    ' 初回日を優先的に試す
                    Dim 初回日割り当て As Long
                    初回日割り当て = 割り当て可能数を算出(品番, 初回日, 品番マスタ, 月間残数, 当日割り当て, 当日型番割り当て, 日次目標, 誤差許容率, 誤差許容個数)

                    If 初回日割り当て > 0 Then
                        対象稼働日 = 初回日
                    End If
                End If

                ' 割り当て可能数算出
                割り当て = 割り当て可能数を算出(品番, 対象稼働日, 品番マスタ, 月間残数, 当日割り当て, 当日型番割り当て, 日次目標, 誤差許容率, 誤差許容個数)

                If 割り当て > 0 Then
                    ' グループ初回日記録
                    If グループID <> "" And Not グループ初回日.Exists(グループID) Then
                        グループ初回日(グループID) = 対象稼働日
                        Debug.Print "グループ[" & グループID & "]初回日: " & 対象稼働日 & "日"
                    End If

                    ' 転記データ記録
                    転記キー = 品番 & "_" & 対象稼働日
                    If 転記データ.Exists(転記キー) Then
                        転記データ(転記キー) = CLng(転記データ(転記キー)) + 割り当て
                    Else
                        転記データ(転記キー) = 割り当て
                    End If

                    ' 残数更新
                    月間残数(品番) = CLng(月間残数(品番)) - 割り当て

                    ' 当日割り当て累積
                    If 当日割り当て.Exists(対象稼働日) Then
                        当日割り当て(対象稼働日) = CLng(当日割り当て(対象稼働日)) + 割り当て
                    Else
                        当日割り当て(対象稼働日) = 割り当て
                    End If

                    Debug.Print "品番[" & 品番 & "] " & 対象稼働日 & "日: " & 割り当て & "個 (残数: " & 月間残数(品番) & ")"

                    ' 残数ゼロなら次の品番へ
                    If 月間残数(品番) = 0 Then Exit For
                End If
            Next wdIdx

NextItem1:
        Next key

        ' ==========================================
        ' 7.2. セット品番の処理（ペア単位）
        ' ==========================================
        Debug.Print "=== セット品番処理（ペア単位） ==="
        For Each key In セットペアマスタ.Keys
            セットベース = CStr(key)
            Dim ペア品番群 As Object
            Set ペア品番群 = セットペアマスタ(セットベース)

            ' ペアの優先度チェック（代表品番で確認）
            Dim 代表品番 As String
            Dim ペア品番 As Variant
            For Each ペア品番 In ペア品番群.Keys
                代表品番 = CStr(ペア品番)
                Exit For
            Next ペア品番

            If 品番マスタ(代表品番)("優先度") <> 優先度 Then GoTo NextPair

            ' ペア全体の残数チェック
            Dim ペア残数あり As Boolean
            ペア残数あり = False
            For Each ペア品番 In ペア品番群.Keys
                If 月間残数(CStr(ペア品番)) > 0 Then
                    ペア残数あり = True
                    Exit For
                End If
            Next ペア品番
            If Not ペア残数あり Then GoTo NextPair

            ' グループ制約チェック（ペア全体で共通）
            グループID = 品番マスタ(代表品番)("グループ")

            ' 稼働日ループ
            For wdIdx = 1 To workDays.Count
                稼働日 = CLng(workDays(wdIdx))
                対象稼働日 = 稼働日

                If グループID <> "" And グループ初回日.Exists(グループID) Then
                    初回日 = CLng(グループ初回日(グループID))

                    ' 初回日を優先的に試す（ペア全体で）
                    Dim ペア初回日割り当て As Long
                    ペア初回日割り当て = セットペア割り当て可能数を算出(ペア品番群, 初回日, 品番マスタ, 月間残数, 当日割り当て, 当日型番割り当て, 日次目標, 誤差許容率, 誤差許容個数)

                    If ペア初回日割り当て > 0 Then
                        対象稼働日 = 初回日
                    End If
                End If

                ' セットペア割り当て可能数算出
                割り当て = セットペア割り当て可能数を算出(ペア品番群, 対象稼働日, 品番マスタ, 月間残数, 当日割り当て, 当日型番割り当て, 日次目標, 誤差許容率, 誤差許容個数)

                If 割り当て > 0 Then
                    ' グループ初回日記録
                    If グループID <> "" And Not グループ初回日.Exists(グループID) Then
                        グループ初回日(グループID) = 対象稼働日
                        Debug.Print "グループ[" & グループID & "]初回日: " & 対象稼働日 & "日"
                    End If

                    ' 全ペア品番に同数割り当て
                    For Each ペア品番 In ペア品番群.Keys
                        品番 = CStr(ペア品番)

                        ' 転記データ記録
                        転記キー = 品番 & "_" & 対象稼働日
                        転記データ(転記キー) = 割り当て

                        ' 残数更新
                        月間残数(品番) -= 割り当て

                        ' 当日割り当て累積（ペア品番ごとに加算）
                        If 当日割り当て.Exists(対象稼働日) Then
                            当日割り当て(対象稼働日) += 割り当て
                        Else
                            当日割り当て(対象稼働日) = 割り当て
                        End If

                        Debug.Print "セットペア[" & セットベース & "] 品番[" & 品番 & "] " & 対象稼働日 & "日: " & 割り当て & "個 (残数: " & 月間残数(品番) & ")"
                    Next ペア品番

                    ' 型番累積更新（ペア全体で）
                    Dim 型番 As String
                    型番 = CStr(品番マスタ(代表品番)("型番"))
                    If 型番 <> "" Then
                        Dim 型番日キー As String
                        型番日キー = 型番 & "_" & 対象稼働日
                        Dim ペア品番数 As Long
                        ペア品番数 = ペア品番群.Count

                        If 当日型番割り当て.Exists(型番日キー) Then
                            当日型番割り当て(型番日キー) += 割り当て * ペア品番数
                        Else
                            当日型番割り当て(型番日キー) = 割り当て * ペア品番数
                        End If
                        Debug.Print "  → 型番[" & 型番 & "] 累積: " & 当日型番割り当て(型番日キー) & "個 (ペア" & ペア品番数 & "件×" & 割り当て & "個)"
                    End If

                    ' ペア全体の残数チェック
                    Dim 全ペア完了 As Boolean
                    全ペア完了 = True
                    For Each ペア品番 In ペア品番群.Keys
                        If 月間残数(CStr(ペア品番)) > 0 Then
                            全ペア完了 = False
                            Exit For
                        End If
                    Next ペア品番
                    If 全ペア完了 Then Exit For
                End If
            Next wdIdx

NextPair:
        Next key
    Next 優先度

    ' ==========================================
    ' 8. 残数処理
    ' ==========================================
    If 月末処理モード = "自動" Then
        Application.StatusBar = "残数を処理中..."
        Debug.Print "--- 残数処理開始 ---"

        ' ==========================================
        ' 8.1. 非セット品番の残数処理
        ' ==========================================
        Debug.Print "=== 非セット品番残数処理 ==="
        For Each key In 月間残数.Keys
            品番 = CStr(key)

            ' セット品はスキップ（後でペア単位で処理）
            If 品番マスタ(品番)("セット") = "SET" Then GoTo NextItem2

            If CLng(月間残数(品番)) > 0 Then
                Debug.Print "残数あり: 品番[" & 品番 & "] " & 月間残数(品番) & "個"

                ' 月末稼働日から逆順
                For wdIdx = workDays.Count To 1 Step -1
                    稼働日 = CLng(workDays(wdIdx))

                    割り当て = 割り当て可能数を算出(品番, 稼働日, 品番マスタ, 月間残数, 当日割り当て, 当日型番割り当て, 日次目標, 誤差許容率, 誤差許容個数)

                    If 割り当て > 0 Then
                        転記キー = 品番 & "_" & 稼働日
                        If 転記データ.Exists(転記キー) Then
                            転記データ(転記キー) = CLng(転記データ(転記キー)) + 割り当て
                        Else
                            転記データ(転記キー) = 割り当て
                        End If

                        月間残数(品番) = CLng(月間残数(品番)) - 割り当て

                        If 当日割り当て.Exists(稼働日) Then
                            当日割り当て(稼働日) = CLng(当日割り当て(稼働日)) + 割り当て
                        Else
                            当日割り当て(稼働日) = 割り当て
                        End If

                        Debug.Print "  → " & 稼働日 & "日に" & 割り当て & "個追加 (残数: " & 月間残数(品番) & ")"

                        If CLng(月間残数(品番)) = 0 Then Exit For
                    End If
                Next wdIdx
            End If
NextItem2:
        Next key

        ' ==========================================
        ' 8.2. セット品番の残数処理（ペア単位）
        ' ==========================================
        Debug.Print "=== セット品番残数処理（ペア単位） ==="
        For Each key In セットペアマスタ.Keys
            セットベース = CStr(key)
            Set ペア品番群 = セットペアマスタ(セットベース)

            ' ペア全体の残数チェック
            ペア残数あり = False
            For Each ペア品番 In ペア品番群.Keys
                If 月間残数(CStr(ペア品番)) > 0 Then
                    ペア残数あり = True
                    Exit For
                End If
            Next ペア品番
            If Not ペア残数あり Then GoTo NextPair2

            Debug.Print "残数あり: セットペア[" & セットベース & "]"

            ' 月末稼働日から逆順
            For wdIdx = workDays.Count To 1 Step -1
                稼働日 = CLng(workDays(wdIdx))

                ' セットペア割り当て可能数算出
                割り当て = セットペア割り当て可能数を算出(ペア品番群, 稼働日, 品番マスタ, 月間残数, 当日割り当て, 当日型番割り当て, 日次目標, 誤差許容率, 誤差許容個数)

                If 割り当て > 0 Then
                    ' 全ペア品番に同数割り当て
                    For Each ペア品番 In ペア品番群.Keys
                        品番 = CStr(ペア品番)

                        転記キー = 品番 & "_" & 稼働日
                        If 転記データ.Exists(転記キー) Then
                            転記データ(転記キー) += 割り当て
                        Else
                            転記データ(転記キー) = 割り当て
                        End If

                        月間残数(品番) -= 割り当て

                        If 当日割り当て.Exists(稼働日) Then
                            当日割り当て(稼働日) += 割り当て
                        Else
                            当日割り当て(稼働日) = 割り当て
                        End If

                        Debug.Print "  → " & 稼働日 & "日に" & 割り当て & "個追加 品番[" & 品番 & "] (残数: " & 月間残数(品番) & ")"
                    Next ペア品番

                    ' 型番累積更新（ペア全体で）
                    For Each ペア品番 In ペア品番群.Keys
                        代表品番 = CStr(ペア品番)
                        Exit For
                    Next ペア品番

                    型番 = CStr(品番マスタ(代表品番)("型番"))
                    If 型番 <> "" Then
                        型番日キー = 型番 & "_" & 稼働日
                        ペア品番数 = ペア品番群.Count

                        If 当日型番割り当て.Exists(型番日キー) Then
                            当日型番割り当て(型番日キー) += 割り当て * ペア品番数
                        Else
                            当日型番割り当て(型番日キー) = 割り当て * ペア品番数
                        End If
                    End If

                    ' ペア全体の残数チェック
                    全ペア完了 = True
                    For Each ペア品番 In ペア品番群.Keys
                        If 月間残数(CStr(ペア品番)) > 0 Then
                            全ペア完了 = False
                            Exit For
                        End If
                    Next ペア品番
                    If 全ペア完了 Then Exit For
                End If
            Next wdIdx
NextPair2:
        Next key
    End If

    ' ==========================================
    ' 9. 転記データをテーブルに書き込み
    ' ==========================================
    Application.StatusBar = "転記データを書き込み中..."

    Dim arr均し As Variant
    ReDim arr均し(1 To UBound(arr展開, 1), 1 To UBound(arr展開, 2))

    ' 元データをコピー（成形品番等のメタデータ）
    For r = 1 To UBound(arr展開, 1)
        For i = 1 To UBound(arr展開, 2)
            If i < 開始列 Or i > 開始列 + maxDay - 1 Then
                arr均し(r, i) = arr展開(r, i)
            Else
                arr均し(r, i) = 0
            End If
        Next i
    Next r

    ' 転記データ反映
    For Each key In 転記データ.Keys
        Dim parts() As String
        parts = Split(CStr(key), "_")
        品番 = parts(0)
        Dim 日 As Long
        日 = CLng(parts(1))
        Dim 数量 As Long
        数量 = CLng(転記データ(key))

        ' 品番の行を探す
        For r = 1 To UBound(arr展開, 1)
            If CStr(arr展開(r, 成形品番列)) = 品番 Then
                arr均し(r, 開始列 + 日 - 1) = 数量
                Exit For
            End If
        Next r
    Next key

    ' テーブルに書き込み
    tbl均し.DataBodyRange.Value = arr均し

    ' ==========================================
    ' 10. 2日連続平均検証
    ' ==========================================
    Debug.Print "--- 2日連続平均検証 ---"

    For wdIdx = 1 To workDays.Count - 1
        Dim 日1 As Long, 日2 As Long
        日1 = CLng(workDays(wdIdx))
        日2 = CLng(workDays(wdIdx + 1))

        Dim 割り当て1 As Long, 割り当て2 As Long
        割り当て1 = 0
        割り当て2 = 0

        If 当日割り当て.Exists(日1) Then 割り当て1 = CLng(当日割り当て(日1))
        If 当日割り当て.Exists(日2) Then 割り当て2 = CLng(当日割り当て(日2))

        Dim 平均 As Double, 誤差 As Double, 許容 As Double
        平均 = (割り当て1 + 割り当て2) / 2
        誤差 = Abs(平均 - 日次目標)
        許容 = 日次目標 * 誤差許容率 / 100 + 誤差許容個数

        If 誤差 > 許容 Then
            Debug.Print "警告: " & 日1 & "-" & 日2 & "日の2日平均(" & Format(平均, "0.0") & ")が許容範囲外 (誤差:" & Format(誤差, "0.0") & ", 許容:" & Format(許容, "0.0") & ")"
        Else
            Debug.Print 日1 & "-" & 日2 & "日: OK (平均=" & Format(平均, "0.0") & ", 誤差=" & Format(誤差, "0.0") & ")"
        End If
    Next wdIdx

    Debug.Print "=== 負荷均し処理完了 ==="

    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "発生箇所: 転記_負荷均し" & vbCrLf & _
           "詳細はイミディエイトウィンドウを確認してください", vbCritical
    Exit Sub

ColumnError:
    Application.StatusBar = False
    MsgBox "列取得エラー: " & Err.Description & vbCrLf & _
           "必要な列がテーブルに存在しません" & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
End Sub

' ==========================================
' 割り当て可能数を算出
' ==========================================
Private Function 割り当て可能数を算出( _
    ByVal 品番 As String, _
    ByVal 稼働日 As Long, _
    ByRef 品番マスタ As Object, _
    ByRef 月間残数 As Object, _
    ByRef 当日割り当て As Object, _
    ByRef 当日型番割り当て As Object, _
    ByVal 日次目標 As Double, _
    ByVal 誤差許容率 As Double, _
    ByVal 誤差許容個数 As Long _
) As Long

    ' 基本制約
    Dim 残数 As Long, 上限 As Long, 単位 As Long
    残数 = CLng(月間残数(品番))
    上限 = CLng(品番マスタ(品番)("上限"))
    単位 = CLng(品番マスタ(品番)("単位"))

    ' 当日既割り当て数
    Dim 当日既割り当て As Long
    当日既割り当て = 0
    If 当日割り当て.Exists(稼働日) Then
        当日既割り当て = CLng(当日割り当て(稼働日))
    End If

    ' 日次目標制約（許容範囲考慮）
    Dim 許容上限 As Long
    許容上限 = CLng(日次目標 * (1 + 誤差許容率 / 100) + 誤差許容個数)

    Dim 当日最大 As Long
    当日最大 = 許容上限 - 当日既割り当て
    If 当日最大 > 上限 Then 当日最大 = 上限
    If 当日最大 < 0 Then 当日最大 = 0

    ' --- 型番制約チェック（セット品のみ） ---
    Dim セット As String
    セット = CStr(品番マスタ(品番)("セット"))

    If セット = "SET" Then
        Dim 型番 As String
        型番 = CStr(品番マスタ(品番)("型番"))

        If 型番 <> "" Then
            ' 型番上限（同一型番の品番は全て同じ上限値）
            Dim 型番上限 As Long
            型番上限 = CLng(品番マスタ(品番)("上限"))

            ' 型番の当日既割り当て数
            Dim 型番当日既割り当て As Long
            Dim 型番日キー As String
            型番日キー = 型番 & "_" & 稼働日

            If 当日型番割り当て.Exists(型番日キー) Then
                型番当日既割り当て = CLng(当日型番割り当て(型番日キー))
            Else
                型番当日既割り当て = 0
            End If

            ' 型番残余枠
            Dim 型番残余 As Long
            型番残余 = 型番上限 - 型番当日既割り当て

            If 型番残余 <= 0 Then
                ' 型番上限に達している
                Debug.Print "型番制約: [" & 型番 & "] " & 稼働日 & "日は上限到達 (上限: " & 型番上限 & ", 既割当: " & 型番当日既割り当て & ")"
                割り当て可能数を算出 = 0
                Exit Function
            End If

            ' 当日最大を型番制約も考慮
            If 当日最大 > 型番残余 Then
                Debug.Print "型番制約: [" & 型番 & "] " & 稼働日 & "日の残余 " & 型番残余 & "個に制限"
                当日最大 = 型番残余
            End If
        End If
    End If
    ' --- 型番制約チェック終わり ---

    ' 単位制約（倍数に丸める）
    Dim 割り当て候補 As Long
    割り当て候補 = 残数
    If 割り当て候補 > 当日最大 Then 割り当て候補 = 当日最大

    割り当て候補 = Int(割り当て候補 / 単位) * 単位

    割り当て可能数を算出 = 割り当て候補
End Function

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

    ' 見つからない場合はエラー
    Err.Raise vbObjectError + 1, "GetColumnIndex", "列[" & colName & "]が見つかりません"
End Function

' ==========================================
' Dictionary値を配列に変換
' ==========================================
Private Function DictValuesToArray(ByRef dict As Object) As Variant
    Dim arr() As Variant
    ReDim arr(1 To dict.Count)

    Dim i As Long
    i = 1
    Dim key As Variant
    For Each key In dict.Keys
        arr(i) = dict(key)
        i = i + 1
    Next key

    DictValuesToArray = arr
End Function

' ==========================================
' セットベース取得（末尾1文字除去）
' ==========================================
Private Function GetSetBase(ByVal 品番 As String) As String
    If Len(品番) > 0 Then
        GetSetBase = Left(品番, Len(品番) - 1)
    Else
        GetSetBase = ""
    End If
End Function

' ==========================================
' セットペア割り当て可能数算出
' ==========================================
Private Function セットペア割り当て可能数を算出( _
    ByRef ペア品番群 As Object, _
    ByVal 稼働日 As Long, _
    ByRef 品番マスタ As Object, _
    ByRef 月間残数 As Object, _
    ByRef 当日割り当て As Object, _
    ByRef 当日型番割り当て As Object, _
    ByVal 日次目標 As Double, _
    ByVal 誤差許容率 As Double, _
    ByVal 誤差許容個数 As Long _
) As Long

    Dim 最小割り当て As Long
    最小割り当て = 999999

    ' 各ペア品番の割り当て可能数を個別算出（型番制約含む）
    Dim ペア品番 As Variant
    For Each ペア品番 In ペア品番群.Keys
        Dim 個別割り当て As Long
        個別割り当て = 割り当て可能数を算出( _
            CStr(ペア品番), 稼働日, 品番マスタ, 月間残数, _
            当日割り当て, 当日型番割り当て, _
            日次目標, 誤差許容率, 誤差許容個数)

        If 個別割り当て < 最小割り当て Then
            最小割り当て = 個別割り当て
        End If
    Next ペア品番

    セットペア割り当て可能数を算出 = 最小割り当て
End Function
