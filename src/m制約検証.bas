Attribute VB_Name = "m制約検証"
Option Explicit

' ==========================================
' 制約検証マクロ
' ==========================================
' 既存の負荷均しマクロ実行後に、制約違反をチェックして報告する
'
' 使い方:
' 1. 既存の「転記_負荷均し」を実行
' 2. このマクロを実行して制約違反をチェック
' 3. 違反がある場合は調整マクロで修正
' ==========================================

Sub 制約検証()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.StatusBar = "制約検証を開始します..."

    Debug.Print "========================================="
    Debug.Print "制約検証開始"
    Debug.Print "========================================="

    ' シート参照
    Dim ws品番 As Worksheet, ws均し As Worksheet
    Set ws品番 = ThisWorkbook.Sheets("品番")
    Set ws均し = ThisWorkbook.Sheets("均し")

    ' テーブル参照
    Dim tbl均し As ListObject, tbl品番 As ListObject
    Set tbl均し = ws均し.ListObjects("_成形展開均し")
    Set tbl品番 = ws品番.ListObjects("_品番")

    ' ==========================================
    ' 1. 品番マスタ読み込み
    ' ==========================================
    Application.StatusBar = "品番マスタを読み込み中..."

    Dim arr品番 As Variant
    arr品番 = tbl品番.DataBodyRange.Value

    Dim 品番マスタ As Object
    Set 品番マスタ = CreateObject("Scripting.Dictionary")

    ' 列インデックス取得
    Dim 品番_成形品番列 As Long, 品番_号補列 As Long
    Dim 品番_系列列 As Long, 品番_仕様列 As Long, 品番_セット列 As Long

    品番_成形品番列 = GetColumnIndex(tbl品番, "成形品番")
    品番_号補列 = GetColumnIndex(tbl品番, "号/補")
    品番_系列列 = GetColumnIndex(tbl品番, "系列")
    品番_仕様列 = GetColumnIndex(tbl品番, "仕様")
    品番_セット列 = GetColumnIndex(tbl品番, "セット")

    Dim r As Long
    For r = 1 To UBound(arr品番, 1)
        Dim info As Object
        Set info = CreateObject("Scripting.Dictionary")

        Dim 成形品番値 As Variant
        成形品番値 = arr品番(r, 品番_成形品番列)

        If IsEmpty(成形品番値) Or 成形品番値 = "" Then GoTo NextRow

        Dim 号補Val As Variant, 系列Val As Variant, 仕様Val As Variant, セットVal As Variant
        号補Val = arr品番(r, 品番_号補列)
        info("号/補") = IIf(IsEmpty(号補Val) Or IsNull(号補Val) Or 号補Val = "", "", CStr(号補Val))

        系列Val = arr品番(r, 品番_系列列)
        info("系列") = IIf(IsEmpty(系列Val) Or IsNull(系列Val) Or 系列Val = "", "", CStr(系列Val))

        仕様Val = arr品番(r, 品番_仕様列)
        info("仕様") = IIf(IsEmpty(仕様Val) Or IsNull(仕様Val) Or 仕様Val = "", "", CStr(仕様Val))

        セットVal = arr品番(r, 品番_セット列)
        info("セット") = IIf(IsEmpty(セットVal) Or IsNull(セットVal) Or セットVal = "", "", CStr(セットVal))

        Set 品番マスタ(CStr(成形品番値)) = info

NextRow:
    Next r

    Debug.Print "品番マスタ件数: " & 品番マスタ.Count

    ' ==========================================
    ' 2. 均しシートのデータ読み込み
    ' ==========================================
    Application.StatusBar = "均しデータを読み込み中..."

    Dim arr均し As Variant
    arr均し = tbl均し.DataBodyRange.Value

    Dim 成形品番列 As Long, 開始列 As Long
    成形品番列 = GetColumnIndex(tbl均し, "成形品番")
    開始列 = GetColumnIndex(tbl均し, "1")

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

    Debug.Print "対象年月: " & 対象年 & "/" & 対象月
    Debug.Print "最大日: " & maxDay

    ' ==========================================
    ' 3. 制約違反チェック
    ' ==========================================
    Application.StatusBar = "制約違反をチェック中..."

    Debug.Print ""
    Debug.Print "--- 制約違反チェック ---"

    Dim 違反カウント As Long
    違反カウント = 0

    Dim d As Long, 品番 As String
    Dim 号補 As String, セット As String
    Dim 数量 As Long

    ' 日ごとにチェック
    For d = 1 To maxDay
        Dim 当日号口単品リスト As Object, 当日補給品リスト As Object
        Set 当日号口単品リスト = CreateObject("Scripting.Dictionary")
        Set 当日補給品リスト = CreateObject("Scripting.Dictionary")

        ' 当日の品番を収集
        For r = 1 To UBound(arr均し, 1)
            品番 = CStr(arr均し(r, 成形品番列))

            If Not 品番マスタ.Exists(品番) Then GoTo NextItem

            ' 数量チェック
            数量 = 0
            On Error Resume Next
            数量 = CLng(arr均し(r, 開始列 + d - 1))
            On Error GoTo ErrorHandler

            If 数量 > 0 Then
                号補 = CStr(品番マスタ(品番)("号/補"))
                セット = CStr(品番マスタ(品番)("セット"))

                ' 号口単品
                If 号補 = "号口" And セット <> "SET" Then
                    当日号口単品リスト(品番) = 数量
                End If

                ' 補給品
                If 号補 = "補給品" Then
                    当日補給品リスト(品番) = 数量
                End If
            End If
NextItem:
        Next r

        ' 制約1: 号口単品が複数存在しないか
        If 当日号口単品リスト.Count > 1 Then
            違反カウント = 違反カウント + 1
            Debug.Print "【違反" & 違反カウント & "】" & d & "日: 号口単品が" & 当日号口単品リスト.Count & "件存在（制約：1日1件まで）"

            Dim key As Variant
            For Each key In 当日号口単品リスト.Keys
                Debug.Print "  - " & key & " (" & 当日号口単品リスト(key) & "個)"
            Next key
        End If

        ' 制約2: 補給品と号口単品が同日に存在しないか
        If 当日補給品リスト.Count > 0 And 当日号口単品リスト.Count > 0 Then
            違反カウント = 違反カウント + 1
            Debug.Print "【違反" & 違反カウント & "】" & d & "日: 補給品と号口単品が同日に存在（制約：同日配置禁止）"
            Debug.Print "  補給品: " & 当日補給品リスト.Count & "件"
            Debug.Print "  号口単品: " & 当日号口単品リスト.Count & "件"
        End If
    Next d

    ' ==========================================
    ' 4. 結果報告
    ' ==========================================
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "制約検証完了"
    Debug.Print "========================================="
    Debug.Print "違反件数: " & 違反カウント

    Application.StatusBar = False

    If 違反カウント = 0 Then
        MsgBox "制約違反はありません！" & vbCrLf & vbCrLf & _
               "全ての制約が守られています。", vbInformation, "制約検証完了"
    Else
        MsgBox "制約違反が " & 違反カウント & " 件見つかりました。" & vbCrLf & vbCrLf & _
               "詳細はイミディエイトウィンドウを確認してください。" & vbCrLf & vbCrLf & _
               "調整マクロで修正することを推奨します。", vbExclamation, "制約違反検出"
    End If

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

