Attribute VB_Name = "m調整_グループ日程移動"
Option Explicit

' ==========================================
' グループ日程移動マクロ
' ==========================================
' 目的: 均しマクロ実行後の手動微調整
' 使い方:
'   1. 均しマクロ実行後、このマクロを実行
'   2. グループIDを入力（例: BB）
'   3. 移動先日を選択（絶対日付 or 相対移動）
'   4. 警告を確認して決定
' ==========================================

Sub グループ日程移動()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.StatusBar = "グループ日程移動を開始します..."

    ' シート参照
    Dim ws均し As Worksheet, ws品番 As Worksheet
    Set ws均し = ThisWorkbook.Sheets("均し")
    Set ws品番 = ThisWorkbook.Sheets("品番")

    ' テーブル参照
    Dim tbl均し As ListObject, tbl品番 As ListObject
    Set tbl均し = ws均し.ListObjects("_成形展開均し")
    Set tbl品番 = ws品番.ListObjects("_品番")

    ' ==========================================
    ' 1. グループID入力
    ' ==========================================
    Dim グループID As String
    グループID = InputBox("移動するグループIDを入力してください" & vbCrLf & _
                          "（例: BB、CC、DD）", "グループ日程移動")

    If グループID = "" Then
        MsgBox "キャンセルされました", vbInformation
        Application.StatusBar = False
        Exit Sub
    End If

    ' ==========================================
    ' 2. 該当グループの品番を抽出
    ' ==========================================
    Dim arr品番 As Variant
    arr品番 = tbl品番.DataBodyRange.Value

    Dim 品番_成形品番列 As Long, 品番_グループ列 As Long
    品番_成形品番列 = GetColumnIndex(tbl品番, "成形品番")
    品番_グループ列 = GetColumnIndex(tbl品番, "グループ")

    Dim グループ品番リスト As Object
    Set グループ品番リスト = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 1 To UBound(arr品番, 1)
        Dim grpVal As Variant
        grpVal = arr品番(r, 品番_グループ列)

        If Not IsEmpty(grpVal) And Not IsNull(grpVal) Then
            If UCase(CStr(grpVal)) = UCase(グループID) Then
                Dim 成形品番 As String
                成形品番 = CStr(arr品番(r, 品番_成形品番列))
                グループ品番リスト(成形品番) = True
            End If
        End If
    Next r

    If グループ品番リスト.Count = 0 Then
        MsgBox "グループ[" & グループID & "]に該当する品番が見つかりません", vbExclamation
        Application.StatusBar = False
        Exit Sub
    End If

    Debug.Print "=== グループ日程移動 ==="
    Debug.Print "グループID: " & グループID
    Debug.Print "該当品番数: " & グループ品番リスト.Count

    ' ==========================================
    ' 3. 現在の割り当て日を特定
    ' ==========================================
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

    ' 現在の割り当て日を検出
    Dim 現在割り当て日 As Object
    Set 現在割り当て日 = CreateObject("Scripting.Dictionary")

    Dim key As Variant, 品番 As String, d As Long
    For Each key In グループ品番リスト.Keys
        品番 = CStr(key)

        For r = 1 To UBound(arr均し, 1)
            If CStr(arr均し(r, 成形品番列)) = 品番 Then
                For d = 1 To maxDay
                    If 開始列 + d - 1 <= UBound(arr均し, 2) Then
                        Dim 数量 As Long
                        数量 = 0
                        On Error Resume Next
                        数量 = CLng(arr均し(r, 開始列 + d - 1))
                        On Error GoTo ErrorHandler

                        If 数量 > 0 Then
                            If Not 現在割り当て日.Exists(d) Then
                                現在割り当て日(d) = 数量
                            Else
                                現在割り当て日(d) = CLng(現在割り当て日(d)) + 数量
                            End If
                        End If
                    End If
                Next d
                Exit For
            End If
        Next r
    Next key

    If 現在割り当て日.Count = 0 Then
        MsgBox "グループ[" & グループID & "]の割り当てが見つかりません", vbExclamation
        Application.StatusBar = False
        Exit Sub
    End If

    ' 現在の主要割り当て日（最も数量が多い日）
    Dim 主要日 As Long, 最大数量 As Long
    主要日 = 0
    最大数量 = 0

    For Each key In 現在割り当て日.Keys
        If CLng(現在割り当て日(key)) > 最大数量 Then
            最大数量 = CLng(現在割り当て日(key))
            主要日 = CLng(key)
        End If
    Next key

    Debug.Print "現在の主要日: " & 主要日 & "日 (" & 最大数量 & "個)"

    ' ==========================================
    ' 4. 移動先日の選択
    ' ==========================================
    Dim 選択 As Long
    選択 = MsgBox("現在の主要割り当て日: " & 主要日 & "日" & vbCrLf & vbCrLf & _
                  "絶対日付指定: [はい]" & vbCrLf & _
                  "相対移動: [いいえ]", vbYesNoCancel + vbQuestion, "移動先日の選択")

    If 選択 = vbCancel Then
        MsgBox "キャンセルされました", vbInformation
        Application.StatusBar = False
        Exit Sub
    End If

    Dim 移動先日 As Long

    If 選択 = vbYes Then
        ' 絶対日付指定
        Dim 入力日 As String
        入力日 = InputBox("移動先の日を入力してください（例: 5、12）", "絶対日付指定", 主要日)

        If 入力日 = "" Then
            MsgBox "キャンセルされました", vbInformation
            Application.StatusBar = False
            Exit Sub
        End If

        If Not IsNumeric(入力日) Then
            MsgBox "日付は数値で入力してください", vbExclamation
            Application.StatusBar = False
            Exit Sub
        End If

        移動先日 = CLng(入力日)
    Else
        ' 相対移動
        Dim ずらし日数 As String
        ずらし日数 = InputBox("何日ずらすか入力してください" & vbCrLf & _
                             "（例: +2、-3）", "相対移動", "+0")

        If ずらし日数 = "" Then
            MsgBox "キャンセルされました", vbInformation
            Application.StatusBar = False
            Exit Sub
        End If

        If Not IsNumeric(ずらし日数) Then
            MsgBox "日数は数値で入力してください（+2、-3など）", vbExclamation
            Application.StatusBar = False
            Exit Sub
        End If

        移動先日 = 主要日 + CLng(ずらし日数)
    End If

    ' 日付妥当性チェック
    If 移動先日 < 1 Or 移動先日 > maxDay Then
        MsgBox "移動先日が範囲外です（1～" & maxDay & "日）", vbExclamation
        Application.StatusBar = False
        Exit Sub
    End If

    ' 移動先日が休日かチェック
    Dim dt As Date, wd As Long
    dt = DateSerial(対象年, 対象月, 移動先日)
    wd = Weekday(dt)

    If wd = 1 Or wd = 7 Then
        Dim 休日確認 As Long
        休日確認 = MsgBox("移動先日（" & 移動先日 & "日）は土日です。" & vbCrLf & _
                         "それでも移動しますか？", vbYesNo + vbExclamation, "休日警告")
        If 休日確認 = vbNo Then
            MsgBox "キャンセルされました", vbInformation
            Application.StatusBar = False
            Exit Sub
        End If
    End If

    Debug.Print "移動先日: " & 移動先日 & "日"

    ' ==========================================
    ' 5. 移動実行
    ' ==========================================
    Application.StatusBar = "グループ[" & グループID & "]を" & 移動先日 & "日に移動中..."

    ' 該当品番の全日程をクリア
    For Each key In グループ品番リスト.Keys
        品番 = CStr(key)

        For r = 1 To UBound(arr均し, 1)
            If CStr(arr均し(r, 成形品番列)) = 品番 Then
                For d = 1 To maxDay
                    If 開始列 + d - 1 <= UBound(arr均し, 2) Then
                        arr均し(r, 開始列 + d - 1) = 0
                    End If
                Next d
                Exit For
            End If
        Next r
    Next key

    ' 移動先日に割り当て
    For Each key In 現在割り当て日.Keys
        Dim 元日 As Long
        元日 = CLng(key)
        数量 = CLng(現在割り当て日(元日))

        For Each 品番 In グループ品番リスト.Keys
            For r = 1 To UBound(arr均し, 1)
                If CStr(arr均し(r, 成形品番列)) = CStr(品番) Then
                    ' 元日の数量を取得（個別品番の数量）
                    Dim 品番数量 As Long
                    品番数量 = 0

                    ' 元データから元日の数量を取得
                    Dim arr元 As Variant
                    arr元 = tbl均し.DataBodyRange.Value

                    If 開始列 + 元日 - 1 <= UBound(arr元, 2) Then
                        On Error Resume Next
                        品番数量 = CLng(arr元(r, 開始列 + 元日 - 1))
                        On Error GoTo ErrorHandler
                    End If

                    ' 移動先日に割り当て
                    If 品番数量 > 0 And 開始列 + 移動先日 - 1 <= UBound(arr均し, 2) Then
                        arr均し(r, 開始列 + 移動先日 - 1) = 品番数量
                        Debug.Print "品番[" & 品番 & "]: " & 元日 & "日(" & 品番数量 & "個) → " & 移動先日 & "日"
                    End If

                    Exit For
                End If
            Next r
        Next 品番
    Next key

    ' テーブルに書き込み
    tbl均し.DataBodyRange.Value = arr均し

    ' ==========================================
    ' 6. 警告表示
    ' ==========================================
    Debug.Print "--- 移動後の再計算 ---"

    ' 日次合計を再計算
    Dim 日次合計 As Object
    Set 日次合計 = CreateObject("Scripting.Dictionary")

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

    ' 日次目標（簡易算出）
    Dim 全体合計 As Long
    全体合計 = 0
    For d = 1 To maxDay
        全体合計 = 全体合計 + CLng(日次合計(d))
    Next d

    ' 稼働日数（簡易: 土日除外のみ）
    Dim 稼働日数 As Long
    稼働日数 = 0
    For d = 1 To maxDay
        dt = DateSerial(対象年, 対象月, d)
        wd = Weekday(dt)
        If wd <> 1 And wd <> 7 Then 稼働日数 = 稼働日数 + 1
    Next d

    Dim 日次平均 As Double
    If 稼働日数 > 0 Then
        日次平均 = 全体合計 / 稼働日数
    Else
        日次平均 = 0
    End If

    Debug.Print "日次平均: " & Format(日次平均, "#,##0.0") & "個"

    ' 移動先日の警告チェック
    Dim 移動先数量 As Long
    移動先数量 = CLng(日次合計(移動先日))

    Dim 平均比率 As Double
    If 日次平均 > 0 Then
        平均比率 = 移動先数量 / 日次平均 * 100
    Else
        平均比率 = 0
    End If

    Debug.Print "移動先日(" & 移動先日 & "日): " & Format(移動先数量, "#,##0") & "個 (平均比: " & Format(平均比率, "0.0") & "%)"

    Dim 警告メッセージ As String
    警告メッセージ = ""

    If 平均比率 > 120 Then
        警告メッセージ = 警告メッセージ & "⚠ 移動先日が平均の120%超です（過剰）" & vbCrLf
        Debug.Print "警告: 移動先日が平均の120%超（過剰）"
    ElseIf 平均比率 < 80 Then
        警告メッセージ = 警告メッセージ & "⚠ 移動先日が平均の80%未満です（過少）" & vbCrLf
        Debug.Print "警告: 移動先日が平均の80%未満（過少）"
    End If

    Debug.Print "=== グループ日程移動完了 ==="

    Application.StatusBar = False

    Dim 完了メッセージ As String
    完了メッセージ = "グループ[" & グループID & "]を" & 主要日 & "日 → " & 移動先日 & "日に移動しました" & vbCrLf & vbCrLf & _
                    "移動先日: " & Format(移動先数量, "#,##0") & "個 (平均比: " & Format(平均比率, "0.0") & "%)"

    If 警告メッセージ <> "" Then
        完了メッセージ = 完了メッセージ & vbCrLf & vbCrLf & 警告メッセージ
    End If

    MsgBox 完了メッセージ, vbInformation, "グループ日程移動完了"
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
