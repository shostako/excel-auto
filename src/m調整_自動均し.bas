Attribute VB_Name = "m調整_自動均し"
Option Explicit

' ==========================================
' 自動均し調整マクロ（改良版）
' ==========================================
' 目的: 均しマクロ実行後の結果を段階的に改善
' アプローチ: 最悪日優先・単純貪欲法
'   - 毎回、最も過剰な日と最も過少な日を特定
'   - その間で1品番のみ移動
'   - 移動後の改善効果を事前評価（往復移動防止）
'   - 全稼働日が平均±20%以内で収束
' 使い方: 何度もボタンを押して徐々に改善
' ==========================================

Sub 自動均し調整()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.StatusBar = "自動均し調整を開始します..."

    Debug.Print "========================================="
    Debug.Print "自動均し調整開始"
    Debug.Print "========================================="

    ' シート参照
    Dim ws均し As Worksheet, ws品番 As Worksheet, ws展開 As Worksheet
    Set ws均し = ThisWorkbook.Sheets("均し")
    Set ws品番 = ThisWorkbook.Sheets("品番")
    Set ws展開 = ThisWorkbook.Sheets("展開")

    ' テーブル参照
    Dim tbl均し As ListObject, tbl品番 As ListObject
    Set tbl均し = ws均し.ListObjects("_成形展開均し")
    Set tbl品番 = ws品番.ListObjects("_品番")

    ' 対象年月取得
    Dim 対象年月 As Date
    対象年月 = CDate(ws展開.Range("A3").Value)

    Dim 対象年 As Long, 対象月 As Long
    対象年 = Year(対象年月)
    対象月 = Month(対象年月)

    Dim maxDay As Long
    maxDay = Day(DateSerial(対象年, 対象月 + 1, 0))

    ' ==========================================
    ' 1. 稼働日リスト作成
    ' ==========================================
    Dim workDays As Object
    Set workDays = 稼働日リスト作成(対象年, 対象月, maxDay)

    Debug.Print "稼働日数: " & workDays.Count

    ' ==========================================
    ' 2. 日次合計算出
    ' ==========================================
    Dim 日次合計 As Object
    Set 日次合計 = 日次合計算出(tbl均し, maxDay)

    ' ==========================================
    ' 3. 日次平均算出
    ' ==========================================
    Dim 日次平均 As Double
    日次平均 = 日次平均算出(日次合計, workDays)

    Debug.Print "日次平均: " & Format(日次平均, "#,##0.0") & "個"

    ' ==========================================
    ' 4. 各稼働日の平均乖離を計算
    ' ==========================================
    Dim 乖離 As Object
    Set 乖離 = CreateObject("Scripting.Dictionary")

    Dim wdIdx As Long, 稼働日 As Long, 数量 As Long, 平均比率 As Double

    For wdIdx = 1 To workDays.Count
        稼働日 = CLng(workDays(wdIdx))
        数量 = CLng(日次合計(稼働日))
        乖離(稼働日) = Abs(数量 - 日次平均)
    Next wdIdx

    ' ==========================================
    ' 5. 最悪日ペアを特定
    ' ==========================================
    Dim 最過剰日 As Long, 最過少日 As Long
    Dim 最大過剰乖離 As Double, 最大過少乖離 As Double
    最大過剰乖離 = 0
    最大過少乖離 = 0

    For wdIdx = 1 To workDays.Count
        稼働日 = CLng(workDays(wdIdx))
        数量 = CLng(日次合計(稼働日))

        ' 過剰日（平均超）
        If 数量 > 日次平均 Then
            If 乖離(稼働日) > 最大過剰乖離 Then
                最大過剰乖離 = 乖離(稼働日)
                最過剰日 = 稼働日
            End If
        End If

        ' 過少日（平均未満）
        If 数量 < 日次平均 Then
            If 乖離(稼働日) > 最大過少乖離 Then
                最大過少乖離 = 乖離(稼働日)
                最過少日 = 稼働日
            End If
        End If
    Next wdIdx

    ' ==========================================
    ' 6. 収束判定
    ' ==========================================
    Dim 過剰日数量 As Long, 過少日数量 As Long
    過剰日数量 = CLng(日次合計(最過剰日))
    過少日数量 = CLng(日次合計(最過少日))

    Dim 過剰日平均比率 As Double, 過少日平均比率 As Double
    過剰日平均比率 = 過剰日数量 / 日次平均 * 100
    過少日平均比率 = 過少日数量 / 日次平均 * 100

    Debug.Print ""
    Debug.Print "最過剰日: " & 最過剰日 & "日 (" & Format(過剰日数量, "#,##0") & "個, 平均比" & Format(過剰日平均比率, "0.0") & "%)"
    Debug.Print "最過少日: " & 最過少日 & "日 (" & Format(過少日数量, "#,##0") & "個, 平均比" & Format(過少日平均比率, "0.0") & "%)"

    ' 収束判定: 全稼働日が平均±20%以内
    If 過剰日平均比率 <= 120 And 過少日平均比率 >= 80 Then
        Debug.Print ""
        Debug.Print "========================================="
        Debug.Print "改善完了: 全稼働日が平均±20%以内です"
        Debug.Print "========================================="

        Application.StatusBar = False
        MsgBox "改善完了しました！" & vbCrLf & vbCrLf & _
               "全稼働日が平均±20%以内に収まっています。" & vbCrLf & _
               "分析マクロで詳細を確認してください。", vbInformation
        Exit Sub
    End If

    ' ==========================================
    ' 7. 最適品番選択と移動
    ' ==========================================
    Debug.Print ""
    Debug.Print "--- 品番選択 ---"

    Dim arr均し As Variant
    arr均し = tbl均し.DataBodyRange.Value

    Dim 成形品番列 As Long, 開始列 As Long
    成形品番列 = GetColumnIndex(tbl均し, "成形品番")
    開始列 = GetColumnIndex(tbl均し, "1")

    ' 過剰日の品番を数量昇順でソート（小さい品番から試す）
    Dim 品番リスト As Object
    Set 品番リスト = CreateObject("Scripting.Dictionary")

    Dim r As Long, 品番 As String, 品番数量 As Long

    For r = 1 To UBound(arr均し, 1)
        品番 = CStr(arr均し(r, 成形品番列))

        If 開始列 + 最過剰日 - 1 <= UBound(arr均し, 2) Then
            品番数量 = 0
            On Error Resume Next
            品番数量 = CLng(arr均し(r, 開始列 + 最過剰日 - 1))
            On Error GoTo ErrorHandler

            If 品番数量 > 0 Then
                ' セット品番は移動しない（複雑すぎるため）
                Dim arr品番 As Variant
                arr品番 = tbl品番.DataBodyRange.Value

                Dim 品番_成形品番列 As Long, 品番_セット列 As Long
                品番_成形品番列 = GetColumnIndex(tbl品番, "成形品番")
                品番_セット列 = GetColumnIndex(tbl品番, "セット")

                Dim セットフラグ As String
                セットフラグ = ""

                Dim i As Long
                For i = 1 To UBound(arr品番, 1)
                    If CStr(arr品番(i, 品番_成形品番列)) = 品番 Then
                        セットフラグ = CStr(arr品番(i, 品番_セット列))
                        Exit For
                    End If
                Next i

                If セットフラグ <> "SET" Then
                    品番リスト(品番) = 品番数量
                End If
            End If
        End If
    Next r

    If 品番リスト.Count = 0 Then
        Debug.Print "移動可能な品番がありません（セット品番のみ）"
        Application.StatusBar = False
        MsgBox "移動可能な品番がありません（過剰日にセット品番のみ存在）", vbExclamation
        Exit Sub
    End If

    ' 数量昇順でソート
    Dim ソート済品番リスト As Object
    Set ソート済品番リスト = 品番を数量順でソート(品番リスト)

    Debug.Print "移動候補品番: " & ソート済品番リスト.Count & "件"

    ' 最適品番を探索
    Dim 移動品番 As String
    Dim 移動数量 As Long
    Dim 見つかった As Boolean
    見つかった = False

    Dim key As Variant
    For Each key In ソート済品番リスト.Keys
        品番 = CStr(key)
        品番数量 = CLng(ソート済品番リスト(品番))

        ' 移動後シミュレーション
        Dim 移動後_過剰日数量 As Long, 移動後_過少日数量 As Long
        移動後_過剰日数量 = 過剰日数量 - 品番数量
        移動後_過少日数量 = 過少日数量 + 品番数量

        Dim 移動後_過剰日乖離 As Double, 移動後_過少日乖離 As Double
        移動後_過剰日乖離 = Abs(移動後_過剰日数量 - 日次平均)
        移動後_過少日乖離 = Abs(移動後_過少日数量 - 日次平均)

        ' 両日とも改善するか判定
        If 移動後_過剰日乖離 < 乖離(最過剰日) And _
           移動後_過少日乖離 < 乖離(最過少日) Then

            Debug.Print ""
            Debug.Print "選択: 品番[" & 品番 & "] " & Format(品番数量, "#,##0") & "個"
            Debug.Print "  移動前: 過剰日=" & Format(過剰日数量, "#,##0") & "個, 過少日=" & Format(過少日数量, "#,##0") & "個"
            Debug.Print "  移動後: 過剰日=" & Format(移動後_過剰日数量, "#,##0") & "個, 過少日=" & Format(移動後_過少日数量, "#,##0") & "個"
            Debug.Print "  乖離改善: 過剰日 " & Format(乖離(最過剰日), "#,##0.0") & " → " & Format(移動後_過剰日乖離, "#,##0.0")
            Debug.Print "           過少日 " & Format(乖離(最過少日), "#,##0.0") & " → " & Format(移動後_過少日乖離, "#,##0.0")

            移動品番 = 品番
            移動数量 = 品番数量
            見つかった = True
            Exit For
        End If
    Next key

    If Not 見つかった Then
        Debug.Print ""
        Debug.Print "改善可能な品番が見つかりませんでした"
        Debug.Print "（全品番が移動後に悪化する可能性）"

        Application.StatusBar = False
        MsgBox "改善可能な品番が見つかりませんでした。" & vbCrLf & vbCrLf & _
               "これ以上の自動調整は困難です。" & vbCrLf & _
               "手動調整マクロ「m調整_グループ日程移動」の使用を検討してください。", vbInformation
        Exit Sub
    End If

    ' ==========================================
    ' 8. 品番移動実行
    ' ==========================================
    Debug.Print ""
    Debug.Print "--- 移動実行 ---"

    Application.StatusBar = "品番[" & 移動品番 & "]を移動中..."

    For r = 1 To UBound(arr均し, 1)
        If CStr(arr均し(r, 成形品番列)) = 移動品番 Then
            ' 移動元をゼロに
            If 開始列 + 最過剰日 - 1 <= UBound(arr均し, 2) Then
                arr均し(r, 開始列 + 最過剰日 - 1) = 0
            End If

            ' 移動先に追加
            If 開始列 + 最過少日 - 1 <= UBound(arr均し, 2) Then
                Dim 既存数量 As Long
                既存数量 = 0
                On Error Resume Next
                既存数量 = CLng(arr均し(r, 開始列 + 最過少日 - 1))
                On Error GoTo ErrorHandler

                arr均し(r, 開始列 + 最過少日 - 1) = 既存数量 + 移動数量
            End If

            Debug.Print "品番[" & 移動品番 & "]: " & 最過剰日 & "日(" & Format(移動数量, "#,##0") & "個) → " & 最過少日 & "日"
            Exit For
        End If
    Next r

    ' テーブルに書き込み
    tbl均し.DataBodyRange.Value = arr均し

    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "自動均し調整完了（1品番移動）"
    Debug.Print "========================================="

    Application.StatusBar = False
    MsgBox "調整完了: 品番[" & 移動品番 & "]を移動しました" & vbCrLf & vbCrLf & _
           最過剰日 & "日(" & Format(移動数量, "#,##0") & "個) → " & 最過少日 & "日" & vbCrLf & vbCrLf & _
           "さらに改善する場合は、再度このマクロを実行してください。", vbInformation

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
End Sub

' ==========================================
' 品番を数量順でソート
' ==========================================
Private Function 品番を数量順でソート(ByRef 品番リスト As Object) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    ' 数量を配列に変換
    Dim 数量配列() As Long
    ReDim 数量配列(1 To 品番リスト.Count)

    Dim 品番配列() As String
    ReDim 品番配列(1 To 品番リスト.Count)

    Dim idx As Long
    idx = 1

    Dim key As Variant
    For Each key In 品番リスト.Keys
        品番配列(idx) = CStr(key)
        数量配列(idx) = CLng(品番リスト(key))
        idx = idx + 1
    Next key

    ' バブルソート（数量昇順）
    Dim i As Long, j As Long
    Dim temp数量 As Long, temp品番 As String

    For i = 1 To 品番リスト.Count - 1
        For j = i + 1 To 品番リスト.Count
            If 数量配列(i) > 数量配列(j) Then
                ' 数量入れ替え
                temp数量 = 数量配列(i)
                数量配列(i) = 数量配列(j)
                数量配列(j) = temp数量

                ' 品番入れ替え
                temp品番 = 品番配列(i)
                品番配列(i) = 品番配列(j)
                品番配列(j) = temp品番
            End If
        Next j
    Next i

    ' 結果Dictionaryに格納
    For i = 1 To 品番リスト.Count
        result(品番配列(i)) = 数量配列(i)
    Next i

    Set 品番を数量順でソート = result
End Function

' ==========================================
' 稼働日リスト作成
' ==========================================
Private Function 稼働日リスト作成(ByVal 対象年 As Long, _
                                 ByVal 対象月 As Long, _
                                 ByVal maxDay As Long) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

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

    Dim d As Long, dt As Date, wd As Long, dateKey As Long
    For d = 1 To maxDay
        dt = DateSerial(対象年, 対象月, d)
        wd = Weekday(dt)
        dateKey = Year(dt) * 10000 + Month(dt) * 100 + Day(dt)

        If wd <> 1 And wd <> 7 And Not holidays.Exists(dateKey) Then
            result.Add result.Count + 1, d
        End If
    Next d

    Set 稼働日リスト作成 = result
End Function

' ==========================================
' 日次合計算出
' ==========================================
Private Function 日次合計算出(ByRef tbl均し As ListObject, _
                            ByVal maxDay As Long) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")

    Dim arr均し As Variant
    arr均し = tbl均し.DataBodyRange.Value

    Dim 開始列 As Long
    開始列 = GetColumnIndex(tbl均し, "1")

    Dim d As Long, r As Long, 数量 As Long

    For d = 1 To maxDay
        result(d) = 0
    Next d

    For r = 1 To UBound(arr均し, 1)
        For d = 1 To maxDay
            If 開始列 + d - 1 <= UBound(arr均し, 2) Then
                数量 = 0
                On Error Resume Next
                数量 = CLng(arr均し(r, 開始列 + d - 1))
                On Error GoTo 0

                result(d) = CLng(result(d)) + 数量
            End If
        Next d
    Next r

    Set 日次合計算出 = result
End Function

' ==========================================
' 日次平均算出
' ==========================================
Private Function 日次平均算出(ByRef 日次合計 As Object, _
                            ByRef workDays As Object) As Double
    Dim 稼働日合計 As Long
    稼働日合計 = 0

    Dim wdIdx As Long, 稼働日 As Long
    For wdIdx = 1 To workDays.Count
        稼働日 = CLng(workDays(wdIdx))
        稼働日合計 = 稼働日合計 + CLng(日次合計(稼働日))
    Next wdIdx

    日次平均算出 = 稼働日合計 / workDays.Count
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

    Err.Raise vbObjectError + 1, "GetColumnIndex", "列[" & colName & "]が見つかりません"
End Function
