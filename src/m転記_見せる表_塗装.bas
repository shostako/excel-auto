Attribute VB_Name = "m転記_見せる表_塗装"
Option Explicit

Sub 転記_見せる表_塗装()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.StatusBar = "見せる表を作成中..."

    ' 1. 期間情報の取得
    Dim 期間Tbl As ListObject
    Set 期間Tbl = Sheets("塗装G").ListObjects("_集計期間塗装G")

    Dim 開始日 As Date, 終了日 As Date
    Dim i As Long

    ' 期間="期間1"の行を検索
    For i = 1 To 期間Tbl.ListRows.Count
        If 期間Tbl.ListColumns("期間").DataBodyRange(i, 1).Value = "期間1" Then
            開始日 = 期間Tbl.ListColumns("開始日").DataBodyRange(i, 1).Value
            終了日 = 期間Tbl.ListColumns("終了日").DataBodyRange(i, 1).Value
            Exit For
        End If
    Next i

    ' 2. シート名決定と既存シート削除
    Dim newName As String
    newName = "塗装" & Format(開始日, "M.d") & "～" & Format(終了日, "M.d")

    ' 同名シートが存在する場合は削除（確認なし）
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = Sheets(newName)
    On Error GoTo ErrorHandler
    If Not sh Is Nothing Then
        Application.DisplayAlerts = False
        sh.Delete
        Application.DisplayAlerts = True
        Set sh = Nothing
    End If

    ' 原紙シートをコピー
    Sheets("原紙塗装").Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = newName
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' 3. ソーステーブルからデータ取得
    Dim srcTbl As ListObject
    Set srcTbl = Sheets("塗装G").ListObjects("_流出G_塗装_期間1")

    ' 4. 不良項目数のカウント
    Dim 不良項目数 As Long
    不良項目数 = 0
    Dim item As String

    For i = 1 To srcTbl.ListRows.Count
        item = srcTbl.ListColumns("項目").DataBodyRange(i, 1).Value
        If Trim(item) <> "ショット数" And Trim(item) <> "不良数" Then
            不良項目数 = 不良項目数 + 1
        End If
    Next i

    Application.StatusBar = "不良項目数: " & 不良項目数 & "件"

    ' 5. 余分な行を削除（41行目から）
    Dim deleteCount As Long
    deleteCount = 20 - 不良項目数

    If deleteCount > 0 Then
        ws.Rows("41:" & (40 + deleteCount)).Delete Shift:=xlUp
    End If

    ' 6. 列マッピング設定
    ' 不良項目用（40行目以降）
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    colMap("合計") = "G"
    colMap("58050FrLH") = "L"
    colMap("58050FrRH") = "O"
    colMap("58050RrLH") = "R"
    colMap("58050RrRH") = "U"
    colMap("28050FrLH") = "X"
    colMap("28050FrRH") = "AA"
    colMap("28050RrLH") = "AD"
    colMap("28050RrRH") = "AG"
    colMap("補給品") = "AJ"

    ' 不良数行専用（39行目）- セル結合が異なるため
    Dim colMap不良数 As Object
    Set colMap不良数 = CreateObject("Scripting.Dictionary")
    colMap不良数("合計") = "D"
    colMap不良数("58050FrLH") = "K"
    colMap不良数("58050FrRH") = "N"
    colMap不良数("58050RrLH") = "Q"
    colMap不良数("58050RrRH") = "T"
    colMap不良数("28050FrLH") = "W"
    colMap不良数("28050FrRH") = "Z"
    colMap不良数("28050RrLH") = "AC"
    colMap不良数("28050RrRH") = "AF"
    colMap不良数("補給品") = "AI"

    ' 7. データ転記
    Dim targetRow As Long
    targetRow = 40

    For i = 1 To srcTbl.ListRows.Count
        item = srcTbl.ListColumns("項目").DataBodyRange(i, 1).Value
        Debug.Print "行" & i & ": [" & item & "] (長さ:" & Len(item) & ")"

        If Trim(item) = "不良数" Then
            Debug.Print "→ 不良数を39行目に転記"
            Call 転記行データ(ws, 39, srcTbl, i, colMap不良数)
        ElseIf Trim(item) <> "ショット数" And Trim(item) <> "不良数" Then
            Debug.Print "→ 不良項目を" & targetRow & "行目に転記"
            ' 不良項目を40行目以降に転記
            ws.Range("E" & targetRow).Value = item
            Call 転記行データ(ws, targetRow, srcTbl, i, colMap)
            targetRow = targetRow + 1
        Else
            Debug.Print "→ スキップ"
        End If
    Next i

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
End Sub

Private Sub 転記行データ(ws As Worksheet, targetRow As Long, _
                        srcTbl As ListObject, srcRowIndex As Long, _
                        colMap As Object)
    ' colMapに従ってデータを転記
    Dim key As Variant
    Dim val As Variant
    Debug.Print "  転記行データ開始: 行" & targetRow
    For Each key In colMap.Keys
        val = srcTbl.ListColumns(key).DataBodyRange(srcRowIndex, 1).Value
        ws.Range(colMap(key) & targetRow).Value = val
        Debug.Print "    " & key & " → " & colMap(key) & targetRow & " = " & val
    Next key
End Sub
