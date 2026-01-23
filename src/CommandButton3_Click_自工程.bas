Attribute VB_Name = "CommandButton3_Click_自工程"
' ========================================
' コマンドボタン用マクロ（シートモジュールにコピペ）
' 処理概要: モードフィルタ適用 + グラフ軸自動調整
' 対象シート: ゾーンFrRr自工程
' 依存モジュール: mグラフ軸設定（先にインポートしておくこと）
' ========================================

Private Sub CommandButton3_Click()
    ' モードフィルタ適用 + グラフ軸調整マクロ

    Dim ws As Worksheet
    Dim selectedMode As String
    Dim pt As PivotTable
    Dim ptArray As Variant
    Dim chartArray As Variant
    Dim i As Long
    Dim targetSheetNames As Variant
    Dim sheetName As Variant

    targetSheetNames = Array("ゾーンFrRr自工程")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    ' --- シート保護解除 ---
    For Each sheetName In targetSheetNames
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(sheetName))
        If Err.Number = 0 Then
            ws.Unprotect Password:=""
        End If
        Err.Clear
    Next sheetName
    On Error GoTo ErrorHandler

    Set ws = ThisWorkbook.Worksheets("ゾーンFrRr自工程")
    selectedMode = ws.Range("T3").Value

    If selectedMode = "" Or selectedMode = "モード項目なし" Then
        MsgBox "モードが選択されていません。", vbExclamation
        GoTo Cleanup
    End If

    ' ============================================
    ' モードフィルタ適用
    ' ============================================
    ptArray = Array("ピボットテーブル41", "ピボットテーブル42", _
                    "ピボットテーブル43", "ピボットテーブル44")

    For i = 0 To UBound(ptArray)
        Set pt = ws.PivotTables(ptArray(i))

        On Error Resume Next
        With pt.PivotFields("モード2")
            .ClearAllFilters
            .CurrentPage = selectedMode
        End With
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo ErrorHandler
    Next i

    ' ============================================
    ' グラフ軸自動調整（共通モジュール呼び出し）
    ' ============================================
    chartArray = Array("グラフ1", "グラフ2", "グラフ3", "グラフ4")
    Call ApplyChartAxisFromPivots(ws, ptArray, chartArray)

Cleanup:
    ' --- シート再保護 ---
    For Each sheetName In targetSheetNames
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(CStr(sheetName))
        If Err.Number = 0 Then
            ws.Protect Password:="", DrawingObjects:=True, Contents:=True, Scenarios:=True
        End If
        Err.Clear
    Next sheetName
    On Error GoTo 0

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    Resume Cleanup
End Sub
