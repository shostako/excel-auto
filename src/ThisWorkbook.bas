Attribute VB_Name = "ThisWorkbook"
' ========================================
' モジュール名: ThisWorkbook
' 処理概要: ブック全体のイベント処理
' 設置場所: VBE → ThisWorkbook（シートモジュールではなく）
' ========================================

Option Explicit

' --------------------------------------------
' 全シート共通：セル変更時のフィルター自動実行
' B3, C3, D3, E3の変更を検知してフィルターマクロを呼び出す
' --------------------------------------------
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' B3:E3の範囲と交差するか判定
    If Not Intersect(Target, Sh.Range("B3:E3")) Is Nothing Then
        Application.EnableEvents = False  ' 無限ループ防止

        Select Case Target.Address
            Case "$B$3"
                If Len(Target.Value) > 0 Then Call 完成品フィルター
            Case "$C$3"
                If Len(Target.Value) > 0 Then Call 側板フィルター
            Case "$D$3"
                If Len(Target.Value) > 0 Then Call 小部品フィルター
            Case "$E$3"
                Call 項目フィルター  ' 「全項目」も含めて常に実行
        End Select

        Application.EnableEvents = True
    End If
End Sub
