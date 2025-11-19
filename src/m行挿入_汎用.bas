Attribute VB_Name = "m行挿入_汎用"

' ========================================
' マクロ名: 選択行挿入
' 処理概要: 指定テーブル内で選択行を上から挿入し、同数だけテーブル末尾を削除する
' 引数: tblName - 対象とするテーブル(リストオブジェクト)名
' 備考: シート保護を解除してから実行し、完了後に再保護する
' ========================================
Sub 選択行挿入(tblName As String)

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' ============================================
    ' アクティブセルとスクロール位置の記憶
    ' ============================================
    Dim activeCellAddress As String
    activeCellAddress = ActiveCell.Address

    Dim topRow As Long
    topRow = ActiveWindow.ScrollRow

    ' ============================================
    ' シート保護状況の確認と一時解除
    ' ============================================
    Dim isProtected As Boolean
    isProtected = (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios)

    Dim userInterfaceOnly As Boolean
    userInterfaceOnly = ws.ProtectionMode

    ' パスワードがある場合は、ここで設定しておく
    ' Dim pw As String
    ' pw = "password"

    ' 保護状況詳細の保持
    Dim allowFormattingCells As Boolean, allowFormattingColumns As Boolean, allowFormattingRows As Boolean
    Dim allowEditingObjects As Boolean
    Dim allowInsertingColumns As Boolean, allowInsertingRows As Boolean, allowInsertingHyperlinks As Boolean
    Dim allowDeletingColumns As Boolean, allowDeletingRows As Boolean, allowSorting As Boolean
    Dim allowFiltering As Boolean, allowUsingPivotTables As Boolean

    If isProtected Then
        With ws.Protection
            allowEditingObjects = Not ws.ProtectDrawingObjects
            allowFormattingCells = .allowFormattingCells
            allowFormattingColumns = .allowFormattingColumns
            allowFormattingRows = .allowFormattingRows
            allowInsertingColumns = .allowInsertingColumns
            allowInsertingRows = .allowInsertingRows
            allowInsertingHyperlinks = .allowInsertingHyperlinks
            allowDeletingColumns = .allowDeletingColumns
            allowDeletingRows = .allowDeletingRows
            allowSorting = .allowSorting
            allowFiltering = .allowFiltering
            allowUsingPivotTables = .allowUsingPivotTables
        End With

        ' シートの保護を解除 (パスワードがあるなら ws.Unprotect Password:=pw)
        ws.Unprotect
    End If

    ' ============================================
    ' テーブルと選択範囲の判定・行挿入
    ' ============================================
    Dim tbl As ListObject

    ' 引数で指定されたテーブルが存在するか確認
    On Error Resume Next
    Set tbl = ws.ListObjects(tblName)
    On Error GoTo 0

    If tbl Is Nothing Then
        MsgBox "テーブル名 """ & tblName & """ がアクティブシート上に見つかりません。", vbExclamation
        GoTo ExitProcedure
    End If

    ' 選択範囲が指定テーブルに属しているか（交差判定）
    If Intersect(Selection, tbl.Range) Is Nothing Then
        MsgBox "指定されたテーブル """ & tblName & """ 内のセルを選択してください。", vbExclamation
        GoTo ExitProcedure
    End If

    Dim firstRow As Long, lastRow As Long, rowsToInsert As Long
    Dim firstCol As Long, lastCol As Long

    firstRow = Selection.Row
    lastRow = Selection.Rows(Selection.Rows.Count).Row
    firstCol = Selection.Column
    lastCol = Selection.Columns(Selection.Columns.Count).Column

    rowsToInsert = lastRow - firstRow + 1

    ' ============================================
    ' テーブル上側への行挿入とテーブル末尾からの行削除
    ' ============================================

    ' 行をテーブルの上側に挿入
    Dim i As Long
    For i = 1 To rowsToInsert
        If (firstRow - tbl.Range.Row) > 0 Then
            tbl.ListRows.Add (firstRow - tbl.Range.Row)
        Else
            tbl.ListRows.Add 1
        End If
    Next i

    ' テーブル末尾から同数の行を削除
    For i = 1 To rowsToInsert
        If tbl.ListRows.Count > 0 Then
            tbl.ListRows(tbl.ListRows.Count).Delete
        End If
    Next i

    ' 挿入された範囲を再選択
    ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol)).Select

ExitProcedure:

    ' ============================================
    ' シートの再保護
    ' ============================================
    If isProtected Then
        ws.Protect DrawingObjects:=Not allowEditingObjects, _
                   Contents:=True, _
                   Scenarios:=True, _
                   userInterfaceOnly:=userInterfaceOnly, _
                   allowFormattingCells:=allowFormattingCells, _
                   allowFormattingColumns:=allowFormattingColumns, _
                   allowFormattingRows:=allowFormattingRows, _
                   allowInsertingColumns:=allowInsertingColumns, _
                   allowInsertingRows:=allowInsertingRows, _
                   allowInsertingHyperlinks:=allowInsertingHyperlinks, _
                   allowDeletingColumns:=allowDeletingColumns, _
                   allowDeletingRows:=allowDeletingRows, _
                   allowSorting:=allowSorting, _
                   allowFiltering:=allowFiltering, _
                   allowUsingPivotTables:=allowUsingPivotTables
    End If

    ' ============================================
    ' アクティブセルとスクロール位置の復元
    ' ============================================
    ActiveWindow.ScrollRow = topRow

    On Error Resume Next
    ws.Range(activeCellAddress).Select
    On Error GoTo 0

End Sub
