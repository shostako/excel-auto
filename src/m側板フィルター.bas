Attribute VB_Name = "m側板フィルター"
Option Explicit

' ========================================
' マクロ名: 側板フィルター
' 処理概要: C3セルの値と一致する「側板」列の行のみ表示
' 参照セル: C3（フィルター条件）, E3（項目フィルター条件）
' フィルター対象: _完成品テーブルの「側板」列のみ
' 特殊動作: 他のテーブル（_core, _slitter, _acf）は項目フィルターのみ適用
' 複合フィルター: 項目フィルター（E3）の条件も同時に適用
' 特殊行: 「稼働日」行は常に表示
' 最適化: 配列一括読み込み + 計算/イベント抑制
' ========================================

Sub 側板フィルター()
    ' --------------------------------------------
    ' 変数宣言
    ' --------------------------------------------
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim filterValue As String
    Dim dataArr As Variant
    Dim itemArr As Variant
    Dim i As Long
    Dim startRow As Long

    ' 項目フィルター用
    Dim filterItem As String
    Dim filterMode As String

    ' 他テーブル（項目フィルターのみ適用）
    Dim otherTables As Variant
    otherTables = Array("_core", "_slitter", "_acf")

    ' --------------------------------------------
    ' 画面更新・計算・イベント抑制（高速化）
    ' --------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    Set ws = ActiveSheet

    ' --------------------------------------------
    ' フィルター条件の取得
    ' --------------------------------------------
    filterValue = ws.Range("C3").Value

    ' 項目フィルター条件の取得
    filterItem = ws.Range("E3").Value
    filterMode = GetItemFilterMode(filterItem)

    ' --------------------------------------------
    ' 排他処理：他のフィルター参照セルをクリア
    ' （E3項目フィルターは変更しない）
    ' --------------------------------------------
    ws.Range("B3").Value = ""
    ws.Range("D3").Value = ""

    ' --------------------------------------------
    ' _完成品テーブル：側板列 AND 項目列でフィルター
    ' --------------------------------------------
    Set tbl = FindTableByPattern(ws, "_完成品")
    startRow = tbl.DataBodyRange.Row
    tbl.DataBodyRange.EntireRow.Hidden = False
    dataArr = tbl.ListColumns("側板").DataBodyRange.Value
    itemArr = tbl.ListColumns("項目").DataBodyRange.Value
    For i = 1 To UBound(dataArr, 1)
        ' 「稼働日」行は常に表示
        If itemArr(i, 1) <> "稼働日" Then
            If dataArr(i, 1) <> filterValue Or Not MatchItemFilter(CStr(itemArr(i, 1)), filterMode, filterItem) Then
                ws.Rows(startRow + i - 1).Hidden = True
            End If
        End If
    Next i

    ' --------------------------------------------
    ' 他テーブル：項目フィルターのみ適用
    ' --------------------------------------------
    Dim tblName As Variant
    For Each tblName In otherTables
        Set tbl = FindTableByPattern(ws, CStr(tblName))
        startRow = tbl.DataBodyRange.Row
        tbl.DataBodyRange.EntireRow.Hidden = False
        itemArr = tbl.ListColumns("項目").DataBodyRange.Value
        For i = 1 To UBound(itemArr, 1)
            ' 「稼働日」行は常に表示
            If itemArr(i, 1) <> "稼働日" Then
                If Not MatchItemFilter(CStr(itemArr(i, 1)), filterMode, filterItem) Then
                    ws.Rows(startRow + i - 1).Hidden = True
                End If
            End If
        Next i
    Next tblName

    ' --------------------------------------------
    ' 垂直スクロールのみ先頭に移動（水平位置は維持）
    ' --------------------------------------------
    ActiveWindow.ScrollRow = 1

    ' --------------------------------------------
    ' 終了処理
    ' --------------------------------------------
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "エラーが発生しました" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "詳細: " & Err.Description, vbCritical
End Sub
