Attribute VB_Name = "m側板フィルター"
Option Explicit

' ========================================
' マクロ名: 側板フィルター
' 処理概要: C3セルの値と一致する「側板」列の行のみ表示
' 参照セル: C3（フィルター条件）
' フィルター対象: _完成品テーブルの「側板」列のみ
' 特殊動作: 他のテーブル（_core, _slitter, _acf）はフィルター解除
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
    Dim i As Long
    Dim startRow As Long

    ' 他テーブル（フィルター解除対象）
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

    ' --------------------------------------------
    ' 排他処理：他のフィルター参照セルをクリア
    ' --------------------------------------------
    ws.Range("B3").Value = ""
    ws.Range("D3").Value = ""
    ws.Range("E3").Value = "全項目"

    ' --------------------------------------------
    ' _完成品テーブル：側板列でフィルター
    ' --------------------------------------------
    Set tbl = ws.ListObjects("_完成品")
    startRow = tbl.DataBodyRange.Row
    tbl.DataBodyRange.EntireRow.Hidden = False
    dataArr = tbl.ListColumns("側板").DataBodyRange.Value
    For i = 1 To UBound(dataArr, 1)
        If dataArr(i, 1) <> filterValue Then
            ws.Rows(startRow + i - 1).Hidden = True
        End If
    Next i

    ' --------------------------------------------
    ' 他テーブル：フィルター解除
    ' --------------------------------------------
    Dim tblName As Variant
    For Each tblName In otherTables
        Set tbl = ws.ListObjects(CStr(tblName))
        tbl.DataBodyRange.EntireRow.Hidden = False
    Next tblName

    ' --------------------------------------------
    ' スクロール位置を先頭に移動
    ' --------------------------------------------
    Application.Goto ws.Range("A1"), True

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
