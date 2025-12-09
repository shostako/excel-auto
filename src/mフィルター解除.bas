Attribute VB_Name = "mフィルター解除"
Option Explicit

' ========================================
' マクロ名: フィルター解除
' 処理概要: 全テーブルの行非表示を解除（全行を表示）
' 対象テーブル: _完成品, _core, _slitter, _acf
' ========================================

Sub フィルター解除()
    ' --------------------------------------------
    ' 変数宣言
    ' --------------------------------------------
    Dim ws As Worksheet
    Dim tbl As ListObject

    ' 対象テーブル名
    Dim tables As Variant
    tables = Array("_完成品", "_core", "_slitter", "_acf")

    ' --------------------------------------------
    ' 画面更新・イベント抑制（高速化）
    ' --------------------------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    Set ws = ActiveSheet

    ' --------------------------------------------
    ' 全テーブルの行を表示（フィルター解除）
    ' --------------------------------------------
    Dim tblName As Variant
    For Each tblName In tables
        Set tbl = FindTableByPattern(ws, CStr(tblName))
        tbl.DataBodyRange.EntireRow.Hidden = False
    Next tblName

    ' --------------------------------------------
    ' フィルター参照セルをリセット
    ' --------------------------------------------
    ws.Range("B3").Value = ""
    ws.Range("C3").Value = ""
    ws.Range("D3").Value = ""
    ws.Range("E3").Value = "全項目"

    ' --------------------------------------------
    ' スクロール位置を先頭に移動
    ' --------------------------------------------
    Application.Goto ws.Range("A1"), True

    ' --------------------------------------------
    ' 終了処理
    ' --------------------------------------------
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "エラーが発生しました" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "詳細: " & Err.Description, vbCritical
End Sub
