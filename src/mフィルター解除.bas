Attribute VB_Name = "mフィルター解除"
Option Explicit

' ========================================
' マクロ名: フィルター解除
' 処理概要: 行非表示フィルターを全解除（全行を表示）
' 対象範囲: 行7〜190
' データ範囲: B7:GP190
' ========================================

Sub フィルター解除()
    ' --------------------------------------------
    ' 変数宣言
    ' --------------------------------------------
    Dim ws As Worksheet

    Const START_ROW As Long = 7      ' データ開始行
    Const END_ROW As Long = 190      ' データ終了行

    ' --------------------------------------------
    ' 画面更新抑制（高速化）
    ' --------------------------------------------
    Application.ScreenUpdating = False

    On Error GoTo ErrorHandler

    Set ws = ActiveSheet

    ' --------------------------------------------
    ' 全行を表示（フィルター解除）
    ' --------------------------------------------
    ws.Rows(START_ROW & ":" & END_ROW).Hidden = False

    ' --------------------------------------------
    ' スクロール位置を先頭に移動
    ' --------------------------------------------
    Application.Goto ws.Range("A7"), True

    ' --------------------------------------------
    ' 終了処理
    ' --------------------------------------------
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "詳細: " & Err.Description, vbCritical
End Sub
