Attribute VB_Name = "mフィルター解除"
Option Explicit

' ========================================
' マクロ名: フィルター解除
' 処理概要: 行非表示フィルターを全解除（全行を表示）
' 対象範囲:
'   - 行7〜190（完成品・小部品）
'   - 行194〜265（コア小部品）
'   - 行269〜304（スリッター小部品）
'   - 行308〜343（ACF小部品）
' ========================================

Sub フィルター解除()
    ' --------------------------------------------
    ' 変数宣言
    ' --------------------------------------------
    Dim ws As Worksheet
    Dim j As Long

    ' 解除対象範囲定義（開始行, 終了行）
    Dim ranges(1 To 4, 1 To 2) As Long
    ranges(1, 1) = 7: ranges(1, 2) = 190      ' 完成品・小部品
    ranges(2, 1) = 194: ranges(2, 2) = 265    ' コア小部品
    ranges(3, 1) = 269: ranges(3, 2) = 304    ' スリッター小部品
    ranges(4, 1) = 308: ranges(4, 2) = 343    ' ACF小部品

    ' --------------------------------------------
    ' 画面更新抑制（高速化）
    ' --------------------------------------------
    Application.ScreenUpdating = False

    On Error GoTo ErrorHandler

    Set ws = ActiveSheet

    ' --------------------------------------------
    ' 全範囲を表示（フィルター解除）
    ' --------------------------------------------
    For j = 1 To 4
        ws.Rows(ranges(j, 1) & ":" & ranges(j, 2)).Hidden = False
    Next j

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
