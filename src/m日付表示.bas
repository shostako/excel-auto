Attribute VB_Name = "m日付表示"
Option Explicit

' ========================================
' マクロ名: 日付表示
' 処理概要: A3の日付を画面中央に表示するようスクロール
' 参照セル: A3（日付）
' 日付行: 6行目（G列以降に日付が並ぶ）
' ウィンドウ枠固定: G7
' 動作: 見つけた日付列から5列左をスクロール開始位置に
' ========================================

Sub 日付表示()
    ' --------------------------------------------
    ' 変数宣言
    ' --------------------------------------------
    Dim ws As Worksheet
    Dim targetDate As Date
    Dim searchRange As Range
    Dim cell As Range
    Dim foundCol As Long
    Dim scrollCol As Long
    Const OFFSET_COLS As Long = 0  ' オフセットなし（日付列が左端）
    Const DATE_ROW As Long = 6     ' 日付が並ぶ行
    Const START_COL As Long = 7    ' G列（日付開始列）

    ' --------------------------------------------
    ' 画面更新抑制
    ' --------------------------------------------
    Application.ScreenUpdating = False

    On Error GoTo ErrorHandler

    Set ws = ActiveSheet

    ' --------------------------------------------
    ' A3から日付を取得
    ' --------------------------------------------
    If Not IsDate(ws.Range("A3").Value) Then
        ' 日付でなければ先頭に戻る
        Application.Goto ws.Range("G7"), True
        GoTo CleanUp
    End If
    targetDate = ws.Range("A3").Value

    ' --------------------------------------------
    ' 6行目のG列以降で日付を検索
    ' --------------------------------------------
    foundCol = 0
    Set searchRange = ws.Range(ws.Cells(DATE_ROW, START_COL), ws.Cells(DATE_ROW, ws.Cells(DATE_ROW, ws.Columns.Count).End(xlToLeft).Column))

    For Each cell In searchRange
        If IsDate(cell.Value) Then
            If CDate(cell.Value) = targetDate Then
                foundCol = cell.Column
                Exit For
            End If
        End If
    Next cell

    ' --------------------------------------------
    ' スクロール位置の決定
    ' --------------------------------------------
    If foundCol > 0 Then
        ' 見つかった：5列左をスクロール開始位置に（最小G列）
        scrollCol = foundCol - OFFSET_COLS
        If scrollCol < START_COL Then scrollCol = START_COL
        Application.Goto ws.Cells(7, scrollCol), True
    Else
        ' 見つからない：先頭（G7）に戻る
        Application.Goto ws.Range("G7"), True
    End If

CleanUp:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "詳細: " & Err.Description, vbCritical
End Sub
