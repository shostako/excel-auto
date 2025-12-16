Attribute VB_Name = "m側板フィルター"
Option Explicit

' ========================================
' マクロ名: 側板フィルター
' 処理概要: C3セルの値でオートフィルターを適用
' 参照セル: C3（フィルター条件）
' フィルター対象: _完成品テーブルの「側板」列のみ
' 条件分岐:
'   - 「全品番」→ フィルター解除
'   - それ以外 → 指定値 + 「稼働日」「合計」でフィルター
' 複合フィルター: 他のフィルターは維持（オートフィルター方式）
' ========================================

Sub 側板フィルター()
    ' --------------------------------------------
    ' 変数宣言
    ' --------------------------------------------
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim filterValue As String
    Dim colIndex As Long
    Dim filterArray(0 To 2) As String

    ' --------------------------------------------
    ' 画面更新・イベント抑制（高速化）
    ' --------------------------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    Set ws = ActiveSheet

    ' --------------------------------------------
    ' フィルター条件の取得
    ' --------------------------------------------
    filterValue = ws.Range("C3").Value

    ' フィルター条件配列（「稼働日」「合計」を追加）
    filterArray(0) = filterValue
    filterArray(1) = "稼働日"
    filterArray(2) = "合計"

    ' --------------------------------------------
    ' _完成品テーブル：側板列にオートフィルター
    ' --------------------------------------------
    Set tbl = FindTableByPattern(ws, "_完成品")
    colIndex = tbl.ListColumns("側板").Index

    If filterValue = "全品番" Or filterValue = "" Then
        ' フィルター解除
        If tbl.AutoFilter.FilterMode Then
            tbl.Range.AutoFilter Field:=colIndex
        End If
    Else
        ' フィルター適用（「稼働日」「合計」を含む）
        tbl.Range.AutoFilter Field:=colIndex, _
            Criteria1:=filterArray, _
            Operator:=xlFilterValues
    End If

    ' --------------------------------------------
    ' 垂直スクロールのみ先頭に移動（水平位置は維持）
    ' --------------------------------------------
    ActiveWindow.ScrollRow = 1

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
