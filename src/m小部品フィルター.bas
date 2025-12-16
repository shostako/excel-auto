Attribute VB_Name = "m小部品フィルター"
Option Explicit

' ========================================
' マクロ名: 小部品フィルター
' 処理概要: D3セルの値でオートフィルターを適用
' 参照セル: D3（フィルター条件）
' フィルター対象: 全テーブルの「小部品」列
' 条件分岐:
'   - 「全品番」→ フィルター解除
'   - それ以外 → 指定値 + 「稼働日」「合計」でフィルター
' 複合フィルター: 他のフィルターは維持（オートフィルター方式）
' ========================================

Sub 小部品フィルター()
    ' --------------------------------------------
    ' 変数宣言
    ' --------------------------------------------
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim filterValue As String
    Dim colIndex As Long
    Dim filterArray(0 To 2) As String

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
    ' フィルター条件の取得
    ' --------------------------------------------
    filterValue = ws.Range("D3").Value

    ' フィルター条件配列（「稼働日」「合計」を追加）
    filterArray(0) = filterValue
    filterArray(1) = "稼働日"
    filterArray(2) = "合計"

    ' --------------------------------------------
    ' 全テーブル：小部品列にオートフィルター
    ' --------------------------------------------
    Dim tblName As Variant
    For Each tblName In tables
        Set tbl = FindTableByPattern(ws, CStr(tblName))
        colIndex = tbl.ListColumns("小部品").Index

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
    Next tblName

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
