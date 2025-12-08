Attribute VB_Name = "m小部品フィルター"
Option Explicit

' ========================================
' マクロ名: 小部品フィルター
' 処理概要: D3セルの値と一致するD列の行のみ表示（不一致行は非表示）
' 参照セル: D3（フィルター条件）
' フィルター対象: D7:D190（小部品名列）
' データ範囲: B7:GP190
' 最適化: 配列一括読み込み + 計算/イベント抑制
' ========================================

Sub 小部品フィルター()
    ' --------------------------------------------
    ' 変数宣言
    ' --------------------------------------------
    Dim ws As Worksheet
    Dim filterValue As String
    Dim i As Long
    Dim dataArr As Variant

    Const START_ROW As Long = 7      ' データ開始行
    Const END_ROW As Long = 190      ' データ終了行
    Const FILTER_COL As Long = 4     ' D列（フィルター対象列）
    Const CRITERIA_CELL As String = "D3"  ' フィルター条件セル

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
    filterValue = ws.Range(CRITERIA_CELL).Value

    ' --------------------------------------------
    ' 全行を表示（リセット）
    ' --------------------------------------------
    ws.Rows(START_ROW & ":" & END_ROW).Hidden = False

    ' --------------------------------------------
    ' データを配列に一括読み込み
    ' --------------------------------------------
    dataArr = ws.Range(ws.Cells(START_ROW, FILTER_COL), ws.Cells(END_ROW, FILTER_COL)).Value

    ' --------------------------------------------
    ' フィルタリング：不一致行を非表示
    ' --------------------------------------------
    For i = 1 To UBound(dataArr, 1)
        If dataArr(i, 1) <> filterValue Then
            ws.Rows(START_ROW + i - 1).Hidden = True
        End If
    Next i

    ' --------------------------------------------
    ' スクロール位置を先頭に移動
    ' --------------------------------------------
    Application.Goto ws.Range("A7"), True

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
