Attribute VB_Name = "m完成品フィルター"
Option Explicit

' ========================================
' マクロ名: 完成品フィルター
' 処理概要: D2セルの値で完成品をフィルター、末尾4字除去した値で小部品をフィルター
' 参照セル: D2（フィルター条件）
' フィルター対象:
'   - B7:B190（完成品名列）→ D2そのまま
'   - D194:D265（コア小部品名列）→ D2末尾4字除去
'   - D269:D304（スリッター小部品名列）→ D2末尾4字除去
'   - D308:D343（ACF小部品名列）→ D2末尾4字除去
' データ範囲: B7:HK190, B194:HK249, B253:HK280, B284:HK311
' 最適化: 配列一括読み込み + 計算/イベント抑制
' ========================================

Sub 完成品フィルター()
    ' --------------------------------------------
    ' 変数宣言
    ' --------------------------------------------
    Dim ws As Worksheet
    Dim filterValue As String
    Dim filterValueTrimmed As String
    Dim rawValue As String
    Dim i As Long, j As Long
    Dim dataArr As Variant
    Dim currentFilter As String

    ' 小部品フィルター範囲定義（開始行, 終了行, フィルター列）
    Dim ranges(1 To 3, 1 To 3) As Long
    ranges(1, 1) = 194: ranges(1, 2) = 265: ranges(1, 3) = 4  ' D列（コア小部品）
    ranges(2, 1) = 269: ranges(2, 2) = 304: ranges(2, 3) = 4  ' D列（スリッター小部品）
    ranges(3, 1) = 308: ranges(3, 2) = 343: ranges(3, 3) = 4  ' D列（ACF小部品）

    Const START_ROW As Long = 7      ' 完成品データ開始行
    Const END_ROW As Long = 190      ' 完成品データ終了行
    Const FILTER_COL As Long = 2     ' B列（完成品フィルター対象列）

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
    rawValue = ws.Range("D2").Value
    filterValue = rawValue                              ' 完成品用（そのまま）
    filterValueTrimmed = Left(rawValue, Len(rawValue) - 4)  ' 小部品用（末尾4字除去）

    ' --------------------------------------------
    ' 全範囲を表示（リセット）
    ' --------------------------------------------
    ws.Rows(START_ROW & ":" & END_ROW).Hidden = False
    For j = 1 To 3
        ws.Rows(ranges(j, 1) & ":" & ranges(j, 2)).Hidden = False
    Next j

    ' --------------------------------------------
    ' 完成品範囲のフィルタリング（D2そのまま）
    ' --------------------------------------------
    dataArr = ws.Range(ws.Cells(START_ROW, FILTER_COL), ws.Cells(END_ROW, FILTER_COL)).Value
    For i = 1 To UBound(dataArr, 1)
        If dataArr(i, 1) <> filterValue Then
            ws.Rows(START_ROW + i - 1).Hidden = True
        End If
    Next i

    ' --------------------------------------------
    ' 小部品範囲のフィルタリング（D2末尾4字除去）
    ' --------------------------------------------
    For j = 1 To 3
        dataArr = ws.Range(ws.Cells(ranges(j, 1), ranges(j, 3)), _
                          ws.Cells(ranges(j, 2), ranges(j, 3))).Value
        For i = 1 To UBound(dataArr, 1)
            If dataArr(i, 1) <> filterValueTrimmed Then
                ws.Rows(ranges(j, 1) + i - 1).Hidden = True
            End If
        Next i
    Next j

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
