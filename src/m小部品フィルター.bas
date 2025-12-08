Attribute VB_Name = "m小部品フィルター"
Option Explicit

' ========================================
' マクロ名: 小部品フィルター
' 処理概要: D3セルの値と一致するD列の行のみ表示（不一致行は非表示）
' 参照セル: D3（フィルター条件）
' フィルター対象:
'   - D7:D190（小部品名列）
'   - D194:D265（コア小部品名列）
'   - D269:D304（スリッター小部品名列）
'   - D308:D343（ACF小部品名列）
' データ範囲: B7:HK190, B194:HK249, B253:HK280, B284:HK311
' 最適化: 配列一括読み込み + 計算/イベント抑制
' ========================================

Sub 小部品フィルター()
    ' --------------------------------------------
    ' 変数宣言
    ' --------------------------------------------
    Dim ws As Worksheet
    Dim filterValue As String
    Dim i As Long, j As Long
    Dim dataArr As Variant

    ' フィルター範囲定義（開始行, 終了行, フィルター列）
    Dim ranges(1 To 4, 1 To 3) As Long
    ranges(1, 1) = 7: ranges(1, 2) = 190: ranges(1, 3) = 4    ' D列（小部品）
    ranges(2, 1) = 194: ranges(2, 2) = 265: ranges(2, 3) = 4  ' D列（コア小部品）
    ranges(3, 1) = 269: ranges(3, 2) = 304: ranges(3, 3) = 4  ' D列（スリッター小部品）
    ranges(4, 1) = 308: ranges(4, 2) = 343: ranges(4, 3) = 4  ' D列（ACF小部品）

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
    filterValue = ws.Range("D3").Value

    ' --------------------------------------------
    ' 全範囲を表示（リセット）
    ' --------------------------------------------
    For j = 1 To 4
        ws.Rows(ranges(j, 1) & ":" & ranges(j, 2)).Hidden = False
    Next j

    ' --------------------------------------------
    ' 各範囲でフィルタリング
    ' --------------------------------------------
    For j = 1 To 4
        ' データを配列に一括読み込み
        dataArr = ws.Range(ws.Cells(ranges(j, 1), ranges(j, 3)), _
                          ws.Cells(ranges(j, 2), ranges(j, 3))).Value

        ' 不一致行を非表示
        For i = 1 To UBound(dataArr, 1)
            If dataArr(i, 1) <> filterValue Then
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
