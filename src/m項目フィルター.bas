Attribute VB_Name = "m項目フィルター"
Option Explicit

' ========================================
' マクロ名: 項目フィルター
' 処理概要: E3セルの値に基づいて「項目」列をフィルター（複合フィルター対応）
' 参照セル: E3（項目フィルター条件）
' 他のフィルター参照: B3（完成品）, C3（側板）, D3（小部品）
' フィルター対象: 全テーブルの「項目」列
' 条件分岐:
'   - 「計画全体」→「計画」を含む行を表示
'   - 「実績全体」→「実績」を含む行を表示
'   - 「計画実績全体」→「計画」or「実績」を含む行を表示
'   - その他 → 完全一致
' 特殊動作: 一旦リセットして他のフィルターを再適用後、項目フィルターを適用
' 最適化: 配列一括読み込み + 計算/イベント抑制
' ========================================

Sub 項目フィルター()
    ' --------------------------------------------
    ' 変数宣言
    ' --------------------------------------------
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim filterItem As String
    Dim filterProduct As String
    Dim filterSide As String
    Dim filterPart As String
    Dim filterProductTrimmed As String
    Dim filterMode As String
    Dim dataArr As Variant
    Dim i As Long
    Dim startRow As Long
    Dim rowNum As Long
    Dim cellValue As String

    ' 対象テーブル名
    Dim tables As Variant
    tables = Array("_完成品", "_core", "_slitter", "_acf")

    ' 小部品テーブル名（完成品フィルター用）
    Dim subTables As Variant
    subTables = Array("_core", "_slitter", "_acf")

    ' --------------------------------------------
    ' 画面更新・計算・イベント抑制（高速化）
    ' --------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    Set ws = ActiveSheet

    ' --------------------------------------------
    ' 全フィルター条件の取得
    ' --------------------------------------------
    filterProduct = ws.Range("B3").Value      ' 完成品
    filterSide = ws.Range("C3").Value         ' 側板
    filterPart = ws.Range("D3").Value         ' 小部品
    filterItem = ws.Range("E3").Value         ' 項目

    ' 完成品フィルター用（末尾4字除去）
    If Len(filterProduct) > 4 Then
        filterProductTrimmed = Left(filterProduct, Len(filterProduct) - 4)
    Else
        filterProductTrimmed = ""
    End If

    ' 項目フィルターモード判定
    filterMode = "exact"
    Select Case filterItem
        Case "全項目"
            filterMode = "none"  ' フィルター適用しない
        Case "計画全体"
            filterMode = "contains_keikaku"
        Case "実績全体"
            filterMode = "contains_jisseki"
        Case "計画実績全体"
            filterMode = "contains_both"
    End Select

    ' --------------------------------------------
    ' 1. 全テーブルの行を表示（リセット）
    ' --------------------------------------------
    Dim tblName As Variant
    For Each tblName In tables
        Set tbl = FindTableByPattern(ws, CStr(tblName))
        tbl.DataBodyRange.EntireRow.Hidden = False
    Next tblName

    ' --------------------------------------------
    ' 2. 完成品フィルター再適用（B3が空でない場合）
    ' --------------------------------------------
    If Len(filterProduct) > 0 Then
        ' _完成品テーブル：製品名列
        Set tbl = FindTableByPattern(ws, "_完成品")
        startRow = tbl.DataBodyRange.Row
        dataArr = tbl.ListColumns("製品名").DataBodyRange.Value
        For i = 1 To UBound(dataArr, 1)
            If dataArr(i, 1) <> filterProduct Then
                ws.Rows(startRow + i - 1).Hidden = True
            End If
        Next i

        ' 小部品テーブル：小部品列（末尾4字除去）
        If Len(filterProductTrimmed) > 0 Then
            For Each tblName In subTables
                Set tbl = FindTableByPattern(ws, CStr(tblName))
                startRow = tbl.DataBodyRange.Row
                dataArr = tbl.ListColumns("小部品").DataBodyRange.Value
                For i = 1 To UBound(dataArr, 1)
                    If dataArr(i, 1) <> filterProductTrimmed Then
                        ws.Rows(startRow + i - 1).Hidden = True
                    End If
                Next i
            Next tblName
        End If
    End If

    ' --------------------------------------------
    ' 3. 側板フィルター再適用（C3が空でない場合、_完成品のみ）
    ' --------------------------------------------
    If Len(filterSide) > 0 Then
        Set tbl = FindTableByPattern(ws, "_完成品")
        startRow = tbl.DataBodyRange.Row
        dataArr = tbl.ListColumns("側板").DataBodyRange.Value
        For i = 1 To UBound(dataArr, 1)
            rowNum = startRow + i - 1
            If Not ws.Rows(rowNum).Hidden Then
                If dataArr(i, 1) <> filterSide Then
                    ws.Rows(rowNum).Hidden = True
                End If
            End If
        Next i
    End If

    ' --------------------------------------------
    ' 4. 小部品フィルター再適用（D3が空でない場合）
    ' --------------------------------------------
    If Len(filterPart) > 0 Then
        For Each tblName In tables
            Set tbl = FindTableByPattern(ws, CStr(tblName))
            startRow = tbl.DataBodyRange.Row
            dataArr = tbl.ListColumns("小部品").DataBodyRange.Value
            For i = 1 To UBound(dataArr, 1)
                rowNum = startRow + i - 1
                If Not ws.Rows(rowNum).Hidden Then
                    If dataArr(i, 1) <> filterPart Then
                        ws.Rows(rowNum).Hidden = True
                    End If
                End If
            Next i
        Next tblName
    End If

    ' --------------------------------------------
    ' 5. 項目フィルター適用（E3が空でなく、かつ「全項目」でない場合）
    ' --------------------------------------------
    If Len(filterItem) > 0 And filterMode <> "none" Then
        For Each tblName In tables
            Set tbl = FindTableByPattern(ws, CStr(tblName))
            startRow = tbl.DataBodyRange.Row
            dataArr = tbl.ListColumns("項目").DataBodyRange.Value

            For i = 1 To UBound(dataArr, 1)
                rowNum = startRow + i - 1
                If Not ws.Rows(rowNum).Hidden Then
                    cellValue = dataArr(i, 1)
                    Select Case filterMode
                        Case "exact"
                            If cellValue <> filterItem Then
                                ws.Rows(rowNum).Hidden = True
                            End If
                        Case "contains_keikaku"
                            If InStr(cellValue, "計画") = 0 Then
                                ws.Rows(rowNum).Hidden = True
                            End If
                        Case "contains_jisseki"
                            If InStr(cellValue, "実績") = 0 Then
                                ws.Rows(rowNum).Hidden = True
                            End If
                        Case "contains_both"
                            If InStr(cellValue, "計画") = 0 And InStr(cellValue, "実績") = 0 Then
                                ws.Rows(rowNum).Hidden = True
                            End If
                    End Select
                End If
            Next i
        Next tblName
    End If

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
