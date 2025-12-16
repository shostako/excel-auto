Attribute VB_Name = "m項目追加"
Option Explicit

' ========================================
' マクロ名: 項目追加
' 処理概要: 指定テーブルで指定した項目行の下に新しい項目行を追加
' 引数:
'   tableName - 対象テーブル名（パターン検索）
'   targetItem - 検索する項目値（この行の下に追加）
'   newItem - 追加する行の項目値
' 処理詳細:
'   - 製品名/側板/詳細/小部品：存在する列のみ上の行の値・書式をコピー
'   - 項目：newItemを設定
'   - 合計：空白
'   - その他（日付列等）：文字色自動、数値フォーマット、左隣セル参照関数
' 対応テーブル:
'   - _完成品系：製品名、側板、小部品列あり
'   - _core/_slitter/_acf系：詳細、小部品列あり
' 冪等性: 直下にnewItemが既に存在する場合はスキップ
' 注意: フィルターがかかっている場合は自動解除される
' ========================================

Sub 項目追加(tableName As String, targetItem As String, newItem As String)
    ' --------------------------------------------
    ' 変数宣言
    ' --------------------------------------------
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim itemCol As ListColumn
    Dim dataArr As Variant
    Dim i As Long
    Dim insertCount As Long

    ' 列インデックス（0は列なしを意味）
    Dim colProduct As Long      ' 製品名
    Dim colSide As Long         ' 側板
    Dim colDetail As Long       ' 詳細
    Dim colPart As Long         ' 小部品
    Dim colItem As Long         ' 項目
    Dim colTotal As Long        ' 合計
    Dim lastCol As Long         ' 最終列

    ' 行操作用
    Dim newRow As ListRow
    Dim srcRow As Long          ' コピー元（targetItem）の行番号
    Dim tblStartRow As Long     ' テーブルデータ開始行
    Dim j As Long
    Dim formulaCol As Long

    ' targetItem行インデックスを格納する配列
    Dim targetRows() As Long
    Dim targetCount As Long

    ' --------------------------------------------
    ' 画面更新・計算・イベント抑制（高速化）
    ' --------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    Set ws = ActiveSheet

    ' --------------------------------------------
    ' テーブル取得（完全一致）
    ' --------------------------------------------
    Dim tmpTbl As ListObject
    For Each tmpTbl In ws.ListObjects
        If tmpTbl.Name = tableName Then
            Set tbl = tmpTbl
            Exit For
        End If
    Next tmpTbl
    If tbl Is Nothing Then
        MsgBox "「" & tableName & "」テーブルが見つかりません", vbExclamation
        GoTo Cleanup
    End If

    ' --------------------------------------------
    ' フィルター解除（行追加のため）
    ' --------------------------------------------
    If Not tbl.AutoFilter Is Nothing Then
        If tbl.AutoFilter.FilterMode Then
            tbl.AutoFilter.ShowAllData
        End If
    End If

    ' --------------------------------------------
    ' 列インデックス取得（存在チェック付き、0は列なし）
    ' --------------------------------------------
    If HasColumn(tbl, "製品名") Then colProduct = tbl.ListColumns("製品名").Index Else colProduct = 0
    If HasColumn(tbl, "側板") Then colSide = tbl.ListColumns("側板").Index Else colSide = 0
    If HasColumn(tbl, "詳細") Then colDetail = tbl.ListColumns("詳細").Index Else colDetail = 0
    If HasColumn(tbl, "小部品") Then colPart = tbl.ListColumns("小部品").Index Else colPart = 0
    colItem = tbl.ListColumns("項目").Index
    colTotal = tbl.ListColumns("合計").Index
    lastCol = tbl.ListColumns.Count

    tblStartRow = tbl.DataBodyRange.Row

    ' --------------------------------------------
    ' targetItem行のインデックスを収集
    ' --------------------------------------------
    Set itemCol = tbl.ListColumns("項目")
    dataArr = itemCol.DataBodyRange.Value

    targetCount = 0
    ReDim targetRows(1 To UBound(dataArr, 1))

    For i = 1 To UBound(dataArr, 1)
        If dataArr(i, 1) = targetItem Then
            targetCount = targetCount + 1
            targetRows(targetCount) = i
        End If
    Next i

    If targetCount = 0 Then
        MsgBox "「" & targetItem & "」行が見つかりません", vbExclamation
        GoTo Cleanup
    End If

    ' --------------------------------------------
    ' 下から上へ処理（行挿入によるインデックスずれ防止）
    ' --------------------------------------------
    insertCount = 0

    For i = targetCount To 1 Step -1
        Dim rowIdx As Long
        rowIdx = targetRows(i)

        ' 直下の行がnewItemかチェック（スキップ判定）
        If rowIdx < tbl.ListRows.Count Then
            If tbl.DataBodyRange.Cells(rowIdx + 1, colItem).Value = newItem Then
                ' 既に存在するのでスキップ
                GoTo NextIteration
            End If
        End If

        ' 行挿入（targetItemの直後に挿入）
        Set newRow = tbl.ListRows.Add(rowIdx + 1)
        srcRow = tblStartRow + rowIdx - 1  ' targetItemのシート行番号

        ' --------------------------------------------
        ' 製品名/側板/詳細/小部品：存在する列のみ上の行から値・書式をコピー
        ' --------------------------------------------
        If colProduct > 0 Then
            ws.Cells(srcRow, tbl.Range.Column + colProduct - 1).Copy
            ws.Cells(srcRow + 1, tbl.Range.Column + colProduct - 1).PasteSpecial xlPasteAll
        End If

        If colSide > 0 Then
            ws.Cells(srcRow, tbl.Range.Column + colSide - 1).Copy
            ws.Cells(srcRow + 1, tbl.Range.Column + colSide - 1).PasteSpecial xlPasteAll
        End If

        If colDetail > 0 Then
            ws.Cells(srcRow, tbl.Range.Column + colDetail - 1).Copy
            ws.Cells(srcRow + 1, tbl.Range.Column + colDetail - 1).PasteSpecial xlPasteAll
        End If

        If colPart > 0 Then
            ws.Cells(srcRow, tbl.Range.Column + colPart - 1).Copy
            ws.Cells(srcRow + 1, tbl.Range.Column + colPart - 1).PasteSpecial xlPasteAll
        End If

        ' --------------------------------------------
        ' 項目：newItemを設定（書式デフォルト）
        ' --------------------------------------------
        With newRow.Range.Cells(1, colItem)
            .Value = newItem
            .Interior.pattern = xlNone
            .Font.ColorIndex = xlAutomatic
        End With

        ' --------------------------------------------
        ' 合計：空白（書式デフォルト）
        ' --------------------------------------------
        With newRow.Range.Cells(1, colTotal)
            .Value = ""
            .Interior.pattern = xlNone
            .Font.ColorIndex = xlAutomatic
        End With

        ' --------------------------------------------
        ' 日付列等：書式デフォルト、数値フォーマット、左隣セル参照
        ' --------------------------------------------
        For j = colTotal + 1 To lastCol
            With newRow.Range.Cells(1, j)
                .Interior.pattern = xlNone
                .Font.ColorIndex = xlAutomatic
                .NumberFormat = "0"
                ' 左隣セル参照の関数を設定
                formulaCol = tbl.Range.Column + j - 2  ' 左隣の列番号
                .Formula = "=" & ws.Cells(srcRow + 1, formulaCol).Address(False, False)
            End With
        Next j

        insertCount = insertCount + 1

NextIteration:
    Next i

    Application.CutCopyMode = False

    ' --------------------------------------------
    ' 完了（正常終了時はメッセージなし、ただし追加件数0なら通知）
    ' --------------------------------------------
    If insertCount = 0 Then
        MsgBox "追加対象の行がありませんでした（全て「" & newItem & "」が存在済み）", vbInformation
    End If

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "エラーが発生しました" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "詳細: " & Err.Description, vbCritical
End Sub

' ========================================
' 列存在チェック関数
' ========================================
Private Function HasColumn(tbl As ListObject, colName As String) As Boolean
    Dim col As ListColumn
    On Error Resume Next
    Set col = tbl.ListColumns(colName)
    HasColumn = Not col Is Nothing
    On Error GoTo 0
End Function
