Attribute VB_Name = "m転記_統合_塗装P"
Option Explicit

' ========================================
' マクロ名: 転記_統合_塗装P
' 処理概要: 塗装G/T/Hの出力テーブルを塗装Pシートに統合コピーし、印刷用レイアウトを作成
'
' 【処理の特徴】
' 1. 従来マクロ流用：塗装G/T/Hの出力テーブルをそのままコピー
' 2. 統合表示：G→T→Hの順に縦並びで配置
' 3. 期間別ページ：各期間で改ページを挿入
' 4. 印刷対応：全体を印刷範囲に設定
'
' 【テーブル構成】
' 期間テーブル : シート「塗装P」、テーブル「_集計期間塗装P」
' コピー元テーブル :
'   - シート「塗装G」、テーブル「_流出G_塗装_{期間名}」
'   - シート「塗装T」、テーブル「_手直しT_塗装_{期間名}」
'   - シート「塗装H」、テーブル「_廃棄H_塗装_{期間名}」
'
' 【処理フロー】
' 1. 塗装Pシートの既存データ・改ページをクリア
' 2. 期間テーブルから期間名リストを取得
' 3. 各期間でG→T→Hテーブルをコピー（間隔2行）
' 4. 期間ごとに改ページを挿入
' 5. 全体を印刷範囲に設定
' ========================================

Sub 転記_統合_塗装P()
    ' ============================================
    ' 最適化設定の保存と適用
    ' ============================================
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    Dim origEnableEvents As Boolean
    Dim origDisplayAlerts As Boolean

    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation
    origEnableEvents = Application.EnableEvents
    origDisplayAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    On Error GoTo ErrorHandler
    Application.StatusBar = "塗装P統合転記処理を開始します..."

    ' ============================================
    ' シートとテーブルの参照取得
    ' ============================================
    Dim wsTarget As Worksheet
    Dim wsG As Worksheet, wsT As Worksheet, wsH As Worksheet

    Set wsTarget = ThisWorkbook.Worksheets("塗装P")
    Set wsG = ThisWorkbook.Worksheets("塗装G")
    Set wsT = ThisWorkbook.Worksheets("塗装T")
    Set wsH = ThisWorkbook.Worksheets("塗装H")

    Dim tblPeriod As ListObject
    On Error Resume Next
    Set tblPeriod = wsTarget.ListObjects("_集計期間塗装P")
    On Error GoTo ErrorHandler

    If tblPeriod Is Nothing Then
        MsgBox "シート「塗装P」にテーブル「_集計期間塗装P」が見つかりません。", vbCritical
        GoTo Cleanup
    End If

    ' ============================================
    ' 期間情報の読み込み
    ' ============================================
    Dim periodCount As Long
    periodCount = 0
    If Not tblPeriod.DataBodyRange Is Nothing Then
        periodCount = tblPeriod.DataBodyRange.Rows.Count
    End If

    If periodCount = 0 Then
        MsgBox "期間テーブルにデータがありません。", vbExclamation
        GoTo Cleanup
    End If

    ' 期間名を配列に格納（空白期間はスキップ）
    Dim periodNames() As String
    Dim validPeriodCount As Long
    validPeriodCount = 0

    Dim p As Long
    Dim tempPeriodName As String
    Dim tempStartDate As Variant

    For p = 1 To periodCount
        tempStartDate = tblPeriod.DataBodyRange.Cells(p, 2).Value
        ' 開始日が空または0ならスキップ
        If Not IsEmpty(tempStartDate) And tempStartDate <> 0 Then
            validPeriodCount = validPeriodCount + 1
        End If
    Next p

    If validPeriodCount = 0 Then
        MsgBox "有効な期間データがありません。", vbExclamation
        GoTo Cleanup
    End If

    ReDim periodNames(1 To validPeriodCount)
    Dim idx As Long
    idx = 0
    For p = 1 To periodCount
        tempStartDate = tblPeriod.DataBodyRange.Cells(p, 2).Value
        If Not IsEmpty(tempStartDate) And tempStartDate <> 0 Then
            idx = idx + 1
            periodNames(idx) = CStr(tblPeriod.DataBodyRange.Cells(p, 1).Value)
        End If
    Next p

    ' ============================================
    ' 既存データのクリア（設定テーブル以外）
    ' ============================================
    Application.StatusBar = "既存データをクリア中..."

    ' 設定テーブルの最終行を取得
    Dim tblItems As ListObject
    On Error Resume Next
    Set tblItems = wsTarget.ListObjects("_流出項目塗装P")
    On Error GoTo ErrorHandler

    Dim settingsEndRow As Long
    settingsEndRow = 1

    If Not tblItems Is Nothing Then
        settingsEndRow = Application.WorksheetFunction.Max(settingsEndRow, _
            tblItems.Range.Row + tblItems.Range.Rows.Count)
    End If
    If Not tblPeriod Is Nothing Then
        settingsEndRow = Application.WorksheetFunction.Max(settingsEndRow, _
            tblPeriod.Range.Row + tblPeriod.Range.Rows.Count)
    End If

    ' 設定テーブルより下のデータをクリア
    Dim clearStartRow As Long
    clearStartRow = settingsEndRow + 2

    Dim lastUsedRow As Long
    lastUsedRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row

    If lastUsedRow >= clearStartRow Then
        wsTarget.Rows(clearStartRow & ":" & lastUsedRow).Clear
    End If

    ' 改ページをクリア
    wsTarget.ResetAllPageBreaks

    ' 印刷範囲をクリア
    wsTarget.PageSetup.PrintArea = ""

    ' ============================================
    ' テーブルのコピー処理
    ' ============================================
    Dim currentRow As Long
    currentRow = clearStartRow

    Dim printStartRow As Long
    Dim printEndRow As Long
    Dim printLastCol As Long
    printStartRow = currentRow
    printEndRow = 0
    printLastCol = 1

    Dim periodIdx As Long
    Dim tblG As ListObject, tblT As ListObject, tblH As ListObject
    Dim tblNameG As String, tblNameT As String, tblNameH As String
    Dim copyRange As Range
    Dim pasteCell As Range

    For periodIdx = 1 To validPeriodCount
        Application.StatusBar = "期間 " & periodIdx & "/" & validPeriodCount & " を処理中..."

        ' テーブル名を構築
        tblNameG = "_流出G_塗装_" & Replace(periodNames(periodIdx), " ", "_")
        tblNameT = "_手直しT_塗装_" & Replace(periodNames(periodIdx), " ", "_")
        tblNameH = "_廃棄H_塗装_" & Replace(periodNames(periodIdx), " ", "_")

        ' 塗装Gテーブルをコピー（タイトル行含む）
        On Error Resume Next
        Set tblG = wsG.ListObjects(tblNameG)
        On Error GoTo ErrorHandler

        If Not tblG Is Nothing Then
            ' タイトル行（テーブルの1行上）を含む範囲を取得
            Dim titleRowG As Long
            titleRowG = tblG.Range.Row - 1
            Set copyRange = wsG.Range(wsG.Cells(titleRowG, 1), _
                wsG.Cells(tblG.Range.Row + tblG.Range.Rows.Count - 1, _
                          tblG.Range.Column + tblG.Range.Columns.Count - 1))
            Set pasteCell = wsTarget.Cells(currentRow, 1)

            copyRange.Copy
            pasteCell.PasteSpecial xlPasteAll
            Application.CutCopyMode = False

            ' 列幅を更新
            If copyRange.Columns.Count > printLastCol Then
                printLastCol = copyRange.Columns.Count
            End If

            currentRow = currentRow + copyRange.Rows.Count + 2 ' 2行空ける
        End If

        ' 塗装Tテーブルをコピー（タイトル行含む）
        On Error Resume Next
        Set tblT = wsT.ListObjects(tblNameT)
        On Error GoTo ErrorHandler

        If Not tblT Is Nothing Then
            ' タイトル行（テーブルの1行上）を含む範囲を取得
            Dim titleRowT As Long
            titleRowT = tblT.Range.Row - 1
            Set copyRange = wsT.Range(wsT.Cells(titleRowT, 1), _
                wsT.Cells(tblT.Range.Row + tblT.Range.Rows.Count - 1, _
                          tblT.Range.Column + tblT.Range.Columns.Count - 1))
            Set pasteCell = wsTarget.Cells(currentRow, 1)

            copyRange.Copy
            pasteCell.PasteSpecial xlPasteAll
            Application.CutCopyMode = False

            If copyRange.Columns.Count > printLastCol Then
                printLastCol = copyRange.Columns.Count
            End If

            currentRow = currentRow + copyRange.Rows.Count + 2
        End If

        ' 塗装Hテーブルをコピー（タイトル行含む）
        On Error Resume Next
        Set tblH = wsH.ListObjects(tblNameH)
        On Error GoTo ErrorHandler

        If Not tblH Is Nothing Then
            ' タイトル行（テーブルの1行上）を含む範囲を取得
            Dim titleRowH As Long
            titleRowH = tblH.Range.Row - 1
            Set copyRange = wsH.Range(wsH.Cells(titleRowH, 1), _
                wsH.Cells(tblH.Range.Row + tblH.Range.Rows.Count - 1, _
                          tblH.Range.Column + tblH.Range.Columns.Count - 1))
            Set pasteCell = wsTarget.Cells(currentRow, 1)

            copyRange.Copy
            pasteCell.PasteSpecial xlPasteAll
            Application.CutCopyMode = False

            If copyRange.Columns.Count > printLastCol Then
                printLastCol = copyRange.Columns.Count
            End If

            currentRow = currentRow + copyRange.Rows.Count + 2
        End If

        ' この期間の最終行を記録
        printEndRow = currentRow - 2 ' 空白行の前

        ' 改ページを挿入（最後の期間以外）
        If periodIdx < validPeriodCount Then
            wsTarget.HPageBreaks.Add Before:=wsTarget.Rows(currentRow)
        End If
    Next periodIdx

    ' ============================================
    ' 印刷範囲と印刷設定
    ' ============================================
    If printEndRow > printStartRow Then
        With wsTarget.PageSetup
            ' 印刷範囲設定
            .PrintArea = wsTarget.Range( _
                wsTarget.Cells(printStartRow, 1), _
                wsTarget.Cells(printEndRow, printLastCol)).Address

            ' 「シートを1ページに収める」を解除し、改ページを有効にする
            .Zoom = False
            .FitToPagesWide = 1    ' 幅は1ページに収める
            .FitToPagesTall = False ' 高さは自動（改ページに従う）
            .CenterVertically = False ' 上詰め（中央配置しない）
            .CenterHorizontally = False ' 左詰め
        End With
        Debug.Print "印刷範囲設定: " & wsTarget.PageSetup.PrintArea
    End If

    Application.StatusBar = "塗装P統合転記処理が完了しました"
    Debug.Print "塗装P統合転記完了: " & validPeriodCount & "期間処理"

Cleanup:
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.DisplayAlerts = origDisplayAlerts
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
    Resume Cleanup
End Sub
