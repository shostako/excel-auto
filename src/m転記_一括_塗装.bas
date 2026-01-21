Attribute VB_Name = "m転記_一括_塗装"
Option Explicit

' ============================================
' 塗装関連マクロ一括実行
' ============================================
' アクティブシートの期間設定を基準に、塗装関連の全マクロを一括実行する。
'
' 実行対象マクロ：
' - 転記_日報_塗装N
' - 転記_手直し_塗装T
' - 転記_流出_塗装G
' - 転記_日報_塗装NW
' - 転記_廃棄_塗装H
' - 転記_統合_塗装P
'
' 処理フロー：
' 1. アクティブシートの期間テーブル読み込み
' 2. 各塗装シートの期間テーブルをアクティブシートに合わせる
'    - 期間データをコピー（入る分だけ）
'    - 余分な行は空白にクリア（行の挿入・削除なし）
' 3. 各マクロを順次実行
' 4. エラー時は即座に中断
' ============================================

Sub 転記_塗装一括()
    Application.StatusBar = "塗装一括転記を開始します..."

    On Error GoTo ErrorHandler

    ' ============================================
    ' 1. アクティブシートの期間テーブル検出
    ' ============================================
    Dim wsActive As Worksheet
    Set wsActive = ActiveSheet

    Dim tblPeriod As ListObject
    Set tblPeriod = Nothing

    ' ListObjectsから"_集計期間"で始まるテーブルを検索
    Dim tbl As ListObject
    For Each tbl In wsActive.ListObjects
        If Left(tbl.Name, 5) = "_集計期間" Then
            Set tblPeriod = tbl
            Exit For
        End If
    Next tbl

    If tblPeriod Is Nothing Then
        MsgBox "アクティブシートに期間テーブル（_集計期間*）が見つかりません。処理を中止します。", vbExclamation
        GoTo Cleanup
    End If

    ' ============================================
    ' 2. 期間データ読み込み
    ' ============================================
    Dim periodCount As Long
    periodCount = 0
    If Not tblPeriod.DataBodyRange Is Nothing Then
        periodCount = tblPeriod.DataBodyRange.Rows.Count
    End If

    If periodCount = 0 Then
        MsgBox "期間テーブルにデータがありません。処理を中止します。", vbExclamation
        GoTo Cleanup
    End If

    Dim periodInfo() As Variant
    ReDim periodInfo(1 To periodCount, 1 To 3)

    Dim p As Long
    For p = 1 To periodCount
        periodInfo(p, 1) = CStr(tblPeriod.DataBodyRange.Cells(p, 1).Value) ' 期間名
        periodInfo(p, 2) = tblPeriod.DataBodyRange.Cells(p, 2).Value       ' 開始日
        periodInfo(p, 3) = tblPeriod.DataBodyRange.Cells(p, 3).Value       ' 終了日
    Next p

    Application.StatusBar = "各シートの期間テーブルを同期中..."

    ' ============================================
    ' 3. 対象シート×テーブルのリスト定義
    ' ============================================
    Dim targetSheets As Variant
    Dim targetTables As Variant
    targetSheets = Array("塗装N", "塗装T", "塗装G", "塗装NW", "塗装H", "塗装P")
    targetTables = Array("_集計期間日報塗装", "_集計期間塗装T", "_集計期間塗装G", "_集計期間日報塗装W", "_集計期間塗装H", "_集計期間塗装P")

    ' ============================================
    ' 4. 各テーブルをアクティブシートの期間構成に同期
    ' ============================================
    Dim i As Long
    For i = LBound(targetSheets) To UBound(targetSheets)
        Dim ws As Worksheet
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(targetSheets(i))
        On Error GoTo ErrorHandler

        If ws Is Nothing Then
            MsgBox "シート「" & targetSheets(i) & "」が見つかりません。処理を中止します。", vbCritical
            GoTo Cleanup
        End If

        Dim tblTarget As ListObject
        Set tblTarget = Nothing
        On Error Resume Next
        Set tblTarget = ws.ListObjects(targetTables(i))
        On Error GoTo ErrorHandler

        If tblTarget Is Nothing Then
            MsgBox "シート「" & targetSheets(i) & "」にテーブル「" & targetTables(i) & "」が見つかりません。処理を中止します。", vbCritical
            GoTo Cleanup
        End If

        ' 行数確認
        Dim currentCount As Long
        If tblTarget.DataBodyRange Is Nothing Then
            currentCount = 0
        Else
            currentCount = tblTarget.DataBodyRange.Rows.Count
        End If

        ' コピー可能な行数を決定（期間数と既存行数の小さい方）
        Dim copyCount As Long
        copyCount = Application.WorksheetFunction.Min(periodCount, currentCount)

        ' 期間データをコピー（入る分だけ）
        For p = 1 To copyCount
            tblTarget.DataBodyRange.Cells(p, 1).Value = periodInfo(p, 1) ' 期間名
            tblTarget.DataBodyRange.Cells(p, 2).Value = periodInfo(p, 2) ' 開始日
            tblTarget.DataBodyRange.Cells(p, 3).Value = periodInfo(p, 3) ' 終了日
        Next p

        ' 余分な行を空白にクリア
        If currentCount > periodCount Then
            For p = periodCount + 1 To currentCount
                tblTarget.DataBodyRange.Cells(p, 1).Value = "" ' 期間名クリア
                tblTarget.DataBodyRange.Cells(p, 2).Value = "" ' 開始日クリア
                tblTarget.DataBodyRange.Cells(p, 3).Value = "" ' 終了日クリア
            Next p
        End If
    Next i

    ' ============================================
    ' 5. 各マクロを順次実行
    ' ============================================
    Dim macroNames As Variant
    macroNames = Array("転記_日報_塗装N", "転記_手直し_塗装T", "転記_流出_塗装G", "転記_日報_塗装NW", "転記_廃棄_塗装H", "転記_統合_塗装P")

    For i = LBound(macroNames) To UBound(macroNames)
        Application.StatusBar = "実行中: " & macroNames(i) & " (" & (i + 1) & "/" & (UBound(macroNames) + 1) & ")"

        On Error GoTo MacroError
        Application.Run macroNames(i)
        On Error GoTo ErrorHandler
    Next i

    Application.StatusBar = "塗装一括転記が完了しました"
    GoTo Cleanup

MacroError:
    MsgBox "マクロ「" & macroNames(i) & "」の実行中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "エラー内容: " & Err.Description, vbCritical
    GoTo Cleanup

ErrorHandler:
    Application.StatusBar = False
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical

Cleanup:
    Application.StatusBar = False
End Sub
