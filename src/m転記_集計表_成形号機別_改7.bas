Attribute VB_Name = "m転記_集計表_成形号機別_改7"
Option Explicit

' ========================================
' マクロ名: 転記_集計表_成形号機別
' 処理概要: 号機別データを集計表に転記し、ジャンル別ライン分析を可能にする
' ソーステーブル: シート「成形号機別」テーブル「_成形号機別b」
' 基準日付: シート「集計表」セルA1の日付
' 転記先: シート「集計表」4-16行目、F-P列（6列構成）+ T列（段取関連）
' ライン構成:
'   - 1-2-3号機ライン: サンショウ製品（123合計で管理）
'   - 4号機: 単独ライン
'   - 5号機: スペーサー（新規追加）
' 転記指標: 12種類（日ｼｮｯﾄ、日実績、日出来高ｻｲｸﾙ、累計ｼｮｯﾄ、累計実績、平均実績、平均出来高ｻｲｸﾙ、日不良実績、日不良率、累計不良実績、累計不良率、平均不良数）
' 段取指標: 6種類（日平均段取時間、日段取時間、日段取回数、累計段取時間、累計段取回数、平均段取時間）
' 修正内容: 段取時間関連（日段取時間、累計段取時間、平均段取時間）を60倍してH→m換算
' 変更点: 段取関連転記を1号機T4-T9、2号機T11-T16に配置変更、平均段取時間追加
' 改6変更点: 列構成を12合計→123合計に変更、5号機を新規追加
' 改7変更点: 1号機・123合計の不良実績をG/M列に移動、不良率・平均不良数の転記を廃止
' ========================================
Sub 転記_集計表_成形号機別()
    ' 最適化設定の保存
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    Dim origEnableEvents As Boolean
    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation
    origEnableEvents = Application.EnableEvents

    ' 最適化設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' 変数宣言
    Dim wsTarget As Worksheet
    Dim wsSource As Worksheet
    Dim targetDate As Date
    Dim sourceTable As ListObject
    Dim sourceData As Range
    Dim i As Long, j As Long, k As Long
    Dim sourceRow As Long
    Dim totalCombinations As Long
    Dim processedCombinations As Long

    ' ジャンル別ライン構成の定義（6列構成）
    ' F列=1号機、H列=2号機、J列=3号機、L列=123合計、N列=4号機、P列=5号機
    Dim prefixList() As Variant
    prefixList = Array("1号機", "2号機", "3号機", "123合計", "4号機", "5号機")

    ' 転記対象指標の定義（12種類）
    Dim suffixList() As Variant
    suffixList = Array("日ｼｮｯﾄ", "日実績", "日出来高ｻｲｸﾙ", "累計ｼｮｯﾄ", _
                      "累計実績", "平均実績", "平均出来高ｻｲｸﾙ", _
                      "日不良実績", "日不良率", "累計不良実績", _
                      "累計不良率", "平均不良数")

    ' 転記先行番号の配列（各指標の配置行）
    Dim targetRows() As Variant
    targetRows = Array(4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15, 16)

    ' 転記先列番号の配列（ジャンル別ライン構成に対応）
    Dim targetColumns() As Variant
    targetColumns = Array(6, 8, 10, 12, 14, 16)  ' F, H, J, L, N, P列

    ' エラーハンドリング設定
    On Error GoTo ErrorHandler

    ' 進捗表示開始
    Application.StatusBar = "成形データの転記処理を開始します..."

    ' =================================
    ' 第1段階：シートとテーブルの取得・検証
    ' =================================

    ' 転記先シート（集計表）の取得
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("集計表")
    If wsTarget Is Nothing Then
        MsgBox "「集計表」シートが見つかりません。", vbCritical, "シートエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler

    ' 基準日付の取得（集計表A1セルから）
    If Not IsDate(wsTarget.Range("A1").Value) Then
        MsgBox "集計表のセルA1に有効な日付が入力されていません。", vbCritical, "日付エラー"
        GoTo CleanupAndExit
    End If
    targetDate = wsTarget.Range("A1").Value

    ' ソースシート（成形号機別）の取得
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("成形号機別")
    If wsSource Is Nothing Then
        MsgBox "「成形号機別」シートが見つかりません。", vbCritical, "シートエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler

    ' ソーステーブル（_成形号機別b）の取得
    On Error Resume Next
    Set sourceTable = wsSource.ListObjects("_成形号機別b")
    If sourceTable Is Nothing Then
        MsgBox "「成形号機別」シートに「_成形号機別b」テーブルが見つかりません。", vbCritical, "テーブルエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler

    ' データ範囲の確認
    If sourceTable.DataBodyRange Is Nothing Then
        MsgBox "「_成形号機別b」テーブルにデータがありません。", vbCritical, "データエラー"
        GoTo CleanupAndExit
    End If
    Set sourceData = sourceTable.DataBodyRange

    ' =================================
    ' 第2段階：基準日付に一致する行の検索
    ' =================================

    ' 日付列のインデックス取得
    Dim dateColIndex As Long
    On Error Resume Next
    dateColIndex = sourceTable.ListColumns("日付").Index
    If Err.Number <> 0 Then
        MsgBox "「_成形号機別b」テーブルに「日付」列が見つかりません。", vbCritical, "列エラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler

    ' 該当日付の行を検索
    sourceRow = 0
    For j = 1 To sourceData.Rows.Count
        If sourceData.Cells(j, dateColIndex).Value = targetDate Then
            sourceRow = j
            Exit For
        End If
    Next j

    If sourceRow = 0 Then
        MsgBox "日付 " & Format(targetDate, "yyyy/mm/dd") & " のデータが見つかりません。", vbCritical, "データエラー"
        GoTo CleanupAndExit
    End If

    ' =================================
    ' 第3段階：ジャンル別ライン×指標の転記処理
    ' =================================

    ' 処理総数の計算（進捗表示用）
    ' 6ライン × 12指標 = 72件の転記処理
    totalCombinations = (UBound(prefixList) + 1) * (UBound(suffixList) + 1)
    processedCombinations = 0

    ' 各ライン（号機・合計）の処理
    For i = 0 To UBound(prefixList)
        Application.StatusBar = "成形データ転記中... (" & prefixList(i) & ")"

        ' 各指標の転記処理
        For k = 0 To UBound(suffixList)
            processedCombinations = processedCombinations + 1

            ' 列名を構築（例：1号機日実績、123合計累計ｼｮｯﾄ）
            Dim columnName As String
            columnName = prefixList(i) & suffixList(k)

            ' ============================================
            ' 改7: 1号機・123合計の不良率・平均不良数はスキップ
            ' ============================================
            If (prefixList(i) = "1号機" Or prefixList(i) = "123合計") And _
               (suffixList(k) = "日不良率" Or suffixList(k) = "累計不良率" Or suffixList(k) = "平均不良数") Then
                ' 転記しない
                GoTo NextIteration
            End If

            ' 該当列からデータを転記
            On Error Resume Next
            Dim colIndex As Long
            colIndex = sourceTable.ListColumns(columnName).Index

            If Err.Number = 0 Then
                ' ソースデータの値を取得
                Dim sourceValue As Variant
                sourceValue = sourceData.Cells(sourceRow, colIndex).Value

                ' 空白・NULL値の処理（明示的に0に変換）
                If IsEmpty(sourceValue) Or sourceValue = "" Or IsNull(sourceValue) Then
                    sourceValue = 0
                End If

                ' ============================================
                ' 改7: 1号機・123合計の不良実績は隣の列に転記
                ' ============================================
                Dim actualTargetCol As Long
                If (prefixList(i) = "1号機" Or prefixList(i) = "123合計") And _
                   (suffixList(k) = "日不良実績" Or suffixList(k) = "累計不良実績") Then
                    ' 1号機: F(6)→G(7)、123合計: L(12)→M(13)
                    actualTargetCol = targetColumns(i) + 1
                Else
                    actualTargetCol = targetColumns(i)
                End If

                wsTarget.Cells(targetRows(k), actualTargetCol).Value = sourceValue
            Else
                ' 列が見つからない場合の警告（デバッグ用）
                Debug.Print "警告: 列「" & columnName & "」が見つかりません。"
                Err.Clear
            End If
            On Error GoTo ErrorHandler

NextIteration:
            ' 進捗更新（10件ごと）
            If processedCombinations Mod 10 = 0 Then
                Application.StatusBar = "成形データ転記中... (" & _
                    processedCombinations & "/" & totalCombinations & ")"
            End If
        Next k
    Next i

    ' =================================
    ' 第4段階：段取関連データの転記処理
    ' =================================

    Application.StatusBar = "段取データ転記中..."

    ' 段取関連転記の定義（1号機T4-T9、2号機T11-T16）
    ' T列=20列目
    Dim dandoriTransfers() As Variant
    dandoriTransfers = Array( _
        Array("1号機日平均段取時間", 4, 20), _
        Array("1号機日段取時間", 5, 20), _
        Array("1号機日段取回数", 6, 20), _
        Array("1号機累計段取時間", 7, 20), _
        Array("1号機累計段取回数", 8, 20), _
        Array("1号機平均段取時間", 9, 20), _
        Array("2号機日平均段取時間", 11, 20), _
        Array("2号機日段取時間", 12, 20), _
        Array("2号機日段取回数", 13, 20), _
        Array("2号機累計段取時間", 14, 20), _
        Array("2号機累計段取回数", 15, 20), _
        Array("2号機平均段取時間", 16, 20) _
    )

    ' 段取関連データの転記実行
    Dim transferItem As Variant
    For i = 0 To UBound(dandoriTransfers)
        transferItem = dandoriTransfers(i)
        columnName = transferItem(0)  ' 列名
        Dim targetRow As Long: targetRow = transferItem(1)  ' 転記先行
        Dim targetCol As Long: targetCol = transferItem(2)  ' 転記先列

        ' 該当列からデータを転記
        On Error Resume Next
        colIndex = sourceTable.ListColumns(columnName).Index

        If Err.Number = 0 Then
            ' ソースデータの値を取得
            sourceValue = sourceData.Cells(sourceRow, colIndex).Value

            ' 時間関連の処理（H→m換算）
            If InStr(columnName, "時間") > 0 Then
                ' 60倍して分単位に変換
                wsTarget.Cells(targetRow, targetCol).Value = sourceValue * 60
                ' 書式設定：平均段取時間のみ小数点1桁、その他は整数
                If InStr(columnName, "平均段取時間") > 0 Then
                    wsTarget.Cells(targetRow, targetCol).NumberFormatLocal = "_-* 0.0"" 分"""
                Else
                    wsTarget.Cells(targetRow, targetCol).NumberFormatLocal = "_-* 0"" 分"""
                End If
            Else
                ' 回数関連：通常の転記
                wsTarget.Cells(targetRow, targetCol).Value = sourceValue
                ' 書式設定：0を含めて位置統一（数字と単位の間にスペース）
                wsTarget.Cells(targetRow, targetCol).NumberFormatLocal = "_-* 0"" 回"""
            End If
        Else
            ' 列が見つからない場合の警告（デバッグ用）
            Debug.Print "警告: 段取列「" & columnName & "」が見つかりません。"
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Next i

    ' 正常終了処理
    Application.StatusBar = "転記処理完了"
    Application.Wait Now + TimeValue("0:00:01")
    GoTo CleanupAndExit

ErrorHandler:
    ' エラー情報の取得
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear

    ' エラーメッセージの表示
    MsgBox "転記処理中にエラーが発生しました。" & vbCrLf & _
           "エラー番号: " & errNum & vbCrLf & _
           "詳細: " & errDesc, vbCritical, "転記エラー"

CleanupAndExit:
    ' 設定を元に戻す
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.StatusBar = False
End Sub
