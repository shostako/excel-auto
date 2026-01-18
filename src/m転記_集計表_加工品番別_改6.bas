Attribute VB_Name = "m転記_集計表_加工品番別_改6"
Option Explicit

' ========================================
' マクロ名: 転記_集計表_加工品番別
' 処理概要: 加工品番別テーブルデータを集計表シートの加工別欄に転記
' ソーステーブル: シート「加工品番別」テーブル「_加工品番別b」
' ターゲットシート: シート「集計表」
' 転記対象: 6カテゴリ×9項目のマトリックス形式データ + 段取関連項目
' カテゴリ分類: アルヴェルF/R、ノアヴォクF/R、補給品、合計の6分類
' 転記項目: 実績・出来高・不良関連の9種類の集計値 + 段取関連6項目
' 段取項目: 日平均段取時間、日段取時間、日段取回数、累計段取時間、累計段取回数、平均段取時間
' 修正内容: 段取時間関連（日段取時間、累計段取時間、平均段取時間）を60倍してH→m換算
' 変更点: 合計平均段取時間をT51に追加転記
' 改6変更点: アルヴェルF・合計の不良実績/累計不良数をG/Q列に移動、累計不良率・平均不良数の転記を廃止
' ========================================
Sub 転記_集計表_加工品番別()
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

    ' ステータスバー初期化
    Application.StatusBar = "加工品番別 集計表転記を開始..."

    ' エラーハンドリング設定
    On Error GoTo ErrorHandler

    ' =================================
    ' 第1段階：変数宣言と基本設定
    ' =================================

    ' 変数宣言
    Dim wsTarget As Worksheet
    Dim wsSource As Worksheet
    Dim targetDate As Date
    Dim sourceTable As ListObject
    Dim sourceData As Range
    Dim i As Long, j As Long, k As Long
    Dim sourceRow As Long
    Dim totalProcesses As Long
    Dim processedCount As Long

    ' =================================
    ' 第2段階：転記マトリックス設定
    ' =================================

    ' 転記列名の配列（末尾部分）
    Dim columnSuffixes() As Variant
    columnSuffixes = Array("日実績", "日出来高ｻｲｸﾙ", "累計実績", "平均実績", _
                          "累計出来高ｻｲｸﾙ", "日不良実績", "累計不良数", _
                          "累計不良率", "平均不良数")

    ' 転記先行番号の配列
    Dim targetRows() As Variant
    targetRows = Array(46, 47, 48, 49, 50, 52, 53, 54, 55)

    ' 各カテゴリの情報（接頭辞、転記先列）
    Dim categoryInfo() As Variant
    categoryInfo = Array( _
        Array("アルヴェルF", 6), _
        Array("アルヴェルR", 8), _
        Array("ノアヴォクF", 10), _
        Array("ノアヴォクR", 12), _
        Array("補給品", 14), _
        Array("合計", 16) _
    )

    ' =================================
    ' 第3段階：集計表シートの取得と検証
    ' =================================

    Application.StatusBar = "集計表シートを取得中..."

    ' 集計表シート取得
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("集計表")
    If wsTarget Is Nothing Then
        Application.StatusBar = "集計表シートが見つかりません"
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler

    ' 集計表のA1セルから日付取得
    If Not IsDate(wsTarget.Range("A1").Value) Then
        Application.StatusBar = "集計表A1に有効な日付がありません"
        GoTo Cleanup
    End If
    targetDate = wsTarget.Range("A1").Value

    ' =================================
    ' 第4段階：ソースデータの取得と検証
    ' =================================

    Application.StatusBar = "ソースデータを取得中..."

    ' ソースシート取得（加工品番別シート）
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("加工品番別")
    If wsSource Is Nothing Then
        Application.StatusBar = "加工品番別シートが見つかりません"
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler

    ' ソーステーブル取得
    On Error Resume Next
    Set sourceTable = wsSource.ListObjects("_加工品番別b")
    If sourceTable Is Nothing Then
        Application.StatusBar = "テーブル_加工品番別bが見つかりません"
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler

    ' データ範囲取得
    If sourceTable.DataBodyRange Is Nothing Then
        Application.StatusBar = "ソーステーブルにデータがありません"
        GoTo Cleanup
    End If
    Set sourceData = sourceTable.DataBodyRange

    ' =================================
    ' 第5段階：対象日付行の特定
    ' =================================

    ' 日付列のインデックス取得
    Dim dateColIndex As Long
    On Error Resume Next
    dateColIndex = sourceTable.ListColumns("日付").Index
    If Err.Number <> 0 Then
        Application.StatusBar = "日付列が見つかりません"
        GoTo Cleanup
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
        Application.StatusBar = "対象日付 " & Format(targetDate, "yyyy/mm/dd") & " のデータなし"
        GoTo Cleanup
    End If

    ' =================================
    ' 第6段階：マトリックス形式転記処理
    ' =================================

    ' 処理総数の計算（進捗表示用）
    totalProcesses = (UBound(categoryInfo) + 1) * (UBound(columnSuffixes) + 1)
    processedCount = 0

    Application.StatusBar = "マトリックス転記を実行中..."

    ' 各カテゴリ（アルヴェルF/R、ノアヴォクF/R、補給品、合計）について処理
    Dim categoryIndex As Long
    For categoryIndex = 0 To UBound(categoryInfo)
        Dim prefix As String
        Dim targetCol As Long
        prefix = categoryInfo(categoryIndex)(0)
        targetCol = categoryInfo(categoryIndex)(1)

        ' 各項目（9種類）について転記
        For k = 0 To UBound(columnSuffixes)
            processedCount = processedCount + 1

            Dim fullColumnName As String
            Dim colIndex As Long

            ' 完全な列名を構築
            fullColumnName = prefix & columnSuffixes(k)

            ' ============================================
            ' 改6: アルヴェルF・合計の累計不良率・平均不良数はスキップ
            ' ============================================
            If (prefix = "アルヴェルF" Or prefix = "合計") And _
               (columnSuffixes(k) = "累計不良率" Or columnSuffixes(k) = "平均不良数") Then
                ' 転記しない
                GoTo NextIteration
            End If

            ' 列インデックス取得
            On Error Resume Next
            colIndex = sourceTable.ListColumns(fullColumnName).Index

            If Err.Number = 0 Then
                ' ソース値を一旦変数に格納
                Dim sourceValue As Variant
                sourceValue = sourceData.Cells(sourceRow, colIndex).Value

                ' 空白チェック
                If IsEmpty(sourceValue) Or sourceValue = "" Or IsNull(sourceValue) Then
                    sourceValue = 0
                End If

                ' ============================================
                ' 改6: アルヴェルF・合計の不良実績/累計不良数は隣の列に転記
                ' ============================================
                Dim actualTargetCol As Long
                If (prefix = "アルヴェルF" Or prefix = "合計") And _
                   (columnSuffixes(k) = "日不良実績" Or columnSuffixes(k) = "累計不良数") Then
                    ' アルヴェルF: F(6)→G(7)、合計: P(16)→Q(17)
                    actualTargetCol = targetCol + 1
                Else
                    actualTargetCol = targetCol
                End If

                wsTarget.Cells(targetRows(k), actualTargetCol).Value = sourceValue

                ' 不良率列の場合は書式設定（パーセント表示）
                If InStr(fullColumnName, "不良率") > 0 Then
                    wsTarget.Cells(targetRows(k), actualTargetCol).NumberFormatLocal = "0.0%"
                End If
            Else
                ' 列が見つからない場合は警告（デバッグ用）
                Debug.Print "警告: 列「" & fullColumnName & "」が見つかりません。"
                Err.Clear
            End If
            On Error GoTo ErrorHandler

NextIteration:
            ' 進捗更新（5件ごと）
            If processedCount Mod 5 = 0 Or processedCount = totalProcesses Then
                Application.StatusBar = "加工集計表転記中... " & Format(processedCount / totalProcesses, "0%") & _
                                      " (" & processedCount & "/" & totalProcesses & "件)"
            End If
        Next k
    Next categoryIndex

    ' =================================
    ' 第7段階：段取関連データの転記処理
    ' =================================

    Application.StatusBar = "段取データ転記中..."

    ' 段取関連転記の定義（合計系のみ、T列=20列目）
    Dim dandoriTransfers() As Variant
    dandoriTransfers = Array( _
        Array("合計日平均段取時間", 46, 20), _
        Array("合計日段取時間", 47, 20), _
        Array("合計日段取回数", 48, 20), _
        Array("合計累計段取時間", 49, 20), _
        Array("合計累計段取回数", 50, 20), _
        Array("合計平均段取時間", 51, 20) _
    )

    ' 段取関連データの転記実行
    Dim transferItem As Variant
    For i = 0 To UBound(dandoriTransfers)
        transferItem = dandoriTransfers(i)
        Dim columnName As String: columnName = transferItem(0)  ' 列名
        Dim targetRow As Long: targetRow = transferItem(1)  ' 転記先行
        targetCol = transferItem(2)  ' 転記先列

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

    ' 処理完了表示
    Application.StatusBar = "加工集計表転記完了: " & processedCount & "件のデータを転記"
    Application.Wait Now + TimeValue("0:00:01")

    GoTo Cleanup

ErrorHandler:
    ' エラー情報の取得
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear

    ' エラーメッセージの表示
    MsgBox "加工品番別 集計表転記でエラー発生" & vbCrLf & _
           "エラー番号: " & errNum & vbCrLf & _
           "詳細: " & errDesc, vbCritical, "転記エラー"

Cleanup:
    ' 設定を元に戻す
    Application.StatusBar = False
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
End Sub
