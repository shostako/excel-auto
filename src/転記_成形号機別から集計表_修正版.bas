Attribute VB_Name = "m転記_成形号機別_改"
Sub 転記_成形号機別から集計表()
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
    
    ' 号機前置詞の配列
    Dim prefixList() As Variant
    prefixList = Array("1号機", "2号機", "12計", "3号機", "4号機", "34計")
    
    ' 転記列項目名後置の配列
    Dim suffixList() As Variant
    suffixList = Array("生産数", "予定数", "標準出来高", "設計時間", _
                      "設計予定", "標準直接", "標準出来高計", _
                      "実不良数", "実不良率", "設計不良数", _
                      "設計不良率", "標準不良数")
    
    ' 転記先行番号の配列（suffixListに対応）
    Dim targetRows() As Variant
    targetRows = Array(4, 5, 6, 7, 8, 9, 10, 12, 13, 14, 15, 16)
    
    ' 号機に対応する転記列の配列
    Dim targetColumns() As Variant
    targetColumns = Array(6, 8, 10, 12, 14, 16)  ' F, H, J, L, N, P列
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' 進捗表示開始
    Application.StatusBar = "成形データの転記処理を開始します..."
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 集計表シート取得
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("集計表")
    If wsTarget Is Nothing Then
        MsgBox "「集計表」シートが見つかりません。", vbCritical, "シートエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' 集計表のA1セルから日付取得
    If Not IsDate(wsTarget.Range("A1").Value) Then
        MsgBox "集計表のセルA1に有効な日付が入力されていません。", vbCritical, "日付エラー"
        GoTo CleanupAndExit
    End If
    targetDate = wsTarget.Range("A1").Value
    
    ' 成形号機別シート取得
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("成形号機別")
    If wsSource Is Nothing Then
        MsgBox "「成形号機別」シートが見つかりません。", vbCritical, "シートエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' ソーステーブル取得
    On Error Resume Next
    Set sourceTable = wsSource.ListObjects("_成形号機別b")
    If sourceTable Is Nothing Then
        MsgBox "「成形号機別」シートに「_成形号機別b」テーブルが見つかりません。", vbCritical, "テーブルエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' データ範囲取得
    If sourceTable.DataBodyRange Is Nothing Then
        MsgBox "「_成形号機別b」テーブルにデータがありません。", vbCritical, "データエラー"
        GoTo CleanupAndExit
    End If
    Set sourceData = sourceTable.DataBodyRange
    
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
    
    ' 各号機と項目の組み合わせで転記処理
    totalCombinations = (UBound(prefixList) + 1) * (UBound(suffixList) + 1)
    processedCombinations = 0
    
    For i = 0 To UBound(prefixList)
        Application.StatusBar = "成形データ転記中... (" & prefixList(i) & ")"
        
        For k = 0 To UBound(suffixList)
            processedCombinations = processedCombinations + 1
            
            ' 列名を構築（号機前置詞 + 項目後置詞）
            Dim columnName As String
            columnName = prefixList(i) & suffixList(k)
            
            ' 転記処理
            On Error Resume Next
            Dim colIndex As Long
            colIndex = sourceTable.ListColumns(columnName).Index
            
            If Err.Number = 0 Then
                ' 値を一度変数に格納
                Dim cellValue As Variant
                cellValue = sourceData.Cells(sourceRow, colIndex).Value
                
                ' 空白、エラー値のチェック
                If IsEmpty(cellValue) Or IsError(cellValue) Then
                    wsTarget.Cells(targetRows(k), targetColumns(i)).Value = 0
                ElseIf VarType(cellValue) = vbString And Trim(cellValue) = "" Then
                    ' 文字列で空白の場合のみ0を設定
                    wsTarget.Cells(targetRows(k), targetColumns(i)).Value = 0
                Else
                    wsTarget.Cells(targetRows(k), targetColumns(i)).Value = cellValue
                End If
            Else
                ' 列が見つからない場合は0を設定
                wsTarget.Cells(targetRows(k), targetColumns(i)).Value = 0
                Debug.Print "警告: 列「" & columnName & "」が見つかりません。"
                Err.Clear
            End If
            On Error GoTo ErrorHandler
            
            ' 進捗更新（10回ごと）
            If processedCombinations Mod 10 = 0 Then
                Application.StatusBar = "成形データ転記中... (" & _
                    processedCombinations & "/" & totalCombinations & ")"
            End If
        Next k
    Next i
    
    ' 正常終了（エラー以外はメッセージを表示）
    GoTo CleanupAndExit
    
ErrorHandler:
    MsgBox "転記処理中にエラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "転記エラー"
    
CleanupAndExit:
    ' 後処理
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False  ' ステータスバーをクリア
End Sub