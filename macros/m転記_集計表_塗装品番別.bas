Attribute VB_Name = "m転記_集計表_塗装品番別"
Sub 転記_集計表_塗装品番別()
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
    
    ' 品番接頭辞の配列（「合計」を追加）
    Dim prefixList() As Variant
    prefixList = Array("アルヴェルF", "アルヴェルR", "ノアヴォクF", "ノアヴォクR", "補給品", "合計")
    
    ' 転記元列名末尾の配列
    Dim suffixList() As Variant
    suffixList = Array("日実績", "日出来高ｻｲｸﾙ", "累計実績", "平均実績", _
                      "累計出来高ｻｲｸﾙ", "日不良実績", "累計不良数", _
                      "累計不良率", "平均不良数")
    
    ' 転記先行番号の配列（suffixListに対応）
    Dim targetRows() As Variant
    targetRows = Array(20, 21, 22, 23, 24, 26, 27, 28, 29)
    
    ' 品番に対応する転記先列の配列（P列=16を追加）
    Dim targetColumns() As Variant
    targetColumns = Array(6, 8, 10, 12, 14, 16)  ' F, H, J, L, N, P列
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' 高速化設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 進捗表示開始
    Application.StatusBar = "塗装データの転記処理を開始します..."
    
    ' 集計表シート取得
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("集計表")
    If wsTarget Is Nothing Then
        MsgBox "「集計表」シートが見つかりません。" & vbCrLf & _
               "まさか、シート名間違えてないよな？", vbCritical, "シートエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' 集計表のA1セルから日付取得
    If Not IsDate(wsTarget.Range("A1").Value) Then
        MsgBox "集計表のセルA1に有効な日付が入力されていません。" & vbCrLf & _
               "日付も入力できないのか？", vbCritical, "日付エラー"
        GoTo CleanupAndExit
    End If
    targetDate = wsTarget.Range("A1").Value
    
    ' 塗装品番別シート取得
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("塗装品番別")
    If wsSource Is Nothing Then
        MsgBox "「塗装品番別」シートが見つかりません。" & vbCrLf & _
               "シート作るの忘れた？", vbCritical, "シートエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' ソーステーブル取得
    On Error Resume Next
    Set sourceTable = wsSource.ListObjects("_塗装品番別b")
    If sourceTable Is Nothing Then
        MsgBox "「塗装品番別」シートに「_塗装品番別b」テーブルが見つかりません。" & vbCrLf & _
               "テーブル名、ちゃんと確認した？", vbCritical, "テーブルエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' データ範囲取得
    If sourceTable.DataBodyRange Is Nothing Then
        MsgBox "「_塗装品番別b」テーブルにデータがありません。" & vbCrLf & _
               "空のテーブルから何を転記するつもりだ？", vbCritical, "データエラー"
        GoTo CleanupAndExit
    End If
    Set sourceData = sourceTable.DataBodyRange
    
    ' 日付列のインデックス取得
    Dim dateColIndex As Long
    On Error Resume Next
    dateColIndex = sourceTable.ListColumns("日付").Index
    If Err.Number <> 0 Then
        MsgBox "「_塗装品番別b」テーブルに「日付」列が見つかりません。" & vbCrLf & _
               "日付列もないのに日付で検索とか無理だろ", vbCritical, "列エラー"
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
        MsgBox "日付 " & Format(targetDate, "yyyy/mm/dd") & " のデータが見つかりません。" & vbCrLf & _
               "その日のデータ、本当に入力した？", vbCritical, "データエラー"
        GoTo CleanupAndExit
    End If
    
    ' 各品番と末尾の組み合わせで転記処理
    totalCombinations = (UBound(prefixList) + 1) * (UBound(suffixList) + 1)
    processedCombinations = 0
    
    For i = 0 To UBound(prefixList)
        Application.StatusBar = "塗装データ転記中... (" & prefixList(i) & ")"
        
        For k = 0 To UBound(suffixList)
            processedCombinations = processedCombinations + 1
            
            ' 列名を構築（品番接頭辞 + 末尾文字列）
            Dim columnName As String
            columnName = prefixList(i) & suffixList(k)
            
            ' 転記実行
            On Error Resume Next
            Dim colIndex As Long
            colIndex = sourceTable.ListColumns(columnName).Index
            
            If Err.Number = 0 Then
                ' ソース値を一旦変数に格納
                Dim sourceValue As Variant
                sourceValue = sourceData.Cells(sourceRow, colIndex).Value
                
                ' 空白チェックと転記
                If IsEmpty(sourceValue) Or sourceValue = "" Or IsNull(sourceValue) Then
                    wsTarget.Cells(targetRows(k), targetColumns(i)).Value = 0
                Else
                    wsTarget.Cells(targetRows(k), targetColumns(i)).Value = sourceValue
                End If
            Else
                ' 列が見つからない場合は警告（デバッグ用）
                Debug.Print "警告: 列「" & columnName & "」が見つかりません。"
                Err.Clear
            End If
            On Error GoTo ErrorHandler
            
            ' 進捗更新（10件ごと）
            If processedCombinations Mod 10 = 0 Then
                Application.StatusBar = "塗装データ転記中... (" & _
                    processedCombinations & "/" & totalCombinations & ")"
            End If
        Next k
    Next i
    
    ' 正常終了（エラー時以外はメッセージ非表示）
    GoTo CleanupAndExit
    
ErrorHandler:
    MsgBox "転記処理中に予期せぬエラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & vbCrLf & _
           "ちゃんとデバッグしてから実行しろよな", vbCritical, "転記エラー"
    
CleanupAndExit:
    ' 後処理
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False  ' ステータスバーをクリア
End Sub