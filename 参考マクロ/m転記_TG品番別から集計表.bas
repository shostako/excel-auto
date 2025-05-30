Attribute VB_Name = "m転記_TG品番別から集計表"
Sub 転記_TG品番別から集計表()
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
    
    ' 品番接頭辞の配列
    Dim prefixList() As Variant
    prefixList = Array("RH", "LH", "合計")
    
    ' 転記元列名末尾の配列
    Dim suffixList() As Variant
    suffixList = Array("日実績", "日出来高ｻｲｸﾙ", "累計実績", "平均実績", _
                      "累計出来高ｻｲｸﾙ", "日不良実績", "累計不良数", _
                      "累計不良率", "平均不良数")
    
    ' 転記先行番号の配列（suffixListに対応）
    Dim targetRows() As Variant
    targetRows = Array(33, 34, 35, 36, 37, 39, 40, 41, 42)
    
    ' 品番に対応する転記先列の配列
    Dim targetColumns() As Variant
    targetColumns = Array(12, 14, 16)  ' L, N, P列
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' 進捗表示開始
    Application.StatusBar = "TGデータの転記処理を開始します..."
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 集計表シート取得
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("集計表")
    If wsTarget Is Nothing Then
        MsgBox "「集計表」シートが見つかりません。" & vbCrLf & _
               "シート名ぐらい覚えとけよ", vbCritical
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' 集計表のA1セルから日付取得
    If Not IsDate(wsTarget.Range("A1").Value) Then
        MsgBox "集計表のセルA1に有効な日付が入力されていません。" & vbCrLf & _
               "日付の入力もできないとか、大丈夫か？", vbCritical
        GoTo CleanupAndExit
    End If
    targetDate = wsTarget.Range("A1").Value
    
    ' TG品番別シート取得
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("TG品番別")
    If wsSource Is Nothing Then
        MsgBox "「TG品番別」シートが見つかりません。" & vbCrLf & _
               "シート作ってから実行しろよ", vbCritical
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' ソーステーブル取得
    On Error Resume Next
    Set sourceTable = wsSource.ListObjects("_TG品番別b")
    If sourceTable Is Nothing Then
        MsgBox "「TG品番別」シートに「_TG品番別b」テーブルが見つかりません。" & vbCrLf & _
               "テーブル名ぐらい統一しろよな", vbCritical
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' データ範囲取得
    If sourceTable.DataBodyRange Is Nothing Then
        MsgBox "「_TG品番別b」テーブルにデータがありません。" & vbCrLf & _
               "空っぽのテーブルから何を転記する気だ？", vbCritical
        GoTo CleanupAndExit
    End If
    Set sourceData = sourceTable.DataBodyRange
    
    ' 日付列のインデックス取得
    Dim dateColIndex As Long
    On Error Resume Next
    dateColIndex = sourceTable.ListColumns("日付").Index
    If Err.Number <> 0 Then
        MsgBox "「_TG品番別b」テーブルに「日付」列が見つかりません。" & vbCrLf & _
               "日付列もないのに何を基準に転記するんだよ", vbCritical
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
               "その日のデータ、本当に存在するのか？", vbCritical
        GoTo CleanupAndExit
    End If
    
    ' 各品番と末尾の組み合わせで転記処理
    totalCombinations = (UBound(prefixList) + 1) * (UBound(suffixList) + 1)
    processedCombinations = 0
    
    For i = 0 To UBound(prefixList)
        Application.StatusBar = "TGデータ転記中... (" & prefixList(i) & ")"
        
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
                ' 転記先のセルに値を設定
                wsTarget.Cells(targetRows(k), targetColumns(i)).Value = _
                    sourceData.Cells(sourceRow, colIndex).Value
            Else
                ' 列が見つからない場合は警告（デバッグ用）
                Debug.Print "警告: 列「" & columnName & "」が見つかりません。"
                Err.Clear
            End If
            On Error GoTo ErrorHandler
            
            ' 進捗更新（10件ごと）
            If processedCombinations Mod 10 = 0 Then
                Application.StatusBar = "TGデータ転記中... (" & _
                    processedCombinations & "/" & totalCombinations & ")"
            End If
        Next k
    Next i
    
    ' 正常終了メッセージ（コメントアウト済み - エラー時以外は非表示）
    ' MsgBox "TGデータの転記が完了しました。", vbInformation
    GoTo CleanupAndExit
    
ErrorHandler:
    MsgBox "転記処理中に予期せぬエラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & vbCrLf & _
           "デバッグぐらいしてから実行しろよ", vbCritical, "転記エラー"
    
CleanupAndExit:
    ' 後処理
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False  ' ステータスバーをクリア
End Sub

