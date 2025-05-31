Attribute VB_Name = "m転記_集計表_加工作業者別"
Sub 転記_集計表_加工作業者別()
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
    
    ' 転記元列名末尾の配列
    Dim suffixList() As Variant
    suffixList = Array("実績", "日出来高ｻｲｸﾙ", "日時間当り出来高", "累計", _
                      "日平均実績", "平均出来高ｻｲｸﾙ", "平均時間当り数")
    
    ' 転記先行番号の配列（suffixListに対応）
    Dim targetRows() As Variant
    targetRows = Array(59, 60, 61, 62, 63, 64, 65)
    
    ' 作業者名を取得する列の配列（58行目）
    Dim workerColumns() As Variant
    workerColumns = Array(4, 6, 8, 10, 12, 14, 16)  ' D, F, H, J, L, N, P列
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' 高速化設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 進捗表示開始
    Application.StatusBar = "加工作業者別データの転記処理を開始します..."
    
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
    
    ' 加工作業者別シート取得
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("加工作業者別")
    If wsSource Is Nothing Then
        MsgBox "「加工作業者別」シートが見つかりません。", vbCritical, "シートエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' ソーステーブル取得
    On Error Resume Next
    Set sourceTable = wsSource.ListObjects("_加工作業者別b")
    If sourceTable Is Nothing Then
        MsgBox "「加工作業者別」シートに「_加工作業者別b」テーブルが見つかりません。", vbCritical, "テーブルエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' データ範囲取得
    If sourceTable.DataBodyRange Is Nothing Then
        MsgBox "「_加工作業者別b」テーブルにデータがありません。", vbCritical, "データエラー"
        GoTo CleanupAndExit
    End If
    Set sourceData = sourceTable.DataBodyRange
    
    ' 日付列のインデックス取得
    Dim dateColIndex As Long
    On Error Resume Next
    dateColIndex = sourceTable.ListColumns("日付").Index
    If Err.Number <> 0 Then
        MsgBox "「_加工作業者別b」テーブルに「日付」列が見つかりません。", vbCritical, "列エラー"
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
    
    ' 各列の作業者名を取得して転記処理
    totalCombinations = 0
    processedCombinations = 0
    
    ' まず総処理数をカウント（進捗表示用）
    For i = 0 To UBound(workerColumns)
        If wsTarget.Cells(58, workerColumns(i)).Value <> "" Then
            totalCombinations = totalCombinations + (UBound(suffixList) + 1)
        End If
    Next i
    
    ' 各列の処理
    For i = 0 To UBound(workerColumns)
        ' 58行目から作業者名を取得
        Dim workerName As String
        workerName = CStr(wsTarget.Cells(58, workerColumns(i)).Value)
        
        ' 空白セルはスキップ
        If workerName <> "" Then
            Application.StatusBar = "加工作業者別データ転記中... (" & workerName & ")"
            
            ' 各末尾との組み合わせで転記
            For k = 0 To UBound(suffixList)
            processedCombinations = processedCombinations + 1
            
            ' 列名を構築（作業者名 + 末尾文字列）
            Dim columnName As String
            columnName = workerName & suffixList(k)
            
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
                    wsTarget.Cells(targetRows(k), workerColumns(i)).Value = 0
                Else
                    wsTarget.Cells(targetRows(k), workerColumns(i)).Value = sourceValue
                End If
            Else
                ' 列が見つからない場合は警告（デバッグ用）
                Debug.Print "警告: 列「" & columnName & "」が見つかりません。"
                Err.Clear
            End If
            On Error GoTo ErrorHandler
            
                ' 進捗更新
                If processedCombinations Mod 5 = 0 Then
                    Application.StatusBar = "加工作業者別データ転記中... (" & _
                        processedCombinations & "/" & totalCombinations & ")"
                End If
            Next k
        End If
    Next i
    
    ' 正常終了（エラー時以外はメッセージ非表示）
    GoTo CleanupAndExit
    
ErrorHandler:
    MsgBox "転記処理中にエラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "転記エラー"
    
CleanupAndExit:
    ' 後処理
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False  ' ステータスバーをクリア
End Sub