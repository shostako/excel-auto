Attribute VB_Name = "m転記_モールFR別から集計表"
Option Explicit

' モールFR別から集計表への転記マクロ
' 「_モールFR別b」テーブルから「集計表」シートへデータを転記
Sub 転記_モールFR別から集計表()
    ' 変数宣言
    Dim wsTarget As Worksheet
    Dim wsSource As Worksheet
    Dim targetDate As Date
    Dim sourceTable As ListObject
    Dim sourceData As Range
    Dim i As Long, j As Long, k As Long
    Dim sourceRow As Long
    
    ' 転記列名の配列（末尾部分）
    Dim columnSuffixes() As Variant
    columnSuffixes = Array("日実績", "日出来高ｻｲｸﾙ", "累計実績", "平均実績", _
                          "平均出来高ｻｲｸﾙ", "日不良率", "不良累計", _
                          "累計不良率", "平均不良数")
    
    ' 転記先行番号の配列
    Dim targetRows() As Variant
    targetRows = Array(33, 34, 35, 36, 37, 39, 40, 41, 42)
    
    ' 各カテゴリの情報（接頭辞、転記先列）
    Dim categoryInfo() As Variant
    categoryInfo = Array( _
        Array("モールF", 6), _
        Array("モールR", 8), _
        Array("合計", 10) _
    )
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' 進捗表示開始
    Application.StatusBar = "モールFR別データの転記処理を開始します..."
    
    ' 集計表シート取得
    On Error Resume Next
    Set wsTarget = ThisWorkbook.Worksheets("集計表")
    If wsTarget Is Nothing Then
        MsgBox "「集計表」シートが見つかりません。", vbCritical
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' 集計表のA1セルから日付取得
    If Not IsDate(wsTarget.Range("A1").Value) Then
        MsgBox "集計表のセルA1に有効な日付が入力されていません。", vbCritical
        GoTo CleanupAndExit
    End If
    targetDate = wsTarget.Range("A1").Value
    
    ' ソースシート取得（モールFR別シート）
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("モールFR別")
    If wsSource Is Nothing Then
        MsgBox "「モールFR別」シートが見つかりません。", vbCritical
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' ソーステーブル取得
    On Error Resume Next
    Set sourceTable = wsSource.ListObjects("_モールFR別b")
    If sourceTable Is Nothing Then
        MsgBox "「_モールFR別b」テーブルが見つかりません。", vbCritical
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' データ範囲取得
    If sourceTable.DataBodyRange Is Nothing Then
        MsgBox "「_モールFR別b」テーブルにデータがありません。", vbInformation
        GoTo CleanupAndExit
    End If
    Set sourceData = sourceTable.DataBodyRange
    
    ' 日付列のインデックス取得
    Dim dateColIndex As Long
    On Error Resume Next
    dateColIndex = sourceTable.ListColumns("日付").Index
    If Err.Number <> 0 Then
        MsgBox "「_モールFR別b」テーブルに「日付」列が見つかりません。", vbCritical
        Err.Clear
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
        MsgBox "日付 " & Format(targetDate, "yyyy/mm/dd") & " のデータが見つかりません。", vbInformation
        GoTo CleanupAndExit
    End If
    
    ' 各カテゴリ（モールF、モールR、合計）について処理
    Dim categoryIndex As Long
    Dim totalCategories As Long
    totalCategories = UBound(categoryInfo) + 1
    
    For categoryIndex = 0 To UBound(categoryInfo)
        Application.StatusBar = "転記処理中... (" & categoryInfo(categoryIndex)(0) & " - " & (categoryIndex + 1) & "/" & totalCategories & ")"
        
        Dim prefix As String
        Dim targetCol As Long
        prefix = categoryInfo(categoryIndex)(0)
        targetCol = categoryInfo(categoryIndex)(1)
        
        ' 各項目（9種類）について転記
        For k = 0 To UBound(columnSuffixes)
            Dim fullColumnName As String
            Dim colIndex As Long
            
            ' 完全な列名を構築
            fullColumnName = prefix & columnSuffixes(k)
            
            ' 列インデックス取得
            On Error Resume Next
            colIndex = sourceTable.ListColumns(fullColumnName).Index
            If Err.Number = 0 Then
                ' 転記実行
                wsTarget.Cells(targetRows(k), targetCol).Value = sourceData.Cells(sourceRow, colIndex).Value
                
                ' 不良率列の場合は書式設定（パーセント表示）
                If InStr(fullColumnName, "不良率") > 0 Then
                    wsTarget.Cells(targetRows(k), targetCol).NumberFormatLocal = "0.0%"
                End If
            Else
                Debug.Print "警告: 「" & fullColumnName & "」列が見つかりません。"
                Err.Clear
            End If
            On Error GoTo ErrorHandler
        Next k
    Next categoryIndex
    
    ' 正常終了
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    MsgBox "転記処理中に予期しないエラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "転記エラー"
    
CleanupAndExit:
    Application.StatusBar = False
End Sub