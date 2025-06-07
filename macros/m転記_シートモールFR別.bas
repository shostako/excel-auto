Attribute VB_Name = "m転記_シートモールFR別"
Option Explicit

' ==========================================
' モールFR別からシートへの転記マクロ
' 「_モールFR別a」テーブルからF/R別シートへデータを転記
' ==========================================
Sub 転記_シートモールFR別()
    ' ==========================================
    ' 変数宣言
    ' ==========================================
    Dim wsSource As Worksheet
    Dim sourceTable As ListObject
    Dim sourceData As Range
    Dim dateDict As Object
    Dim i As Long, j As Long
    Dim targetDate As Date
    Dim targetSheet As Worksheet
    Dim lastRow As Long
    Dim foundRow As Long
    Dim processedCount As Long
    Dim totalRows As Long
    
    ' F/R種別の配列
    Dim frTypes() As Variant
    Dim sheetNames() As Variant
    frTypes = Array("モールF", "モールR")
    sheetNames = Array("モールF", "モールR")
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' ==========================================
    ' 高速化設定
    ' ==========================================
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 進捗表示開始
    Application.StatusBar = "モールFR別シート転記処理を開始します..."
    
    ' ==========================================
    ' ソースシート・テーブル取得
    ' ==========================================
    ' モールFR別シート取得
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("モールFR別")
    If wsSource Is Nothing Then
        MsgBox "「モールFR別」シートが見つかりません。", vbCritical, "シートエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' ソーステーブル取得
    On Error Resume Next
    Set sourceTable = wsSource.ListObjects("_モールFR別a")
    If sourceTable Is Nothing Then
        MsgBox "「_モールFR別a」テーブルが見つかりません。", vbCritical, "テーブルエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' データ範囲取得
    If sourceTable.DataBodyRange Is Nothing Then
        MsgBox "「_モールFR別a」テーブルにデータがありません。", vbInformation, "データなし"
        GoTo CleanupAndExit
    End If
    Set sourceData = sourceTable.DataBodyRange
    
    ' ==========================================
    ' 日付インデックス作成（高速検索用）
    ' ==========================================
    Set dateDict = CreateObject("Scripting.Dictionary")
    
    ' 日付列のインデックス取得
    Dim dateColIndex As Long
    On Error Resume Next
    dateColIndex = sourceTable.ListColumns("日付").Index
    If Err.Number <> 0 Then
        MsgBox "「_モールFR別a」テーブルに「日付」列が見つかりません。", vbCritical, "列エラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' 日付とその行番号をDictionaryに格納
    For i = 1 To sourceData.Rows.Count
        targetDate = sourceData.Cells(i, dateColIndex).Value
        If Not IsEmpty(targetDate) Then
            dateDict(targetDate) = i
        End If
    Next i
    
    ' ==========================================
    ' メイン処理: 各F/Rタイプのデータを転記
    ' ==========================================
    totalRows = dateDict.Count * UBound(frTypes) + 1
    processedCount = 0
    
    ' 各F/Rタイプについて処理
    Dim k As Long
    For k = 0 To UBound(frTypes)
        ' 転記先シートの存在確認
        On Error Resume Next
        Set targetSheet = ThisWorkbook.Worksheets(sheetNames(k))
        On Error GoTo ErrorHandler
        
        If targetSheet Is Nothing Then
            Debug.Print "警告: シート「" & sheetNames(k) & "」が見つかりません。"
            GoTo NextType
        End If
        
        Application.StatusBar = "転記処理中... (" & sheetNames(k) & ")"
        
        ' 各日付のデータを処理
        Dim dateKey As Variant
        For Each dateKey In dateDict.Keys
            processedCount = processedCount + 1
            i = dateDict(dateKey)
            
            ' 転記先シートで該当日付の行を検索
            lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
            foundRow = 0
            
            For j = 2 To lastRow
                If targetSheet.Cells(j, 1).Value = dateKey Then
                    foundRow = j
                    Exit For
                End If
            Next j
            
            ' 該当日付が見つからない場合は新規行追加
            If foundRow = 0 Then
                foundRow = lastRow + 1
                targetSheet.Cells(foundRow, 1).Value = dateKey
            End If
            
            ' データ転記
            ' 日実績
            Dim colName As String
            Dim colIndex As Long
            colName = frTypes(k) & "日実績"
            On Error Resume Next
            colIndex = sourceTable.ListColumns(colName).Index
            If Err.Number = 0 Then
                If Not IsEmpty(sourceData.Cells(i, colIndex).Value) Then
                    targetSheet.Cells(foundRow, 2).Value = sourceData.Cells(i, colIndex).Value
                End If
            End If
            Err.Clear
            
            ' 日不良数
            colName = frTypes(k) & "日不良数"
            colIndex = sourceTable.ListColumns(colName).Index
            If Err.Number = 0 Then
                If Not IsEmpty(sourceData.Cells(i, colIndex).Value) Then
                    targetSheet.Cells(foundRow, 3).Value = sourceData.Cells(i, colIndex).Value
                End If
            End If
            Err.Clear
            
            ' 日出来高サイクル
            colName = frTypes(k) & "日出来高ｻｲｸﾙ"
            colIndex = sourceTable.ListColumns(colName).Index
            If Err.Number = 0 Then
                If Not IsEmpty(sourceData.Cells(i, colIndex).Value) Then
                    targetSheet.Cells(foundRow, 4).Value = sourceData.Cells(i, colIndex).Value
                End If
            End If
            Err.Clear
            On Error GoTo ErrorHandler
            
            ' 進捗更新
            If processedCount Mod 10 = 0 Then
                Application.StatusBar = "転記処理中... (" & processedCount & "/" & totalRows & ")"
            End If
        Next dateKey
        
NextType:
        Set targetSheet = Nothing
    Next k
    
    ' 正常終了
    GoTo CleanupAndExit
    
ErrorHandler:
    ' エラー処理
    MsgBox "転記処理中にエラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "転記エラー"
    
CleanupAndExit:
    ' 後処理
    Set dateDict = Nothing
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub