Attribute VB_Name = "m転記_シート塗装品番別"
Option Explicit

' ==========================================
' 塗装品番別からシートへの転記マクロ
' 「_塗装品番別a」テーブルから品番別シートへデータを転記
' ==========================================
Sub 転記_シート塗装品番別()
    ' ==========================================
    ' 変数宣言
    ' ==========================================
    Dim wsSource As Worksheet
    Dim sourceTable As ListObject
    Dim sourceData As Range
    Dim i As Long, j As Long
    Dim targetDate As Date
    Dim nickname As String
    Dim targetSheet As Worksheet
    Dim lastRow As Long
    Dim foundRow As Long
    Dim processedCount As Long
    Dim totalRows As Long
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' ==========================================
    ' 高速化設定
    ' ==========================================
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 進捗表示開始
    Application.StatusBar = "塗装品番別シート転記処理を開始します..."
    
    ' ==========================================
    ' ソースシート・テーブル取得
    ' ==========================================
    ' 塗装品番別シート取得
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("塗装品番別")
    If wsSource Is Nothing Then
        MsgBox "「塗装品番別」シートが見つかりません。", vbCritical, "シートエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' ソーステーブル取得
    On Error Resume Next
    Set sourceTable = wsSource.ListObjects("_塗装品番別a")
    If sourceTable Is Nothing Then
        MsgBox "「_塗装品番別a」テーブルが見つかりません。", vbCritical, "テーブルエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' データ範囲取得
    If sourceTable.DataBodyRange Is Nothing Then
        MsgBox "「_塗装品番別a」テーブルにデータがありません。", vbInformation, "データなし"
        GoTo CleanupAndExit
    End If
    Set sourceData = sourceTable.DataBodyRange
    
    ' ==========================================
    ' メイン処理: 各品番のデータを転記
    ' ==========================================
    ' 総行数をカウント（進捗表示用）
    totalRows = sourceData.Rows.Count
    processedCount = 0
    
    ' 各行のデータを処理
    For i = 1 To sourceData.Rows.Count
        ' 日付とニックネームを取得
        targetDate = sourceData.Cells(i, 1).Value  ' 日付列
        nickname = sourceData.Cells(i, 2).Value     ' ニックネーム列
        
        ' 空白行はスキップ
        If nickname = "" Or IsEmpty(nickname) Or IsEmpty(targetDate) Then
            GoTo NextRow
        End If
        
        processedCount = processedCount + 1
        Application.StatusBar = "転記処理中... (" & processedCount & "/" & totalRows & ") " & nickname
        
        ' ニックネームに対応するシートの存在確認
        On Error Resume Next
        Set targetSheet = ThisWorkbook.Worksheets(nickname)
        On Error GoTo ErrorHandler
        
        If targetSheet Is Nothing Then
            ' シートが存在しない場合はスキップ（エラーにはしない）
            Debug.Print "警告: ニックネーム「" & nickname & "」のシートが見つかりません。"
            GoTo NextRow
        End If
        
        ' 転記先シートで該当日付の行を検索
        lastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row
        foundRow = 0
        
        For j = 2 To lastRow  ' ヘッダー行をスキップ
            If targetSheet.Cells(j, 1).Value = targetDate Then
                foundRow = j
                Exit For
            End If
        Next j
        
        ' 該当日付が見つからない場合は新規行追加
        If foundRow = 0 Then
            foundRow = lastRow + 1
            targetSheet.Cells(foundRow, 1).Value = targetDate
        End If
        
        ' データ転記（集計値）
        ' 実績合計
        If Not IsEmpty(sourceData.Cells(i, 3).Value) Then
            targetSheet.Cells(foundRow, 2).Value = sourceData.Cells(i, 3).Value
        Else
            targetSheet.Cells(foundRow, 2).Value = 0
        End If
        
        ' 不良合計
        If Not IsEmpty(sourceData.Cells(i, 4).Value) Then
            targetSheet.Cells(foundRow, 3).Value = sourceData.Cells(i, 4).Value
        Else
            targetSheet.Cells(foundRow, 3).Value = 0
        End If
        
        ' 稼動時間合計
        If Not IsEmpty(sourceData.Cells(i, 5).Value) Then
            targetSheet.Cells(foundRow, 4).Value = sourceData.Cells(i, 5).Value
        Else
            targetSheet.Cells(foundRow, 4).Value = 0
        End If
        
        ' 段取時間合計
        If Not IsEmpty(sourceData.Cells(i, 6).Value) Then
            targetSheet.Cells(foundRow, 5).Value = sourceData.Cells(i, 6).Value
        Else
            targetSheet.Cells(foundRow, 5).Value = 0
        End If
        
NextRow:
        Set targetSheet = Nothing
    Next i
    
    ' 正常終了
    GoTo CleanupAndExit
    
ErrorHandler:
    ' エラー処理
    MsgBox "転記処理中にエラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "処理中のニックネーム: " & nickname, vbCritical, "転記エラー"
    
CleanupAndExit:
    ' 後処理
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub