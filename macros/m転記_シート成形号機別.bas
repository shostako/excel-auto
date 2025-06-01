Attribute VB_Name = "m転記_シート成形号機別"
Option Explicit

' ==========================================
' 成形号機別からシートへの転記マクロ
' 「_成形号機別a」テーブルから各号機シートへデータを転記
' ==========================================
Sub 転記_シート成形号機別()
    ' ==========================================
    ' 変数宣言
    ' ==========================================
    Dim wsSource As Worksheet
    Dim sourceTable As ListObject
    Dim sourceData As Range
    Dim i As Long, j As Long
    Dim targetDate As Date
    Dim targetSheet As Worksheet
    Dim lastRow As Long
    Dim foundRow As Long
    Dim processedCount As Long
    Dim totalRows As Long
    
    ' 号機番号リスト
    Dim machineList() As Variant
    machineList = Array("SS01", "SS02", "SS03", "SS04", "SS05")
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' ==========================================
    ' 高速化設定
    ' ==========================================
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 進捗表示開始
    Application.StatusBar = "成形号機別シート転記処理を開始します..."
    
    ' ==========================================
    ' ソースシート・テーブル取得
    ' ==========================================
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
    Set sourceTable = wsSource.ListObjects("_成形号機別a")
    If sourceTable Is Nothing Then
        MsgBox "「_成形号機別a」テーブルが見つかりません。", vbCritical, "テーブルエラー"
        GoTo CleanupAndExit
    End If
    On Error GoTo ErrorHandler
    
    ' データ範囲取得
    If sourceTable.DataBodyRange Is Nothing Then
        MsgBox "「_成形号機別a」テーブルにデータがありません。", vbInformation, "データなし"
        GoTo CleanupAndExit
    End If
    Set sourceData = sourceTable.DataBodyRange
    
    ' ==========================================
    ' メイン処理: 各号機のデータを転記
    ' ==========================================
    ' 総処理数をカウント（進捗表示用）
    totalRows = sourceData.Rows.Count * (UBound(machineList) + 1)
    processedCount = 0
    
    ' 各号機について処理
    Dim machineIndex As Long
    For machineIndex = 0 To UBound(machineList)
        Dim machineName As String
        machineName = machineList(machineIndex)
        
        ' 転記先シートの存在確認
        On Error Resume Next
        Set targetSheet = ThisWorkbook.Worksheets(machineName)
        On Error GoTo ErrorHandler
        
        If targetSheet Is Nothing Then
            Debug.Print "警告: 号機「" & machineName & "」のシートが見つかりません。"
            processedCount = processedCount + sourceData.Rows.Count
            GoTo NextMachine
        End If
        
        Application.StatusBar = "転記処理中... (" & machineName & ")"
        
        ' 各日付のデータを処理
        For i = 1 To sourceData.Rows.Count
            processedCount = processedCount + 1
            
            ' 日付を取得
            targetDate = sourceData.Cells(i, 1).Value  ' 日付列
            
            ' 空白行はスキップ
            If IsEmpty(targetDate) Then
                GoTo NextDate
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
            
            ' データ転記（号機別のデータ）
            Select Case machineName
                Case "SS01"
                    ' SS01のデータ転記
                    TransferMachineData sourceData, i, targetSheet, foundRow, 2
                
                Case "SS02"
                    ' SS02のデータ転記
                    TransferMachineData sourceData, i, targetSheet, foundRow, 6
                
                Case "SS03"
                    ' SS03のデータ転記
                    TransferMachineData sourceData, i, targetSheet, foundRow, 10
                
                Case "SS04"
                    ' SS04のデータ転記
                    TransferMachineData sourceData, i, targetSheet, foundRow, 14
                
                Case "SS05"
                    ' SS05のデータ転記
                    TransferMachineData sourceData, i, targetSheet, foundRow, 18
            End Select
            
            ' 進捗更新
            If processedCount Mod 10 = 0 Then
                Application.StatusBar = "転記処理中... (" & processedCount & "/" & totalRows & ")"
            End If
            
NextDate:
        Next i
        
NextMachine:
        Set targetSheet = Nothing
    Next machineIndex
    
    ' 正常終了
    GoTo CleanupAndExit
    
ErrorHandler:
    ' エラー処理
    MsgBox "転記処理中にエラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "転記エラー"
    
CleanupAndExit:
    ' 後処理
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

' ==========================================
' 号機データ転記サブルーチン
' ==========================================
Private Sub TransferMachineData(sourceData As Range, sourceRow As Long, _
                               targetSheet As Worksheet, targetRow As Long, _
                               startCol As Long)
    ' 各項目のデータ転記（4項目）
    Dim colOffset As Long
    For colOffset = 0 To 3
        If Not IsEmpty(sourceData.Cells(sourceRow, startCol + colOffset).Value) Then
            targetSheet.Cells(targetRow, 2 + colOffset).Value = _
                sourceData.Cells(sourceRow, startCol + colOffset).Value
        Else
            targetSheet.Cells(targetRow, 2 + colOffset).Value = 0
        End If
    Next colOffset
End Sub