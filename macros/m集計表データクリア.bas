Attribute VB_Name = "m集計表データクリア"
Sub 集計表データクリア()
    ' 変数宣言
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim i As Long
    Dim deletedCount As Long
    Dim totalSheets As Long
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' 高速化設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 進捗表示開始
    Application.StatusBar = "集計表データのクリア処理を開始します..."
    
    ' ワークブック設定
    Set wb = ThisWorkbook
    
    ' 削除対象シート名
    sheetNames = Array("日別集計_モールFR別", "集計表_TG作業者別", "集計表_TG品番別", _
                      "集計表_モールFR別", "集計表_加工作業者別", "集計表_加工品番別", _
                      "集計表_塗装品番別", "集計表_流出廃棄")
    
    ' 初期化
    deletedCount = 0
    totalSheets = UBound(sheetNames) + 1
    
    ' 各シートの削除処理
    For i = 0 To UBound(sheetNames)
        Application.StatusBar = "シート削除中... (" & (i + 1) & "/" & totalSheets & ")"
        
        On Error Resume Next
        Set ws = wb.Sheets(sheetNames(i))
        If Not ws Is Nothing Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            deletedCount = deletedCount + 1
            Set ws = Nothing
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ' 処理結果の表示
    If deletedCount > 0 Then
        MsgBox "集計表データをクリアしました。" & vbCrLf & _
               "削除されたシート数: " & deletedCount & "/" & totalSheets, vbInformation, "処理完了"
    Else
        MsgBox "削除対象のシートが見つかりませんでした。", vbInformation, "処理完了"
    End If
    
    ' 正常終了
    GoTo CleanupAndExit
    
ErrorHandler:
    MsgBox "データクリア処理中にエラーが発生しました。" & vbCrLf & _
           "エラー内容: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "エラー"
    
CleanupAndExit:
    ' 後処理
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.StatusBar = False  ' ステータスバーをクリア
End Sub