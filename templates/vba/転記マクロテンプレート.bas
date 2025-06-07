Attribute VB_Name = "m転記マクロテンプレート"
Sub 転記マクロテンプレート()
    '==========================================
    ' 転記マクロテンプレート
    ' 用途: テーブルから集計表へのデータ転記
    ' 
    ' カスタマイズ箇所:
    ' 1. シート名（転記元・転記先）
    ' 2. テーブル名
    ' 3. 転記条件（日付など）
    ' 4. 転記する列と行の対応
    '==========================================
    
    ' 変数宣言
    Dim wsTarget As Worksheet    ' 転記先シート（集計表など）
    Dim wsSource As Worksheet    ' 転記元シート
    Dim targetDate As Date       ' 転記条件（日付）
    Dim sourceTable As ListObject ' 転記元テーブル
    Dim sourceData As Range      ' テーブルのデータ範囲
    Dim i As Long, j As Long
    Dim processedCount As Long
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' 高速化設定（重要！）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' 進捗表示
    Application.StatusBar = "転記処理を開始します..."
    
    '==========================================
    ' カスタマイズ箇所1: シート名設定
    '==========================================
    ' 転記先シート取得
    Set wsTarget = ThisWorkbook.Worksheets("集計表")  ' ← シート名を変更
    
    ' 転記元シート取得
    Set wsSource = ThisWorkbook.Worksheets("データシート")  ' ← シート名を変更
    
    '==========================================
    ' カスタマイズ箇所2: 転記条件取得
    '==========================================
    ' 例: 集計表のA1セルから日付を取得
    If Not IsDate(wsTarget.Range("A1").Value) Then
        MsgBox "転記条件（日付）が正しく入力されていません。", vbCritical
        GoTo CleanupAndExit
    End If
    targetDate = wsTarget.Range("A1").Value
    
    '==========================================
    ' カスタマイズ箇所3: テーブル取得
    '==========================================
    Set sourceTable = wsSource.ListObjects("テーブル名")  ' ← テーブル名を変更
    
    ' データ範囲取得
    If sourceTable.DataBodyRange Is Nothing Then
        MsgBox "転記元テーブルにデータがありません。", vbCritical
        GoTo CleanupAndExit
    End If
    Set sourceData = sourceTable.DataBodyRange
    
    '==========================================
    ' カスタマイズ箇所4: 転記処理
    '==========================================
    ' テーブルの各行をループ
    For i = 1 To sourceData.Rows.Count
        ' 転記条件のチェック（例: 日付が一致）
        If sourceData.Cells(i, 1).Value = targetDate Then  ' ← 列番号を調整
            
            ' データ転記（例）
            wsTarget.Range("B10").Value = sourceData.Cells(i, 2).Value  ' ← 転記先・元を調整
            wsTarget.Range("C10").Value = sourceData.Cells(i, 3).Value
            
            processedCount = processedCount + 1
            
            ' 進捗表示（100件ごと）
            If processedCount Mod 100 = 0 Then
                Application.StatusBar = "転記処理中... " & processedCount & "件処理済み"
            End If
        End If
    Next i
    
    ' 完了メッセージ（デバッグ用）
    Debug.Print "転記完了: " & processedCount & "件"
    
CleanupAndExit:
    ' 設定を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "エラー"
End Sub