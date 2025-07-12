Attribute VB_Name = "mマクロテスト"
Option Explicit

' ========================================
' マクロ名: テスト実行マクロ
' 処理概要: プロジェクト規約に従った基本機能のテスト実行
' ソーステーブル: アクティブシート（テスト用データ）
' ターゲットテーブル: アクティブシート（結果出力用）
' 処理方式: 規約遵守テスト（Activate排除、最適化設定、エラーハンドリング）
' ========================================

Sub テスト実行マクロ()
    ' 最適化設定の保存
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    Dim origEnableEvents As Boolean
    Dim origDisplayAlerts As Boolean
    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation
    origEnableEvents = Application.EnableEvents
    origDisplayAlerts = Application.DisplayAlerts
    
    ' 最適化設定（画面ちらつき防止のため最重要）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' ステータスバー初期化
    Application.StatusBar = "マクロテストを開始します..."
    
    ' ============================================
    ' テスト1：基本的なセル操作（Activate使用禁止の確認）
    ' ============================================
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Activateメソッドは絶対に使わない（画面ちらつきの原因）
    ws.Range("A1").Value = "テスト開始時刻: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Range("A2").Value = "プロジェクト規約テスト"
    ws.Range("A3").Value = "Activateメソッド: 使用禁止"
    
    Application.StatusBar = "テスト1完了 - 基本セル操作"
    
    ' ============================================
    ' テスト2：配列処理のテスト（高速化確認）
    ' ============================================
    Dim testData(1 To 5, 1 To 2) As Variant
    Dim i As Long
    
    ' テストデータ作成（配列での処理）
    testData(1, 1) = "項目1": testData(1, 2) = "値1"
    testData(2, 1) = "項目2": testData(2, 2) = "値2"
    testData(3, 1) = "項目3": testData(3, 2) = "値3"
    testData(4, 1) = "項目4": testData(4, 2) = "値4"
    testData(5, 1) = "項目5": testData(5, 2) = "値5"
    
    ' 一括書き込み（セル単位ではなく配列で高速化）
    ws.Range("B1:C5").Value = testData
    
    Application.StatusBar = "テスト2完了 - 配列処理"
    
    ' ============================================
    ' テスト3：計算処理のテスト（Dictionary不使用の軽量版）
    ' ============================================
    Dim result As Double
    result = 0
    
    ' 簡単な計算処理（5項目の合計）
    For i = 1 To 5
        result = result + i * 10
    Next i
    
    ws.Range("D1").Value = "計算結果:"
    ws.Range("D2").Value = result
    
    Application.StatusBar = "テスト3完了 - 計算処理"
    
    ' ============================================
    ' テスト4：With文を使ったオブジェクト参照の効率化
    ' ============================================
    With ws.Range("E1:E3")
        .Cells(1, 1).Value = "With文テスト"
        .Cells(2, 1).Value = "効率的な参照"
        .Cells(3, 1).Value = "オブジェクト最適化"
    End With
    
    Application.StatusBar = "テスト4完了 - With文最適化"
    
    ' ============================================
    ' テスト5：処理結果の整理と完了表示
    ' ============================================
    ws.Range("A5").Value = "==========================="
    ws.Range("A6").Value = "全テスト完了時刻: " & Format(Now, "yyyy-mm-dd hh:mm:ss")
    ws.Range("A7").Value = "規約遵守項目:"
    ws.Range("A8").Value = "・Activateメソッド使用: なし"
    ws.Range("A9").Value = "・ScreenUpdating制御: あり"
    ws.Range("A10").Value = "・エラーハンドリング: あり"
    ws.Range("A11").Value = "・ステータスバー使用: あり"
    ws.Range("A12").Value = "==========================="
    
    ' 処理完了のステータスバー表示（1秒間表示）
    Application.StatusBar = "マクロテストが正常に完了しました"
    Application.Wait Now + TimeValue("00:00:01")
    
    GoTo Cleanup
    
ErrorHandler:
    ' エラー情報の詳細化
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear
    
    MsgBox "エラーが発生しました" & vbCrLf & _
           "エラー番号: " & errNum & vbCrLf & _
           "詳細: " & errDesc, vbCritical, "マクロテストエラー"
    
Cleanup:
    ' 設定を確実に復元（最重要）
    Application.StatusBar = False
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.DisplayAlerts = origDisplayAlerts
End Sub