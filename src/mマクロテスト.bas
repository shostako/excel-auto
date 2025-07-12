Option Explicit

' ========================================
' マクロ名: マクロテスト
' 処理概要: Claude Codeの基本動作確認とプロジェクト規約テスト
' ソーステーブル: アクティブシート（動的範囲）
' ターゲットテーブル: 同シート「テスト結果」範囲
' 処理方式: 基本的なセル操作とメッセージ表示のテスト
' ========================================

Sub マクロテスト()
    ' 最適化設定の保存
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    Dim origEnableEvents As Boolean
    Dim origDisplayAlerts As Boolean
    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation
    origEnableEvents = Application.EnableEvents
    origDisplayAlerts = Application.DisplayAlerts
    
    ' 最適化設定（プロジェクト標準）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' ステータスバー初期化
    Application.StatusBar = "マクロテストを開始します..."
    
    ' ============================================
    ' テスト処理1：現在時刻の記録
    ' ============================================
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 重要：Activateは使わない（画面ちらつきの原因）
    ws.Range("A1").Value = "マクロテスト実行結果"
    ws.Range("A2").Value = "実行日時："
    ws.Range("B2").Value = Now
    
    Application.StatusBar = "マクロテスト：基本データ設定完了..."
    
    ' ============================================
    ' テスト処理2：簡単な計算処理
    ' ============================================
    Dim i As Long
    For i = 1 To 5
        ws.Range("A" & (i + 3)).Value = "テスト項目" & i
        ws.Range("B" & (i + 3)).Value = i * 10
        
        ' 進捗表示（少ないデータだが、パターンを示す）
        Application.StatusBar = "マクロテスト：項目" & i & "/5 処理中..."
    Next i
    
    ' ============================================
    ' テスト処理3：結果の整理
    ' ============================================
    ws.Range("A9").Value = "処理完了ステータス："
    ws.Range("B9").Value = "正常終了"
    ws.Range("A10").Value = "処理方式："
    ws.Range("B10").Value = "Activate未使用・高速化対応"
    
    ' 処理完了のステータスバー表示
    Application.StatusBar = "マクロテストが正常に完了しました"
    Application.Wait Now + TimeValue("00:00:01")
    
    GoTo Cleanup
    
ErrorHandler:
    ' エラー情報の詳細化
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear
    
    MsgBox "マクロテストでエラーが発生しました" & vbCrLf & _
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