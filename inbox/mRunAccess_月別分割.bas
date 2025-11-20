Attribute VB_Name = "mRunAccess_月別分割"
Option Explicit

' ========================================
' マクロ名: Access月別分割マクロ実行
' 処理概要: ExcelからAccessデータベースの月別分割処理を外部実行
' ソーステーブル: なし（Access DB操作）
' ターゲットテーブル: なし
' ========================================

Sub RunAccess_月別分割()
    ' ============================================
    ' 変数宣言：Access連携用オブジェクトとパス定義
    ' ============================================
    Dim acc As Object
    Dim dbPath As String: dbPath = "Z:\全社共有\オート事業部\日報\不良集計\不良集計表\2025年\不良調査表DB-2025.accdb"

    ' ============================================
    ' 初期設定：画面更新停止・キャンセル無効化
    ' ============================================
    Application.ScreenUpdating = False
    Application.EnableCancelKey = 0   ' xlDisable（Ctrl+Break無効化でAccess処理中の中断防止）
    Application.StatusBar = "月別分割: Access起動中..."

    On Error GoTo EH

    ' ============================================
    ' Access起動とDB接続：遅延バインディングで外部実行
    ' ============================================
    Set acc = CreateObject("Access.Application")
    acc.Visible = False
    acc.OpenCurrentDatabase dbPath, False
    Application.StatusBar = "月別分割: 実行中..."

    ' ============================================
    ' マクロ実行：関数→UIマクロの2段階フォールバック処理
    ' 理由：関数形式とUIマクロ形式の実装の違いに対応
    ' ============================================
    ' まず関数を試す → 失敗したらUIマクロ
    On Error Resume Next
    acc.Run "月別分割_Run"
    If Err.Number <> 0 Then
        Err.Clear
        acc.DoCmd.RunMacro "月別分割"
    End If
    On Error GoTo EH

    ' ============================================
    ' クリーンアップ：Access終了と設定復元
    ' ============================================
CleanUp:
    On Error Resume Next
    If Not acc Is Nothing Then
        acc.CloseCurrentDatabase
        acc.Quit
        Set acc = Nothing
    End If
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.EnableCancelKey = 1   ' xlInterrupt（通常状態に復帰）
    Exit Sub

    ' ============================================
    ' エラーハンドリング：ステータスバー表示とCleanUp実行
    ' ============================================
EH:
    Application.StatusBar = "月別分割: 失敗 (" & Err.Number & ") " & Err.Description
    Resume CleanUp
End Sub
