VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ========================================
' モジュール名: Sheet1（シートモジュール）
' 処理概要: セル変更時のフィルター・日付表示自動実行
' 設置場所: VBE → 対象シートのコードウィンドウに貼り付け
' 使用方法: フィルター機能を使うシートごとにこのコードを設置
' ========================================
Option Explicit

' --------------------------------------------
' セル変更イベント
' A3:E3のセル変更時に対応するマクロを自動実行
' --------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim eventsOff As Boolean
    On Error GoTo ErrHandler

    ' 複数セル同時変更は無視
    If Target.CountLarge > 1 Then Exit Sub
    ' A3:E3以外の変更は無視
    If Intersect(Target, Me.Range("A3:E3")) Is Nothing Then Exit Sub

    ' イベント無効化（無限ループ防止）
    Application.EnableEvents = False
    eventsOff = True

    ' 変更されたセルに応じてマクロを呼び出し
    Select Case Target.Address(False, False)
        Case "A3"
            ' 日付セル：値があれば日付表示マクロ実行
            If Len(Target.Value) > 0 Then Call 日付表示
        Case "B3"
            ' 完成品フィルター：値があればフィルター実行
            If Len(Target.Value) > 0 Then Call 完成品フィルター
        Case "C3"
            ' 側板フィルター：値があればフィルター実行
            If Len(Target.Value) > 0 Then Call 側板フィルター
        Case "D3"
            ' 小部品フィルター：値があればフィルター実行
            If Len(Target.Value) > 0 Then Call 小部品フィルター
        Case "E3"
            ' 項目フィルター：「全項目」も含め常に実行
            Call 項目フィルター
    End Select

CleanExit:
    ' イベント再有効化
    If eventsOff Then Application.EnableEvents = True
    Exit Sub

ErrHandler:
    ' エラー時もイベント再有効化してから再スロー
    If eventsOff Then Application.EnableEvents = True
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
