Attribute VB_Name = "mCommon"
Option Explicit

' ========================================
' モジュール名: mCommon
' 処理概要: 共通ヘルパー関数
' ========================================

' --------------------------------------------
' パターンに一致するテーブルを検索
' 引数: ws - 対象ワークシート
'       pattern - 検索パターン（部分一致）
' 戻り値: 一致したListObject、見つからない場合はNothing
' 例: FindTableByPattern(ws, "_完成品") → "_完成品", "_完成品2" 等にマッチ
' --------------------------------------------
Public Function FindTableByPattern(ws As Worksheet, pattern As String) As ListObject
    Dim tbl As ListObject
    For Each tbl In ws.ListObjects
        If InStr(tbl.Name, pattern) > 0 Then
            Set FindTableByPattern = tbl
            Exit Function
        End If
    Next tbl
    Set FindTableByPattern = Nothing
End Function
