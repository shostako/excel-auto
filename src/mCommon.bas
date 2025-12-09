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

' --------------------------------------------
' 項目フィルターモードを判定
' 引数: filterItem - E3の値
' 戻り値: フィルターモード文字列
'   "none" - フィルターなし（全項目/空）
'   "exact" - 完全一致（単一値またはカンマ区切り）
' --------------------------------------------
Public Function GetItemFilterMode(filterItem As String) As String
    Select Case filterItem
        Case "全項目", ""
            GetItemFilterMode = "none"
        Case Else
            GetItemFilterMode = "exact"
    End Select
End Function

' --------------------------------------------
' 項目フィルター条件に一致するか判定
' 引数: cellValue - 項目列のセル値
'       filterMode - フィルターモード（GetItemFilterModeの戻り値）
'       filterItem - E3の値（カンマ区切り対応）
' 戻り値: True=条件に一致, False=一致しない
' カンマ区切りの場合、各要素にOR条件で完全一致判定
' --------------------------------------------
Public Function MatchItemFilter(cellValue As String, filterMode As String, filterItem As String) As Boolean
    Dim items As Variant
    Dim item As Variant

    Select Case filterMode
        Case "none"
            MatchItemFilter = True
        Case "exact"
            ' カンマ区切りの場合は分割してOR判定
            If InStr(filterItem, ",") > 0 Then
                items = Split(filterItem, ",")
                MatchItemFilter = False
                For Each item In items
                    If cellValue = Trim(CStr(item)) Then
                        MatchItemFilter = True
                        Exit For
                    End If
                Next item
            Else
                ' 単一値の場合は完全一致
                MatchItemFilter = (cellValue = filterItem)
            End If
        Case Else
            MatchItemFilter = True
    End Select
End Function
