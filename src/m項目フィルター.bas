Attribute VB_Name = "m項目フィルター"
Option Explicit

' ========================================
' マクロ名: 項目フィルター
' 処理概要: E3セルの値でオートフィルターを適用
' 参照セル: E3（項目フィルター条件）
' フィルター対象: 全テーブルの「項目」列
' 条件分岐:
'   - 「全項目」→ フィルター解除
'   - カンマ区切り → 各要素でOR条件フィルター
'   - 単一値 → その値でフィルター
'   - 「合計」「稼働日」行 → 常に表示（フィルター条件に自動追加）
' 複合フィルター: 他のフィルターは維持（オートフィルター方式）
' ========================================

Sub 項目フィルター()
    ' --------------------------------------------
    ' 変数宣言
    ' --------------------------------------------
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim filterItem As String
    Dim colIndex As Long
    Dim filterArray() As String
    Dim i As Long

    ' 対象テーブル名
    Dim tables As Variant
    tables = Array("_完成品", "_core", "_slitter", "_acf")

    ' --------------------------------------------
    ' 画面更新・イベント抑制（高速化）
    ' --------------------------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    Set ws = ActiveSheet

    ' --------------------------------------------
    ' フィルター条件の取得
    ' --------------------------------------------
    filterItem = ws.Range("E3").Value

    ' --------------------------------------------
    ' 全テーブル：項目列にオートフィルター
    ' --------------------------------------------
    Dim tblName As Variant
    For Each tblName In tables
        Set tbl = FindTableByPattern(ws, CStr(tblName))
        colIndex = tbl.ListColumns("項目").Index

        If filterItem = "全項目" Or filterItem = "" Then
            ' フィルター解除
            If tbl.AutoFilter.FilterMode Then
                ' 該当列のフィルターのみ解除
                tbl.Range.AutoFilter Field:=colIndex
            End If
        Else
            ' フィルター条件を配列に変換（「合計」「稼働日」を自動追加）
            filterArray = BuildItemFilterArray(filterItem)

            ' フィルター適用（複数値）
            tbl.Range.AutoFilter Field:=colIndex, _
                Criteria1:=filterArray, _
                Operator:=xlFilterValues
        End If
    Next tblName

    ' --------------------------------------------
    ' 垂直スクロールのみ先頭に移動（水平位置は維持）
    ' --------------------------------------------
    ActiveWindow.ScrollRow = 1

    ' --------------------------------------------
    ' 終了処理
    ' --------------------------------------------
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "エラーが発生しました" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "詳細: " & Err.Description, vbCritical
End Sub

' ========================================
' フィルター条件配列を構築（「合計」「稼働日」を自動追加）
' ========================================
Public Function BuildItemFilterArray(ByVal filterItem As String) As String()
    Dim baseItems() As String
    Dim result() As String
    Dim i As Long
    Dim count As Long
    Dim hasGoukei As Boolean
    Dim hasKadoubi As Boolean

    ' カンマ区切りで分割
    baseItems = Split(filterItem, ",")

    ' 既存の「合計」「稼働日」チェック
    hasGoukei = False
    hasKadoubi = False
    For i = 0 To UBound(baseItems)
        baseItems(i) = Trim(baseItems(i))
        If baseItems(i) = "合計" Then hasGoukei = True
        If baseItems(i) = "稼働日" Then hasKadoubi = True
    Next i

    ' 結果配列のサイズを計算
    count = UBound(baseItems) + 1
    If Not hasGoukei Then count = count + 1
    If Not hasKadoubi Then count = count + 1

    ReDim result(0 To count - 1)

    ' 基本項目をコピー
    For i = 0 To UBound(baseItems)
        result(i) = baseItems(i)
    Next i

    ' 「合計」「稼働日」を追加（まだない場合）
    Dim idx As Long
    idx = UBound(baseItems) + 1
    If Not hasGoukei Then
        result(idx) = "合計"
        idx = idx + 1
    End If
    If Not hasKadoubi Then
        result(idx) = "稼働日"
    End If

    BuildItemFilterArray = result
End Function
