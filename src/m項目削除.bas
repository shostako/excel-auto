Attribute VB_Name = "m項目削除"
Option Explicit

' ========================================
' マクロ名: 項目削除
' 処理概要: 指定テーブルで指定した項目行を全て削除
' 引数:
'   tableName - 対象テーブル名（パターン検索）
'   targetItem - 削除する項目値
' 冪等性: 該当する行がなければ通知のみ
' 注意: フィルターがかかっている場合は自動解除される
' ========================================

Sub 項目削除(tableName As String, targetItem As String)
    ' --------------------------------------------
    ' 変数宣言
    ' --------------------------------------------
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim itemCol As ListColumn
    Dim dataArr As Variant
    Dim i As Long
    Dim deleteCount As Long
    Dim colItem As Long

    ' targetItem行インデックスを格納する配列
    Dim targetRows() As Long
    Dim targetCount As Long

    ' --------------------------------------------
    ' 画面更新・計算・イベント抑制（高速化）
    ' --------------------------------------------
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo ErrorHandler

    Set ws = ActiveSheet

    ' --------------------------------------------
    ' テーブル取得（完全一致）
    ' --------------------------------------------
    Dim tmpTbl As ListObject
    For Each tmpTbl In ws.ListObjects
        If tmpTbl.Name = tableName Then
            Set tbl = tmpTbl
            Exit For
        End If
    Next tmpTbl
    If tbl Is Nothing Then
        MsgBox "「" & tableName & "」テーブルが見つかりません", vbExclamation
        GoTo Cleanup
    End If

    ' --------------------------------------------
    ' フィルター解除（行削除のため）
    ' --------------------------------------------
    If Not tbl.AutoFilter Is Nothing Then
        If tbl.AutoFilter.FilterMode Then
            tbl.AutoFilter.ShowAllData
        End If
    End If

    ' --------------------------------------------
    ' 列インデックス取得
    ' --------------------------------------------
    colItem = tbl.ListColumns("項目").Index

    ' --------------------------------------------
    ' targetItem行のインデックスを収集
    ' --------------------------------------------
    Set itemCol = tbl.ListColumns("項目")
    dataArr = itemCol.DataBodyRange.Value

    targetCount = 0
    ReDim targetRows(1 To UBound(dataArr, 1))

    For i = 1 To UBound(dataArr, 1)
        If dataArr(i, 1) = targetItem Then
            targetCount = targetCount + 1
            targetRows(targetCount) = i
        End If
    Next i

    If targetCount = 0 Then
        MsgBox "「" & targetItem & "」行が見つかりません", vbInformation
        GoTo Cleanup
    End If

    ' --------------------------------------------
    ' 下から上へ処理（行削除によるインデックスずれ防止）
    ' --------------------------------------------
    deleteCount = 0

    For i = targetCount To 1 Step -1
        Dim rowIdx As Long
        rowIdx = targetRows(i)

        ' 行削除
        tbl.ListRows(rowIdx).Delete

        deleteCount = deleteCount + 1
    Next i

    ' --------------------------------------------
    ' 完了（正常終了時はメッセージなし）
    ' --------------------------------------------

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "エラーが発生しました" & vbCrLf & _
           "エラー番号: " & Err.Number & vbCrLf & _
           "詳細: " & Err.Description, vbCritical
End Sub
