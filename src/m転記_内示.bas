Attribute VB_Name = "m転記_内示"
Option Explicit

' ========================================
' 転記_内示
' ========================================
' 概要：
'   「_内示抽出」テーブルのデータを「_成形展開」テーブルに転記する
'
' 処理フロー：
'   1. 転記対象列（1〜31、翌月、翌々月、計）のデータ行をクリア
'   2. 品番による行照合（_内示抽出の品番列 ⇔ _成形展開の製品品番列）
'   3. 列名による列照合（両テーブルの列名で一致）
'   4. 交点セルに値を転記（1品番に複数行ある場合は全て転記）
'   5. 存在しない品番（転記データが0でない）をエラー報告
'
' 転記対象列：
'   1, 2, 3, ..., 30, 31, 計, 翌月, 翌々月
' ========================================

Sub 転記_内示()
    Application.StatusBar = "内示データの転記を開始します..."

    On Error GoTo ErrorHandler

    ' ========== オブジェクト参照取得 ==========
    Dim ws内示 As Worksheet
    Dim ws展開 As Worksheet
    Dim lo内示 As ListObject
    Dim lo展開 As ListObject

    Set ws内示 = ThisWorkbook.Worksheets("内示")
    Set ws展開 = ThisWorkbook.Worksheets("展開")
    Set lo内示 = ws内示.ListObjects("_内示抽出")
    Set lo展開 = ws展開.ListObjects("_成形展開")

    ' ========== 画面更新停止 ==========
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Application.StatusBar = "転記範囲をクリアしています..."

    ' ========== 転記対象列のデータ行をクリア ==========
    Dim i As Long
    For i = 1 To 31
        lo展開.ListColumns(CStr(i)).DataBodyRange.ClearContents
    Next i
    lo展開.ListColumns("計").DataBodyRange.ClearContents
    lo展開.ListColumns("翌月").DataBodyRange.ClearContents
    lo展開.ListColumns("翌々月").DataBodyRange.ClearContents

    Application.StatusBar = "列マッピングを作成しています..."

    ' ========== _内示抽出テーブルの列インデックス取得 ==========
    Dim col品番 As Long, col計 As Long, col翌月 As Long, col翌々月 As Long
    Dim col日付(1 To 31) As Long

    col品番 = lo内示.ListColumns("品番").Index
    col計 = lo内示.ListColumns("計").Index
    col翌月 = lo内示.ListColumns("翌月").Index
    col翌々月 = lo内示.ListColumns("翌々月").Index

    For i = 1 To 31
        col日付(i) = lo内示.ListColumns(CStr(i)).Index
    Next i

    ' ========== _成形展開テーブルの列インデックス取得 ==========
    Dim tgt_col製品品番 As Long, tgt_col計 As Long, tgt_col翌月 As Long, tgt_col翌々月 As Long
    Dim tgt_col日付(1 To 31) As Long

    tgt_col製品品番 = lo展開.ListColumns("製品品番").Index
    tgt_col計 = lo展開.ListColumns("計").Index
    tgt_col翌月 = lo展開.ListColumns("翌月").Index
    tgt_col翌々月 = lo展開.ListColumns("翌々月").Index

    For i = 1 To 31
        tgt_col日付(i) = lo展開.ListColumns(CStr(i)).Index
    Next i

    ' ========== _成形展開テーブルの製品品番マップ作成 ==========
    Dim 品番行マップ As Object
    Set 品番行マップ = CreateObject("Scripting.Dictionary")

    Dim 展開データ As Variant
    展開データ = lo展開.DataBodyRange.Value

    Dim r As Long
    For r = 1 To UBound(展開データ, 1)
        Dim 製品品番 As String
        製品品番 = CStr(展開データ(r, tgt_col製品品番))
        If 製品品番 <> "" Then
            If Not 品番行マップ.exists(製品品番) Then
                品番行マップ.Add 製品品番, New Collection
            End If
            品番行マップ(製品品番).Add r
        End If
    Next r

    Application.StatusBar = "データを転記しています..."

    ' ========== データ転記処理 ==========
    Dim 未転記品番リスト As String
    未転記品番リスト = ""

    Dim 内示データ As Variant
    内示データ = lo内示.DataBodyRange.Value

    Dim 処理件数 As Long
    処理件数 = 0

    For r = 1 To UBound(内示データ, 1)
        Dim src品番 As String
        src品番 = CStr(内示データ(r, col品番))

        ' 100件ごとに進捗表示
        処理件数 = 処理件数 + 1
        If 処理件数 Mod 100 = 0 Then
            Application.StatusBar = "データを転記しています... (" & 処理件数 & "/" & UBound(内示データ, 1) & ")"
        End If

        ' _成形展開テーブルに該当品番が存在するか確認
        If 品番行マップ.exists(src品番) Then
            ' 該当する全ての行に転記
            Dim targetRows As Collection
            Set targetRows = 品番行マップ(src品番)

            Dim targetRow As Variant
            For Each targetRow In targetRows
                Dim targetRowIdx As Long
                targetRowIdx = CLng(targetRow)

                ' 日付列（1〜31）を個別に書き込み
                For i = 1 To 31
                    lo展開.DataBodyRange.Cells(targetRowIdx, tgt_col日付(i)).Value = 内示データ(r, col日付(i))
                Next i

                ' 計
                lo展開.DataBodyRange.Cells(targetRowIdx, tgt_col計).Value = 内示データ(r, col計)

                ' 翌月
                lo展開.DataBodyRange.Cells(targetRowIdx, tgt_col翌月).Value = 内示データ(r, col翌月)

                ' 翌々月
                lo展開.DataBodyRange.Cells(targetRowIdx, tgt_col翌々月).Value = 内示データ(r, col翌々月)
            Next targetRow
        Else
            ' 品番が_成形展開テーブルに存在しない場合、転記データが全て0かチェック
            Dim hasNonZero As Boolean
            hasNonZero = False

            ' 日付列チェック
            For i = 1 To 31
                If 内示データ(r, col日付(i)) <> 0 Then
                    hasNonZero = True
                    Exit For
                End If
            Next i

            ' 計、翌月、翌々月チェック
            If Not hasNonZero Then
                If 内示データ(r, col計) <> 0 Or 内示データ(r, col翌月) <> 0 Or 内示データ(r, col翌々月) <> 0 Then
                    hasNonZero = True
                End If
            End If

            ' 0でないデータがある場合のみリストに追加
            If hasNonZero Then
                If 未転記品番リスト <> "" Then
                    未転記品番リスト = 未転記品番リスト & vbCrLf
                End If
                未転記品番リスト = 未転記品番リスト & src品番
            End If
        End If
    Next r

    Application.StatusBar = "テーブルの列幅を調整しています..."

    ' ========== 列幅調整 ==========
    ' 幅12の列
    lo展開.ListColumns("客先").Range.ColumnWidth = 12
    lo展開.ListColumns("製品品番").Range.ColumnWidth = 12
    lo展開.ListColumns("成形品番").Range.ColumnWidth = 12

    ' 幅5の列
    lo展開.ListColumns("型番").Range.ColumnWidth = 5
    lo展開.ListColumns("区分").Range.ColumnWidth = 5
    lo展開.ListColumns("仕様").Range.ColumnWidth = 5
    lo展開.ListColumns("成形号機").Range.ColumnWidth = 5
    lo展開.ListColumns("進捗").Range.ColumnWidth = 5
    lo展開.ListColumns("先月在庫").Range.ColumnWidth = 5

    ' 幅3の列（1〜31）
    For i = 1 To 31
        lo展開.ListColumns(CStr(i)).Range.ColumnWidth = 3
    Next i

    ' 幅5の列（計、翌月、翌々月）
    lo展開.ListColumns("計").Range.ColumnWidth = 5
    lo展開.ListColumns("翌月").Range.ColumnWidth = 5
    lo展開.ListColumns("翌々月").Range.ColumnWidth = 5

    ' ========== 縮小して全体を表示 ==========
    With lo展開.Range
        .ShrinkToFit = True
    End With

    ' ========== エラーメッセージ表示 ==========
    If 未転記品番リスト <> "" Then
        MsgBox "以下の品番が展開シートに存在しないため転記できませんでした:" & vbCrLf & vbCrLf & _
               未転記品番リスト, vbExclamation, "転記エラー"
    End If

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
End Sub
