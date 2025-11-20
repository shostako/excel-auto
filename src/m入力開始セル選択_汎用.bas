Attribute VB_Name = "m入力開始セル選択_汎用"
' ========================================
' マクロ名: 入力開始セル選択マクロ
' 処理概要: ブック内の全シートから指定テーブルを検索し、最初の空行を選択して画面スクロール
' ソーステーブル: 引数で指定されたListObject
' ターゲットテーブル: なし（カーソル移動のみ）
' ========================================

Sub 入力開始セル選択マクロ(ByVal targetTableName As String)
    ' -------------------------------------------------------------------------
    ' 機能　　：ブック内の全シートを探し、指定したテーブル名(ListObject)を見つけて
    ' 　　　　最初の空行(A列)を選択し、画面を適切にスクロールする
    ' 引数　　：targetTableName … 検索対象のテーブル名
    ' -------------------------------------------------------------------------

    ' ----- 変数宣言 -----
    Dim ws As Worksheet              ' ワークシート（検索用）
    Dim tbl As ListObject            ' 対象テーブル
    Dim firstEmptyRow As Long        ' 最初の空行番号
    Dim visibleRows As Long          ' 画面表示可能な行数
    Dim lastCell As Range            ' A列の最終セル
    Dim found As Boolean             ' テーブル発見フラグ

    found = False

    ' ============================================
    ' ブック内の全シートからテーブル名を探す
    ' ============================================
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set tbl = ws.ListObjects(targetTableName)
        On Error GoTo 0

        If Not tbl Is Nothing Then
            found = True
            Exit For
        End If
    Next ws

    ' テーブルが見つからない場合は処理を終了
    If Not found Then
        ' （必要に応じてエラーメッセージを出すなど）
        Exit Sub
    End If

    ' ============================================
    ' テーブル内で最初の空のセル(A列)を探す
    ' ============================================
    If Not tbl.DataBodyRange Is Nothing Then
        ' データ範囲が存在する場合
        Set lastCell = tbl.DataBodyRange.Columns(1).Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

        If Not lastCell Is Nothing Then
            ' A列の最後のセルが見つかった場合、その次の行を空の行とする
            firstEmptyRow = lastCell.Row + 1
        Else
            ' A列が完全に空白の場合、テーブル開始行の次の行を選択
            firstEmptyRow = tbl.Range.Row + 1
        End If
    Else
        ' テーブルが空の場合、テーブル開始行の次の行を選択
        firstEmptyRow = tbl.Range.Row + 1
    End If

    ' ============================================
    ' 空のセルを選択
    ' ============================================
    ws.Activate
    ws.Cells(firstEmptyRow, 1).Select

    ' ============================================
    ' 画面表示の調整：入力セルが画面の中央に来るようにスクロール
    ' ============================================
    '    表示可能な行数を大まかに計算
    visibleRows = Application.RoundUp(ActiveWindow.VisibleRange.Rows.Count / 2, 0)

    If firstEmptyRow > visibleRows - 3 Then
        ActiveWindow.ScrollRow = firstEmptyRow - visibleRows + 3
    Else
        ActiveWindow.ScrollRow = 1
    End If
End Sub
