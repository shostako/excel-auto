Attribute VB_Name = "mクエリ西暦更新"
Option Explicit

' ========================================
' マクロ名: クエリ西暦更新
' 処理概要: Power Queryの接続先DBを動的に変更
' 対象クエリ: 不良集計ゾーン別ADO
' 年の取得元: シート「不良集計ゾーン別ADO」セルG2
' 作成日: 2026-01-21
' ========================================

' DBパス設定
Private Const DB_BASE_PATH As String = "Z:\全社共有\オート事業部\日報\不良集計\不良集計表\"
Private Const DB_FILE_PREFIX As String = "不良調査表DB-"

' ============================================
' メイン処理: クエリ西暦更新
' 引数: queryName - 更新対象のクエリ名
' ============================================
Sub クエリ西暦更新(queryName As String)
    Dim ws As Worksheet
    Dim yearValue As Integer
    Dim newFormula As String
    Dim dbPath As String

    On Error GoTo ErrorHandler

    ' シートから年を取得
    Set ws = ThisWorkbook.Worksheets(queryName)

    If Not IsNumeric(ws.Range("G2").Value) Then
        MsgBox "シート「" & queryName & "」のG2に有効な年が入力されていません。", vbExclamation
        Exit Sub
    End If

    yearValue = CInt(ws.Range("G2").Value)

    ' DBパス構築
    dbPath = DB_BASE_PATH & yearValue & "年\" & DB_FILE_PREFIX & yearValue & ".accdb"

    ' M言語コード生成
    newFormula = BuildQueryFormula(dbPath)

    ' クエリ更新
    ThisWorkbook.Queries(queryName).Formula = newFormula

    ' クエリを更新して反映（バックグラウンド更新をオフにして同期実行）
    ws.ListObjects(1).QueryTable.Refresh BackgroundQuery:=False

    Application.StatusBar = "クエリ「" & queryName & "」を " & yearValue & "年 に更新しました。"

    Exit Sub

ErrorHandler:
    MsgBox "クエリ更新中にエラーが発生しました: " & Err.Description, vbCritical
    Application.StatusBar = False
End Sub

' ============================================
' 補助関数: M言語コード生成
' 引数: dbPath - DBファイルのフルパス
' 戻り値: Power Query M言語コード
' ============================================
Private Function BuildQueryFormula(dbPath As String) As String
    Dim m As String

    m = "let" & vbLf
    m = m & "    // 年間データ" & vbLf
    m = m & "    ソース = Access.Database(File.Contents(""" & dbPath & """), [CreateNavigationProperties=true])," & vbLf
    m = m & "    テーブル = ソース{[Schema="""",Item=""_不良集計ゾーン別""]}[Data]," & vbLf
    m = m & "    削除された他の列 = Table.SelectColumns(テーブル,{""ID"", ""日付"", ""品番"", ""ロット"", ""発見"", ""ゾーン"", ""番号"", ""数量"", ""差戻し""})," & vbLf
    m = m & "    変更された型 = Table.TransformColumnTypes(削除された他の列,{{""数量"", Int64.Type}, {""差戻し"", Int64.Type}})" & vbLf
    m = m & "in" & vbLf
    m = m & "    変更された型"

    BuildQueryFormula = m
End Function
