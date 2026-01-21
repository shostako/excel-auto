Attribute VB_Name = "m番号転送ADO"
' ========================================
' マクロ名: 番号転送ADO
' 処理概要: ExcelテーブルからAccessデータベースへ番号データをADO経由で全とっかえ転送
' ソーステーブル: シート「番号S」テーブル「_番号S」
' ターゲットテーブル: Accessデータベース「不良調査表DB-{年}.accdb」テーブル「_番号」
' 転送フィールド: 番号、モード、発生
' 処理方式: 全削除→全挿入（毎回全とっかえ）
' 西暦取得元: シート「不良集計ゾーン別ADO」セルG2
' ========================================

Option Explicit

Sub 番号転送ADO()
    ' ----- 変数宣言 -----
    ' ADO関連
    Dim conn As Object               ' ADODB.Connection（データベース接続）
    Dim cmd As Object                ' ADODB.Command（SQLコマンド実行用）
    Dim rs As Object                 ' ADODB.Recordset（レコードセット操作用）

    ' テーブル関連
    Dim tbl As ListObject            ' 転送元Excelテーブル（_番号S）

    ' ループ・カウンター
    Dim i As Long, j As Long         ' ループカウンター（i:行、j:列）
    Dim rowCount As Long             ' 転送対象行数
    Dim successCount As Long         ' 転送成功カウント
    Dim skippedCount As Long         ' 空白行スキップカウント

    ' フラグ
    Dim transStarted As Boolean      ' トランザクション開始フラグ

    ' DBパス関連
    Dim yearValue As Integer         ' 西暦（G2セルから取得）
    Dim dbPath As String             ' DBファイルパス

    ' 転送対象フィールド定義
    Dim targetFields As Variant      ' 転送するフィールド名の配列
    targetFields = Array("番号", "モード", "発生")

    ' トランザクション開始フラグを初期化
    transStarted = False

    ' 進捗表示用
    Application.ScreenUpdating = False
    Application.StatusBar = "番号データ転送処理を開始します..."

    ' エラー処理
    On Error GoTo ErrorHandler

    ' ============================================
    ' テーブル取得と検証
    ' ============================================
    Set tbl = ActiveSheet.ListObjects("_番号S")
    If tbl Is Nothing Then
        MsgBox "テーブル「_番号S」が見つかりません。", vbExclamation
        GoTo CleanExit
    End If

    rowCount = tbl.ListRows.Count
    If rowCount = 0 Then
        Application.StatusBar = "転送するデータがありません。"
        ' 3秒後にステータスバーをクリア
        Application.OnTime Now + TimeValue("00:00:03"), "ステータスバークリア"
        GoTo CleanExit
    End If

    ' ============================================
    ' 西暦取得・DBパス動的構築
    ' ============================================
    Application.StatusBar = "DBパスを構築しています..."

    yearValue = ThisWorkbook.Worksheets("不良集計ゾーン別ADO").Range("G2").Value

    If yearValue < 2020 Or yearValue > 2100 Then
        MsgBox "西暦の値が不正です: " & yearValue & vbCrLf & _
               "「不良集計ゾーン別ADO」シートのG2セルを確認してください。", vbExclamation
        GoTo CleanExit
    End If

    dbPath = "Z:\全社共有\オート事業部\日報\不良集計\不良集計表\" & _
             yearValue & "年\不良調査表DB-" & yearValue & ".accdb"

    ' DB存在確認
    If Dir(dbPath) = "" Then
        MsgBox "DBファイルが見つかりません:" & vbCrLf & dbPath, vbExclamation
        GoTo CleanExit
    End If

    ' ============================================
    ' ADO接続とコマンドオブジェクト設定
    ' ============================================
    Set conn = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")

    ' 接続文字列（動的パス使用）
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=" & dbPath & ";"

    ' コマンドオブジェクトの設定
    Set cmd.ActiveConnection = conn
    cmd.CommandType = 1  ' 1 = adCmdText

    ' ============================================
    ' フィールド名と列インデックスの対応マップを作成
    ' ============================================
    Dim fieldIndices As Object
    Set fieldIndices = CreateObject("Scripting.Dictionary")
    Dim fieldName As Variant

    For j = 1 To tbl.HeaderRowRange.Columns.Count
        fieldName = tbl.HeaderRowRange.Cells(1, j).Value
        fieldIndices.Add CStr(fieldName), j
    Next j

    ' ============================================
    ' 指定フィールドの存在チェック（Excel側）
    ' ============================================
    Dim missingFields As String
    missingFields = ""

    For Each fieldName In targetFields
        If Not fieldIndices.Exists(CStr(fieldName)) Then
            If missingFields <> "" Then missingFields = missingFields & ", "
            missingFields = missingFields & fieldName
        End If
    Next

    If missingFields <> "" Then
        MsgBox "以下のフィールドがExcelテーブルに見つかりません:" & vbCrLf & missingFields, vbExclamation
        GoTo CleanExit
    End If

    ' ============================================
    ' テーブル再作成（DROP → CREATE）
    ' ============================================
    Application.StatusBar = "転送先テーブルを準備しています..."

    ' テーブルが存在する場合は削除
    If TableExistsInAccess(conn, "_番号") Then
        conn.Execute "DROP TABLE [_番号]"
    End If

    ' テーブル新規作成（IDカウンターもリセット）
    Dim sqlCreate As String
    sqlCreate = "CREATE TABLE [_番号] (" & _
                "[ID] AUTOINCREMENT PRIMARY KEY, " & _
                "[番号] TEXT(50), " & _
                "[モード] TEXT(50), " & _
                "[発生] TEXT(50));"

    conn.Execute sqlCreate

    ' 作成確認
    If Not TableExistsInAccess(conn, "_番号") Then
        MsgBox "テーブル「_番号」の作成に失敗しました。", vbExclamation
        GoTo CleanExit
    End If

    ' ============================================
    ' SQL用のフィールドリスト作成（INSERT文用）
    ' ============================================
    Dim fieldList As String
    fieldList = ""
    For Each fieldName In targetFields
        If fieldList <> "" Then fieldList = fieldList & ", "
        fieldList = fieldList & "[" & fieldName & "]"
    Next

    ' ============================================
    ' トランザクション開始
    ' ============================================
    conn.BeginTrans
    transStarted = True  ' トランザクション開始フラグを設定

    ' ============================================
    ' データ転送実行ループ
    ' ============================================
    Application.StatusBar = "データを転送しています (0/" & rowCount & ")..."
    successCount = 0
    skippedCount = 0  ' スキップ行カウント初期化

    For i = 1 To rowCount
        ' 空白行チェック
        If IsRowEmpty(tbl, i, targetFields, fieldIndices) Then
            skippedCount = skippedCount + 1
            ' 空白行はスキップ
            GoTo NextRow
        End If

        ' この行のINSERT SQL文を作成（指定フィールドのみ）
        Dim sqlInsert As String
        sqlInsert = "INSERT INTO [_番号] (" & fieldList & ") VALUES (" & _
                    GetSelectedValues(tbl, i, targetFields, fieldIndices) & ");"

        ' SQL実行
        conn.Execute sqlInsert

        ' 成功カウント増加
        successCount = successCount + 1

NextRow:
        ' 進捗表示を更新（10行ごと）
        If i Mod 10 = 0 Or i = rowCount Then
            Application.StatusBar = "データを転送しています (" & i & "/" & rowCount & ")..."
        End If
    Next i

    ' ============================================
    ' トランザクションをコミット
    ' ============================================
    conn.CommitTrans
    transStarted = False  ' トランザクション完了フラグを設定

    ' 完了メッセージをステータスバーに表示
    Application.StatusBar = successCount & "件のデータを転送しました。" & skippedCount & "件の空白行をスキップしました。"

    ' 3秒後にステータスバーをクリア
    Application.OnTime Now + TimeValue("00:00:03"), "ステータスバークリア"

CleanExit:
    ' ============================================
    ' リソース解放
    ' ============================================
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.state = 1 Then rs.Close
    End If

    If Not conn Is Nothing Then
        If conn.state = 1 Then
            If transStarted Then  ' トランザクションが開始されている場合のみロールバック
                conn.RollbackTrans
            End If
            conn.Close
        End If
    End If

    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
    Set fieldIndices = Nothing

    ' 画面更新を再開
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    ' ============================================
    ' エラー処理
    ' ============================================
    Dim errMsg As String
    errMsg = "エラーが発生しました:" & vbCrLf & _
             "説明: " & Err.Description & vbCrLf & _
             "エラー番号: " & Err.Number

    ' トランザクションのロールバック
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.state = 1 Then
            If transStarted Then  ' トランザクションが開始されている場合のみロールバック
                conn.RollbackTrans
            End If
        End If
    End If

    MsgBox errMsg, vbCritical, "番号データ転送エラー"
    Resume CleanExit

End Sub

' ============================================
' 関数: TableExistsInAccess
' 機能: Accessテーブルが存在するかどうかを確認する
' 引数: conn - ADODB.Connection オブジェクト
'       tableName - 確認対象のテーブル名
' 戻り値: True = テーブル存在、False = テーブル不在
' ============================================
Function TableExistsInAccess(conn As Object, tableName As String) As Boolean
    Dim tempRS As Object

    ' エラーハンドリングを有効化
    On Error Resume Next

    ' テーブルに対してクエリを実行してみる
    Set tempRS = conn.Execute("SELECT TOP 1 * FROM [" & tableName & "]")

    ' エラーが発生したかどうかをチェック
    If Err.Number = 0 Then
        ' エラーがない場合、テーブルは存在する
        TableExistsInAccess = True
    Else
        ' エラーが発生した場合、テーブルは存在しない
        TableExistsInAccess = False
    End If

    ' リソース解放
    If Not tempRS Is Nothing Then
        If tempRS.state = 1 Then tempRS.Close
        Set tempRS = Nothing
    End If

    ' エラーハンドリングをリセット
    Err.Clear
    On Error GoTo 0
End Function

' ============================================
' 関数: IsRowEmpty
' 機能: 行が空かどうかをチェックする
' 引数: tbl - ListObject
'       rowIndex - 行インデックス
'       targetFields - チェック対象フィールド配列
'       fieldIndices - フィールド名と列インデックスの対応Dictionary
' 戻り値: True = 空行、False = データあり
' ============================================
Function IsRowEmpty(tbl As ListObject, rowIndex As Long, targetFields As Variant, fieldIndices As Object) As Boolean
    Dim i As Integer
    Dim fieldName As String
    Dim colIndex As Integer
    Dim cellValue As Variant

    ' デフォルトは空として扱う
    IsRowEmpty = True

    ' 少なくとも1つのフィールドに値があれば空ではない
    For i = 0 To UBound(targetFields)
        fieldName = targetFields(i)
        colIndex = fieldIndices(fieldName)
        cellValue = tbl.ListRows(rowIndex).Range(1, colIndex).Value

        ' 値が存在するかチェック（Empty、Null、空文字列以外）
        If Not IsEmpty(cellValue) And Not IsNull(cellValue) Then
            If VarType(cellValue) = vbString Then
                ' 文字列の場合、空でないことを確認
                If Len(Trim(cellValue)) > 0 Then
                    IsRowEmpty = False
                    Exit Function
                End If
            Else
                ' 数値や日付など他のデータ型の場合
                IsRowEmpty = False
                Exit Function
            End If
        End If
    Next i
End Function

' ============================================
' 関数: GetSelectedValues
' 機能: 指定されたフィールドの値のみを取得（INSERT VALUES文用）
' 引数: tbl - ListObject
'       rowIndex - 行インデックス
'       targetFields - 取得対象フィールド配列
'       fieldIndices - フィールド名と列インデックスの対応Dictionary
' 戻り値: SQL VALUES句用の値文字列（カンマ区切り）
' ============================================
Function GetSelectedValues(tbl As ListObject, rowIndex As Long, targetFields As Variant, fieldIndices As Object) As String
    Dim result As String
    Dim i As Integer
    Dim fieldName As String
    Dim colIndex As Integer
    Dim cellValue As Variant

    result = ""

    For i = 0 To UBound(targetFields)
        If i > 0 Then result = result & ", "

        fieldName = targetFields(i)
        colIndex = fieldIndices(fieldName)
        cellValue = tbl.ListRows(rowIndex).Range(1, colIndex).Value

        ' データ型に応じたフォーマット
        If IsEmpty(cellValue) Or IsNull(cellValue) Then
            result = result & "NULL"
        ElseIf IsDate(cellValue) Then
            ' 日付形式
            result = result & "#" & Format(cellValue, "yyyy/mm/dd") & "#"
        ElseIf IsNumeric(cellValue) Then
            ' 数値
            result = result & cellValue
        Else
            ' テキスト（シングルクォートをエスケープ）
            result = result & "'" & Replace(cellValue, "'", "''") & "'"
        End If
    Next i

    GetSelectedValues = result
End Function
