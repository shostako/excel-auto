Attribute VB_Name = "m番号転送ADO"
' ========================================
' マクロ名: 番号転送ADO
' 処理概要: ExcelテーブルからAccessデータベースへ番号データをADO経由で転送
' ソーステーブル: シート「現在のシート」テーブル「_番号S」
' ターゲットテーブル: Accessデータベース「不良調査表DB-2025.accdb」テーブル「_番号」
' 転送フィールド: 番号、モード、発生
' 重複チェック: 番号、モードのキー組み合わせ
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

    ' SQL・キー関連
    Dim sqlCheck As String           ' 既存データ確認SQL
    Dim key As String                ' 重複チェック用キー文字列
    Dim keyFields As String          ' 重複チェック対象フィールド（カンマ区切り）

    ' Dictionary関連
    Dim existingDict As Object       ' 既存データ格納用Dictionary（重複チェック用）

    ' フラグ
    Dim transStarted As Boolean      ' トランザクション開始フラグ
    Dim tableExists As Boolean       ' テーブル存在確認フラグ

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

    ' 重複チェック用のキーフィールド設定
    keyFields = "番号,モード"

    ' ============================================
    ' Dictionary オブジェクト作成（重複チェック用）
    ' ============================================
    Set existingDict = CreateObject("Scripting.Dictionary")

    ' ============================================
    ' ADO接続とコマンドオブジェクト設定
    ' ============================================
    Set conn = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")

    ' 接続文字列
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=Z:¥全社共有¥オート事業部¥日報¥不良集計¥不良集計表¥2025年¥不良調査表DB-2025.accdb;"

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
    ' 転送先テーブルの存在確認と構造検証
    ' ============================================
    Application.StatusBar = "転送先テーブルの構造を確認しています..."

    ' テーブルの存在を確認
    tableExists = TableExistsInAccess(conn, "_番号")

    If Not tableExists Then
        ' ============================================
        ' テーブルが存在しない場合は作成
        ' ============================================
        Application.StatusBar = "転送先テーブルを作成しています..."

        Dim sqlCreate As String
        sqlCreate = "CREATE TABLE [_番号] (" & _
                    "[ID] AUTOINCREMENT PRIMARY KEY, " & _
                    "[番号] TEXT(50), " & _
                    "[モード] TEXT(50), " & _
                    "[発生] TEXT(50));"

        conn.Execute sqlCreate

        ' 作成したテーブルが存在するか再確認
        tableExists = TableExistsInAccess(conn, "_番号")
        If Not tableExists Then
            MsgBox "テーブル「_番号」の作成に失敗しました。", vbExclamation
            GoTo CleanExit
        End If
    Else
        ' ============================================
        ' 既存テーブルの構造確認（Access側フィールド検証）
        ' ============================================
        Set rs = conn.Execute("SELECT TOP 1 * FROM [_番号]")

        ' アクセスのフィールド一覧
        Dim accessFields As Object
        Set accessFields = CreateObject("Scripting.Dictionary")
        Dim f As Object

        For Each f In rs.Fields
            If f.Name <> "ID" Then  ' IDフィールドを除外
                accessFields.Add f.Name, True
            End If
        Next

        ' 指定フィールドがアクセスにあるか確認
        missingFields = ""
        For Each fieldName In targetFields
            If Not accessFields.Exists(CStr(fieldName)) Then
                If missingFields <> "" Then missingFields = missingFields & ", "
                missingFields = missingFields & fieldName
            End If
        Next

        If missingFields <> "" Then
            MsgBox "以下のフィールドがAccess側テーブルに見つかりません:" & vbCrLf & missingFields & vbCrLf & _
                   "テーブル構造が異なる可能性があります。", vbExclamation
            GoTo CleanExit
        End If

        rs.Close
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
    ' 既存データの確認（重複転送防止）
    ' ============================================
    Application.StatusBar = "既存データを確認しています..."
    sqlCheck = "SELECT " & Replace(keyFields, ",", ", ") & " FROM [_番号]"

    Set rs = conn.Execute(sqlCheck)

    ' 既存データをDictionaryに格納
    If Not rs.EOF Then
        rs.MoveFirst
        Do Until rs.EOF
            key = ""
            Dim fieldArray As Variant
            Dim fieldIndex As Integer

            fieldArray = Split(keyFields, ",")
            For fieldIndex = 0 To UBound(fieldArray)
                If Not IsNull(rs(Trim(fieldArray(fieldIndex)))) Then
                    key = key & rs(Trim(fieldArray(fieldIndex))) & "|"
                Else
                    key = key & "NULL|"
                End If
            Next fieldIndex

            If Not existingDict.Exists(key) Then
                existingDict.Add key, True
            End If
            rs.MoveNext
        Loop
    End If
    rs.Close

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

        ' キー値を作成して重複チェック
        key = CreateKeyFromRow(tbl, i, keyFields, fieldIndices)

        If Not existingDict.Exists(key) Then
            ' この行のINSERT SQL文を作成（指定フィールドのみ）
            Dim sqlInsert As String
            sqlInsert = "INSERT INTO [_番号] (" & fieldList & ") VALUES (" & _
                        GetSelectedValues(tbl, i, targetFields, fieldIndices) & ");"

            ' SQL実行
            conn.Execute sqlInsert

            ' 成功カウント増加
            successCount = successCount + 1

            ' Dictionaryに追加
            existingDict.Add key, True
        End If

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
    Set existingDict = Nothing
    Set fieldIndices = Nothing

    If Not accessFields Is Nothing Then
        Set accessFields = Nothing
    End If

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
    Dim hasValue As Boolean

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
' 関数: CreateKeyFromRow
' 機能: テーブル行からキー値を作成する（重複チェック用）
' 引数: tbl - ListObject
'       rowIndex - 行インデックス
'       keyFields - キーフィールド（カンマ区切り文字列）
'       fieldIndices - フィールド名と列インデックスの対応Dictionary
' 戻り値: キー文字列（パイプ区切り）
' ============================================
Function CreateKeyFromRow(tbl As ListObject, rowIndex As Long, keyFields As String, fieldIndices As Object) As String
    Dim key As String
    Dim fieldArray As Variant
    Dim i As Integer
    Dim fieldName As String
    Dim colIndex As Integer
    Dim cellValue As Variant

    key = ""
    fieldArray = Split(keyFields, ",")

    For i = 0 To UBound(fieldArray)
        fieldName = Trim(fieldArray(i))

        If fieldIndices.Exists(fieldName) Then
            colIndex = fieldIndices(fieldName)
            cellValue = tbl.ListRows(rowIndex).Range(1, colIndex).Value

            If IsEmpty(cellValue) Or IsNull(cellValue) Then
                key = key & "NULL|"
            Else
                key = key & CStr(cellValue) & "|"
            End If
        Else
            key = key & "MISSING|"
        End If
    Next i

    CreateKeyFromRow = key
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
