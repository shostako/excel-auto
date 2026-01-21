Attribute VB_Name = "mロット数量データ転送ADO"
Option Explicit

' ========================================
' マクロ名: ロット数量データ転送ADO
' 処理概要: Excelテーブル「_ロット数量S」からAccessデータベースへADO経由でデータ転送
' ソーステーブル: シート「(Active)」テーブル「_ロット数量S」
' ターゲットテーブル: Accessデータベース「不良調査表DB-{年}.accdb」テーブル「_ロット数量」
' 作成日: 不明
' 更新日: 2025-01-12（重複時に日付上書き機能追加）
'
' 処理の流れ:
' 1. Excelテーブルからフィールド構造とデータ範囲を取得
' 2. データから含まれる年を抽出し、全DBファイルの存在を事前確認
' 3. DBファイルが1つでも存在しない場合は処理中止（データは残る）
' 4. 年ごとにADO接続でAccessデータベースに接続
' 5. 既存データをDictionaryにロードして重複チェック準備
' 6. 重複時は日付を上書きUPDATE（跨り生産対応）
' 7. トランザクション制御でバッチINSERT実行（50件ごと）
' 8. 全年の転送成功後、ソーステーブルの指定フィールドをクリア
'
' 技術的特徴:
' - 年別自動振り分け機能（日付から年を抽出し適切なDBに転送）
' - 事前DBファイル存在チェック（All or Nothing方式）
' - ADO接続とトランザクション管理
' - Dictionary重複チェック（日付を除外したキー生成）
' - 重複時は日付を上書きUPDATE（後ろの日付に合わせる）
' - バッチ処理によるパフォーマンス最適化（BATCH_SIZE = 50）
' - 動的フィールドマッピング（ヘッダー行から自動取得）
' - 日付範囲フィルター（±7日マージン）による既存データロード効率化
' - 空白行スキップ機能
' - エラー位置特定機能（errorLocation変数）
' - SQLインジェクション対策（シングルクォートエスケープ）
' ========================================

' ============================================
' 定数定義：バッチサイズ、タイムアウト、DBパス設定
' ============================================
Const BATCH_SIZE As Long = 50        ' バッチ処理サイズ（この件数ごとにコミット）
Const CONNECTION_TIMEOUT As Long = 30 ' 接続タイムアウト(秒)
Const COMMAND_TIMEOUT As Long = 60   ' コマンドタイムアウト(秒)

' DBパス設定（年別振り分け用）
Const DB_BASE_PATH As String = "Z:\全社共有\オート事業部\日報\不良集計\不良集計表\"
Const DB_FILE_PREFIX As String = "不良調査表DB-"

' ============================================
' 補助関数: BuildDBPath
' 役割: 年からDBファイルパスを動的生成
' 引数: yearValue - 年（Integer）
' 戻り値: Z:\...\{年}年\不良調査表DB-{年}.accdb
' ============================================
Function BuildDBPath(yearValue As Integer) As String
    BuildDBPath = DB_BASE_PATH & yearValue & "年\" & DB_FILE_PREFIX & yearValue & ".accdb"
End Function

' ============================================
' 補助関数: DBFileExists
' 役割: DBファイルの存在確認
' 引数: dbPath - DBファイルパス
' 戻り値: True = 存在、False = 不存在
' ============================================
Function DBFileExists(dbPath As String) As Boolean
    On Error Resume Next
    DBFileExists = (Dir(dbPath) <> "")
    On Error GoTo 0
End Function

' ============================================
' 補助関数: ExtractYearsFromData
' 役割: テーブルデータから含まれる年を抽出
' 引数: tbl - ListObject、dateColIndex - 日付列インデックス
' 戻り値: Dictionary (key=年, value=True)
' ============================================
Function ExtractYearsFromData(tbl As ListObject, dateColIndex As Integer) As Object
    Dim years As Object
    Set years = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim dateValue As Variant
    Dim yearValue As Integer

    For i = 1 To tbl.ListRows.Count
        dateValue = tbl.ListRows(i).Range(1, dateColIndex).Value
        If IsDate(dateValue) Then
            yearValue = Year(CDate(dateValue))
            If Not years.Exists(yearValue) Then
                years.Add yearValue, True
            End If
        End If
    Next i

    Set ExtractYearsFromData = years
End Function

' ============================================
' 補助関数: CheckAllDBsExist
' 役割: 全DBファイルの存在確認
' 引数: years - 年のDictionary
' 戻り値: 空文字列=全て存在、それ以外=存在しない年のリスト
' ============================================
Function CheckAllDBsExist(years As Object) As String
    Dim missingYears As String
    Dim yearKey As Variant
    Dim dbPath As String

    missingYears = ""
    For Each yearKey In years.Keys
        dbPath = BuildDBPath(CInt(yearKey))
        If Not DBFileExists(dbPath) Then
            If missingYears <> "" Then missingYears = missingYears & ", "
            missingYears = missingYears & yearKey & "年"
        End If
    Next yearKey

    CheckAllDBsExist = missingYears
End Function

' ============================================
' 補助関数: GroupRowsByYear
' 役割: 年別に行番号をグループ化
' 引数: tbl - ListObject、dateColIndex - 日付列インデックス
' 戻り値: Dictionary (key=年, value=Collection of 行番号)
' ============================================
Function GroupRowsByYear(tbl As ListObject, dateColIndex As Integer) As Object
    Dim yearGroups As Object
    Set yearGroups = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim dateValue As Variant
    Dim yearValue As Integer

    For i = 1 To tbl.ListRows.Count
        dateValue = tbl.ListRows(i).Range(1, dateColIndex).Value
        If IsDate(dateValue) Then
            yearValue = Year(CDate(dateValue))
            If Not yearGroups.Exists(yearValue) Then
                Set yearGroups(yearValue) = New Collection
            End If
            yearGroups(yearValue).Add i
        End If
    Next i

    Set GroupRowsByYear = yearGroups
End Function

' ============================================
' 補助関数: BuildUpdateSQL
' 役割: 重複レコードの日付更新用UPDATE文を生成
' 引数: tbl - ListObject、rowIndex - 行番号、
'       fieldIndices - フィールドインデックス、keyFields - キーフィールド
' 戻り値: UPDATE文
' ============================================
Function BuildUpdateSQL(tbl As ListObject, rowIndex As Long, fieldIndices As Object, keyFields As String) As String
    Dim sql As String
    Dim dateIndex As Integer
    Dim dateValue As Variant
    Dim fArray As Variant
    Dim ii As Integer
    Dim fName As String
    Dim colIndex As Integer
    Dim cellValue As Variant
    Dim whereClause As String

    dateIndex = fieldIndices("日付")
    dateValue = tbl.ListRows(rowIndex).Range(1, dateIndex).Value

    sql = "UPDATE [_ロット数量] SET [日付] = #" & Format(dateValue, "yyyy/mm/dd") & "# WHERE "

    ' WHERE条件を構築
    ' Access側のフィールド型に合わせる（ロット数量のみ数値、他はTEXT）
    whereClause = ""
    fArray = Split(keyFields, ",")

    For ii = 0 To UBound(fArray)
        If ii > 0 Then whereClause = whereClause & " AND "

        fName = Trim(fArray(ii))
        colIndex = fieldIndices(fName)
        cellValue = tbl.ListRows(rowIndex).Range(1, colIndex).Value

        If IsEmpty(cellValue) Or IsNull(cellValue) Then
            whereClause = whereClause & "[" & fName & "] IS NULL"
        ElseIf fName = "ロット数量" Then
            ' ロット数量のみ数値型
            whereClause = whereClause & "[" & fName & "] = " & cellValue
        Else
            ' 品番、品番末尾、注番月、ロット、工程はすべてTEXT型
            whereClause = whereClause & "[" & fName & "] = '" & Replace(CStr(cellValue), "'", "''") & "'"
        End If
    Next ii

    BuildUpdateSQL = sql & whereClause
End Function

Sub ロット数量データ転送ADO()
    ' ============================================
    ' 変数宣言：ADO接続、テーブル操作、処理制御用の変数を定義
    ' ============================================
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    Dim tbl As ListObject
    Dim i As Long, j As Long
    Dim rowCount As Long
    Dim sqlCheck As String
    Dim key As String
    Dim existingDict As Object
    Dim successCount As Long
    Dim updatedCount As Long  ' 日付更新件数
    Dim skippedCount As Long ' 空白行のスキップカウント用
    Dim keyFields As String
    Dim transStarted As Boolean ' トランザクション開始フラグ
    Dim tableExists As Boolean ' テーブル存在フラグ
    Dim batchCounter As Long   ' バッチ処理用カウンター
    Dim startTime As Double    ' 処理時間計測用
    Dim recordCount As Long    ' レコード数カウント用
    Dim errorLocation As String ' エラー発生箇所特定用

    ' 年別振り分け用変数
    Dim years As Object
    Dim yearGroups As Object
    Dim yearKey As Variant
    Dim dbPath As String
    Dim missingDBs As String
    Dim totalSuccess As Long
    Dim totalUpdated As Long  ' 年別合計更新件数
    Dim totalDuplicateSkipped As Long  ' 重複スキップ件数（日付同一）
    Dim duplicateSkippedCount As Long  ' 年別重複スキップ件数
    Dim rowNumbers As Collection
    Dim rowNum As Variant
    Dim dateIndex As Integer

    ' ============================================
    ' 初期設定：処理時間計測開始、フラグ初期化、対象フィールド定義
    ' ============================================
    startTime = Timer

    ' トランザクション開始フラグを初期化
    transStarted = False
    batchCounter = 0
    totalSuccess = 0
    totalUpdated = 0
    totalDuplicateSkipped = 0

    ' 年リストを初期化（月別分割マクロとの連携用）
    InitTransferredYears

    ' 転送対象のフィールドを明示的に指定
    Dim targetFields As Variant
    targetFields = Array("日付", "品番", "品番末尾", "注番月", "ロット", "工程", "ロット数量")

    ' 進捗表示用
    Application.ScreenUpdating = False
    Application.StatusBar = "ADO転送処理を開始します..."

    ' エラー処理
    On Error GoTo ErrorHandler

    ' 処理位置を記録
    errorLocation = "テーブル取得"

    ' ============================================
    ' データ検証：ソーステーブルの取得と存在確認、行数チェック
    ' ============================================
    Set tbl = ActiveSheet.ListObjects("_ロット数量S")
    If tbl Is Nothing Then
        Application.StatusBar = "テーブル「_ロット数量S」が見つかりません。"
        Application.Wait Now + TimeValue("00:00:03") ' 3秒間表示
        GoTo CleanExit
    End If

    rowCount = tbl.ListRows.Count
    If rowCount = 0 Then
        Application.StatusBar = "転送するデータがありません。"
        Application.Wait Now + TimeValue("00:00:03") ' 3秒間表示
        GoTo CleanExit
    End If

    ' 重複チェック用のキーフィールド設定（日付を除外）
    keyFields = "品番,品番末尾,注番月,ロット,工程,ロット数量"

    ' 処理位置を記録
    errorLocation = "フィールドマッピング"

    ' ============================================
    ' フィールドマッピング：Excelヘッダー行から列インデックスを取得
    ' Dictionary使用で動的にフィールド名→列番号を対応付け
    ' ============================================
    Dim fieldIndices As Object
    Set fieldIndices = CreateObject("Scripting.Dictionary")
    Dim fieldName As Variant

    For j = 1 To tbl.HeaderRowRange.Columns.Count
        fieldName = tbl.HeaderRowRange.Cells(1, j).Value
        fieldIndices.Add CStr(fieldName), j
    Next j

    ' 指定フィールドの存在チェック
    Dim missingFields As String
    missingFields = ""

    For Each fieldName In targetFields
        If Not fieldIndices.Exists(CStr(fieldName)) Then
            If missingFields <> "" Then missingFields = missingFields & ", "
            missingFields = missingFields & fieldName
        End If
    Next

    If missingFields <> "" Then
        Application.StatusBar = "以下のフィールドがExcelテーブルに見つかりません: " & missingFields
        Application.Wait Now + TimeValue("00:00:05") ' 5秒間表示
        GoTo CleanExit
    End If

    ' 日付列インデックス取得
    dateIndex = fieldIndices("日付")

    ' ============================================
    ' Phase 1: 事前チェック - 年抽出とDBファイル存在確認
    ' ============================================
    errorLocation = "年抽出・DBチェック"

    Application.StatusBar = "データから年を抽出しています..."

    ' データから含まれる年を抽出
    Set years = ExtractYearsFromData(tbl, dateIndex)

    If years.Count = 0 Then
        Application.StatusBar = "転送するデータに有効な日付がありません。"
        Application.Wait Now + TimeValue("00:00:03")
        GoTo CleanExit
    End If

    ' 全DBファイルの存在チェック
    Application.StatusBar = "DBファイルの存在を確認しています..."
    missingDBs = CheckAllDBsExist(years)

    If missingDBs <> "" Then
        Application.StatusBar = "以下のDBファイルが見つかりません: " & missingDBs
        MsgBox "以下のDBファイルが見つかりません:" & vbCrLf & missingDBs & vbCrLf & _
               vbCrLf & "データは転送されませんでした。" & vbCrLf & _
               "必要なDBファイルを作成してから再実行してください。", vbExclamation, "DBファイル不足"
        GoTo CleanExit  ' データはクリアされない
    End If

    ' ============================================
    ' Phase 2: 年別転送処理
    ' ============================================
    errorLocation = "年別データ転送"

    ' 年別に行番号をグループ化
    Set yearGroups = GroupRowsByYear(tbl, dateIndex)

    ' SQL用のフィールドリスト作成
    Dim fieldList As String
    fieldList = ""
    For Each fieldName In targetFields
        If fieldList <> "" Then fieldList = fieldList & ", "
        fieldList = fieldList & "[" & fieldName & "]"
    Next

    ' 年ごとに処理
    For Each yearKey In yearGroups.Keys
        dbPath = BuildDBPath(CInt(yearKey))

        Application.StatusBar = yearKey & "年のデータを転送中..."

        ' ADO接続オブジェクト作成
        Set conn = CreateObject("ADODB.Connection")
        Set cmd = CreateObject("ADODB.Command")

        ' 接続文字列（タイムアウト設定追加）
        conn.ConnectionTimeout = CONNECTION_TIMEOUT
        conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                  "Data Source=" & dbPath & ";"

        ' コマンドオブジェクトの設定
        Set cmd.ActiveConnection = conn
        cmd.CommandType = 1  ' 1 = adCmdText
        cmd.CommandTimeout = COMMAND_TIMEOUT

        ' テーブル存在確認
        tableExists = TableExistsInAccess(conn, "_ロット数量")

        If Not tableExists Then
            ' テーブルが存在しない場合は作成
            Application.StatusBar = yearKey & "年: 転送先テーブルを作成しています..."

            Dim sqlCreate As String
            sqlCreate = "CREATE TABLE [_ロット数量] (" & _
                        "[ID] AUTOINCREMENT PRIMARY KEY, " & _
                        "[日付] DATETIME, " & _
                        "[品番] TEXT(50), " & _
                        "[品番末尾] TEXT(10), " & _
                        "[注番月] TEXT(10), " & _
                        "[ロット] TEXT(50), " & _
                        "[工程] TEXT(50), " & _
                        "[ロット数量] DOUBLE);"

            conn.Execute sqlCreate
        End If

        ' アクセスのフィールド確認
        sqlCheck = "SELECT TOP 1 * FROM [_ロット数量]"
        Set rs = conn.Execute(sqlCheck)

        Dim accessFields As Object
        Set accessFields = CreateObject("Scripting.Dictionary")
        Dim f As Object

        For Each f In rs.Fields
            accessFields.Add f.Name, True
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
            Application.StatusBar = yearKey & "年: 以下のフィールドがAccess側テーブルに見つかりません: " & missingFields
            rs.Close
            conn.Close
            GoTo CleanExit
        End If

        rs.Close

        ' この年のデータの日付範囲を計算
        Dim minDate As Date
        Dim maxDate As Date
        Dim dateValue As Variant

        minDate = DateSerial(2100, 1, 1)
        maxDate = DateSerial(1900, 1, 1)

        Set rowNumbers = yearGroups(yearKey)
        For Each rowNum In rowNumbers
            dateValue = tbl.ListRows(CLng(rowNum)).Range(1, dateIndex).Value
            If IsDate(dateValue) Then
                If CDate(dateValue) < minDate Then minDate = CDate(dateValue)
                If CDate(dateValue) > maxDate Then maxDate = CDate(dateValue)
            End If
        Next rowNum

        ' 安全マージンを追加（±14日で月跨ぎの跨り生産に対応）
        minDate = minDate - 14
        maxDate = maxDate + 14

        ' 既存データの確認（重複転送防止・日付更新用）
        Application.StatusBar = yearKey & "年: 既存データを確認しています..."

        Set existingDict = CreateObject("Scripting.Dictionary")

        Dim dateFilter As String
        dateFilter = " WHERE [日付] BETWEEN #" & Format(minDate, "yyyy/mm/dd") & "# AND #" & Format(maxDate, "yyyy/mm/dd") & "#"

        ' SELECTに日付を追加（重複時の日付比較用）
        sqlCheck = "SELECT [日付], " & Replace(keyFields, ",", ", ") & " FROM [_ロット数量]" & dateFilter

        Set rs = conn.Execute(sqlCheck)

        recordCount = 0
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
                    ' 日付も保存（重複時の比較用）
                    ' 値をコピーして保存（Recordset参照のまま保存するとClose後にエラー3420）
                    Dim storedDate As Variant
                    If IsNull(rs("日付")) Then
                        storedDate = Null
                    Else
                        storedDate = CDate(rs("日付"))
                    End If
                    existingDict.Add key, storedDate
                End If

                rs.MoveNext
                recordCount = recordCount + 1
            Loop
        End If
        rs.Close

        ' トランザクション開始
        conn.BeginTrans
        transStarted = True

        ' この年のレコードを転送
        Application.StatusBar = yearKey & "年: データを転送しています..."
        successCount = 0
        updatedCount = 0
        skippedCount = 0
        duplicateSkippedCount = 0
        batchCounter = 0

        For Each rowNum In rowNumbers
            i = CLng(rowNum)

            ' 空白行チェック
            If IsRowEmpty(tbl, i, targetFields, fieldIndices) Then
                skippedCount = skippedCount + 1
                GoTo NextRow
            End If

            ' キー値を作成して重複チェック
            key = CreateKeyFromRow(tbl, i, keyFields, fieldIndices)

            If Not existingDict.Exists(key) Then
                ' 新規レコード：INSERT
                Dim sqlInsert As String
                sqlInsert = "INSERT INTO [_ロット数量] (" & fieldList & ") VALUES (" & _
                            GetSelectedValues(tbl, i, targetFields, fieldIndices) & ");"

                conn.Execute sqlInsert

                successCount = successCount + 1
                ' 新規追加したレコードの日付をDictionaryに保存
                existingDict.Add key, tbl.ListRows(i).Range(1, dateIndex).Value

                batchCounter = batchCounter + 1

                If batchCounter >= BATCH_SIZE Then
                    conn.CommitTrans
                    transStarted = False
                    conn.BeginTrans
                    transStarted = True
                    batchCounter = 0
                End If
            Else
                ' 重複レコード：日付が異なれば更新
                Dim existingDate As Variant
                Dim newDate As Variant

                existingDate = existingDict(key)
                newDate = tbl.ListRows(i).Range(1, dateIndex).Value

                If IsDate(newDate) And IsDate(existingDate) Then
                    If CLng(newDate) <> CLng(existingDate) Then
                        ' 日付が異なる場合はUPDATE（日付部分のみ比較するためCLngで整数化）
                        Dim sqlUpdate As String
                        sqlUpdate = BuildUpdateSQL(tbl, i, fieldIndices, keyFields)
                        conn.Execute sqlUpdate

                        ' Dictionaryの日付も更新
                        existingDict(key) = CDate(newDate)
                        updatedCount = updatedCount + 1

                        batchCounter = batchCounter + 1

                        If batchCounter >= BATCH_SIZE Then
                            conn.CommitTrans
                            transStarted = False
                            conn.BeginTrans
                            transStarted = True
                            batchCounter = 0
                        End If
                    Else
                        ' 日付も同一：重複スキップ
                        duplicateSkippedCount = duplicateSkippedCount + 1
                    End If
                End If
            End If

NextRow:
        Next rowNum

        ' 最後のバッチをコミット
        If transStarted Then
            conn.CommitTrans
            transStarted = False
        End If

        totalSuccess = totalSuccess + successCount
        totalUpdated = totalUpdated + updatedCount
        totalDuplicateSkipped = totalDuplicateSkipped + duplicateSkippedCount

        ' この年を処理済みリストに追加（月別分割マクロとの連携用）
        AddTransferredYear CInt(yearKey)

        ' 接続クローズ
        conn.Close
        Set conn = Nothing

        ' 年別ステータス表示
        Dim yearStatus As String
        yearStatus = yearKey & "年: " & successCount & "件転送、" & updatedCount & "件更新"
        If duplicateSkippedCount > 0 Then
            yearStatus = yearStatus & "、" & duplicateSkippedCount & "件スキップ（既存一致）"
        End If
        Application.StatusBar = yearStatus
        DoEvents

    Next yearKey

    ' ============================================
    ' Phase 3: 完了処理とクリア
    ' ============================================
    Dim elapsedTime As String
    elapsedTime = Format((Timer - startTime) / 86400, "hh:mm:ss")

    ' 完了メッセージ作成
    Dim finalStatus As String
    finalStatus = "全" & totalSuccess & "件転送、" & totalUpdated & "件更新"
    If totalDuplicateSkipped > 0 Then
        finalStatus = finalStatus & "、" & totalDuplicateSkipped & "件スキップ（既存一致）"
    End If
    finalStatus = finalStatus & "。処理時間: " & elapsedTime

    Application.StatusBar = finalStatus
    DoEvents
    Application.Wait Now + TimeValue("00:00:03")

    ' ソースデータを自動的にクリア（転送・更新・重複スキップのいずれかがあれば）
    If totalSuccess > 0 Or totalUpdated > 0 Or totalDuplicateSkipped > 0 Then
        Application.StatusBar = "ソースデータをクリアしています..."
        ClearSourceTable tbl, targetFields, fieldIndices
        Application.StatusBar = "ソースデータをクリアしました。" & finalStatus
        Application.Wait Now + TimeValue("00:00:02")
    End If

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
            If transStarted Then
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
    Set accessFields = Nothing
    Set years = Nothing
    Set yearGroups = Nothing

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Dim errMsg As String
    errMsg = "エラーが発生しました[" & errorLocation & "]: " & Err.Description & " (エラー番号: " & Err.Number & ")"

    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.state = 1 Then
            If transStarted Then
                conn.RollbackTrans
            End If
        End If
    End If

    Application.StatusBar = errMsg
    MsgBox errMsg, vbExclamation, "エラー - ロット数量データ転送ADO"
    Resume CleanExit
End Sub

' ============================================
' 補助関数: TableExistsInAccess
' ============================================
Function TableExistsInAccess(conn As Object, tableName As String) As Boolean
    Dim tempRS As Object

    On Error Resume Next
    Set tempRS = conn.Execute("SELECT TOP 1 * FROM [" & tableName & "]")

    If Err.Number = 0 Then
        TableExistsInAccess = True
    Else
        TableExistsInAccess = False
    End If

    If Not tempRS Is Nothing Then
        If tempRS.state = 1 Then tempRS.Close
        Set tempRS = Nothing
    End If

    Err.Clear
    On Error GoTo 0
End Function

' ============================================
' 補助関数: IsRowEmpty
' ============================================
Function IsRowEmpty(tbl As ListObject, rowIndex As Long, targetFields As Variant, fieldIndices As Object) As Boolean
    Dim ii As Integer
    Dim fName As String
    Dim colIndex As Integer
    Dim cellValue As Variant

    IsRowEmpty = True

    For ii = 0 To UBound(targetFields)
        fName = targetFields(ii)
        colIndex = fieldIndices(fName)
        cellValue = tbl.ListRows(rowIndex).Range(1, colIndex).Value

        If Not IsEmpty(cellValue) And Not IsNull(cellValue) Then
            If VarType(cellValue) = vbString Then
                If Len(Trim(cellValue)) > 0 Then
                    IsRowEmpty = False
                    Exit Function
                End If
            Else
                IsRowEmpty = False
                Exit Function
            End If
        End If
    Next ii
End Function

' ============================================
' 補助関数: CreateKeyFromRow
' ============================================
Function CreateKeyFromRow(tbl As ListObject, rowIndex As Long, keyFields As String, fieldIndices As Object) As String
    Dim keyStr As String
    Dim fArray As Variant
    Dim ii As Integer
    Dim fName As String
    Dim colIndex As Integer
    Dim cellValue As Variant

    keyStr = ""
    fArray = Split(keyFields, ",")

    For ii = 0 To UBound(fArray)
        fName = Trim(fArray(ii))

        If fieldIndices.Exists(fName) Then
            colIndex = fieldIndices(fName)
            cellValue = tbl.ListRows(rowIndex).Range(1, colIndex).Value

            If IsEmpty(cellValue) Or IsNull(cellValue) Then
                keyStr = keyStr & "NULL|"
            Else
                keyStr = keyStr & CStr(cellValue) & "|"
            End If
        Else
            keyStr = keyStr & "MISSING|"
        End If
    Next ii

    CreateKeyFromRow = keyStr
End Function

' ============================================
' 補助関数: GetSelectedValues
' ============================================
Function GetSelectedValues(tbl As ListObject, rowIndex As Long, targetFields As Variant, fieldIndices As Object) As String
    Dim result As String
    Dim ii As Integer
    Dim fName As String
    Dim colIndex As Integer
    Dim cellValue As Variant

    result = ""

    For ii = 0 To UBound(targetFields)
        If ii > 0 Then result = result & ", "

        fName = targetFields(ii)
        colIndex = fieldIndices(fName)
        cellValue = tbl.ListRows(rowIndex).Range(1, colIndex).Value

        If IsEmpty(cellValue) Or IsNull(cellValue) Then
            result = result & "NULL"
        ElseIf IsDate(cellValue) Then
            result = result & "#" & Format(cellValue, "yyyy/mm/dd") & "#"
        ElseIf IsNumeric(cellValue) Then
            result = result & cellValue
        Else
            result = result & "'" & Replace(cellValue, "'", "''") & "'"
        End If
    Next ii

    GetSelectedValues = result
End Function

' ============================================
' 補助Sub: ClearSourceTable
' ============================================
Sub ClearSourceTable(tbl As ListObject, targetFields As Variant, fieldIndices As Object)
    Dim ii As Integer
    Dim fName As String
    Dim colIndex As Integer

    If tbl.ListRows.Count = 0 Then Exit Sub

    For ii = 0 To UBound(targetFields)
        fName = targetFields(ii)

        If fieldIndices.Exists(fName) Then
            colIndex = fieldIndices(fName)
            tbl.ListColumns(colIndex).DataBodyRange.ClearContents
        End If
    Next ii
End Sub
