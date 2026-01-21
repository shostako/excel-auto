Attribute VB_Name = "mゾーン別データ転送ADO"
Option Explicit

' ========================================
' マクロ名: ゾーン別データ転送ADO
' 処理概要: Excelテーブル「_不良集計ゾーン別S」からAccessデータベースへADO経由でデータ転送
' ソーステーブル: シート「(Active)」テーブル「_不良集計ゾーン別S」
' ターゲットテーブル: Accessデータベース「不良調査表DB-{年}.accdb」テーブル「_不良集計ゾーン別」
' 作成日: 不明
' 更新日: 2026-01-21（品番末尾・注番月フィールドを削除）
'
' 処理の流れ:
' 1. Excelテーブルからフィールド構造とデータ範囲を取得
' 2. データから含まれる年を抽出し、全DBファイルの存在を事前確認
' 3. DBファイルが1つでも存在しない場合は処理中止（データは残る）
' 4. 年ごとにADO接続でAccessデータベースに接続
' 5. トランザクション制御でバッチINSERT実行（50件ごとコミット）
' 6. 全年の転送成功後、ソーステーブルの指定フィールドをクリア
'
' 技術的特徴:
' - 年別自動振り分け機能（日付から年を抽出し適切なDBに転送）
' - 事前DBファイル存在チェック（All or Nothing方式）
' - ADO接続とトランザクション管理
' - バッチ処理によるパフォーマンス最適化（BATCH_SIZE = 50）
' - 動的フィールドマッピング（ヘッダー行から自動取得）
' - 空白行スキップ機能
' - エラー位置特定機能（errorLocation変数）
' - SQLインジェクション対策（シングルクォートエスケープ）
' ========================================

' ============================================
' 定数定義：バッチサイズとタイムアウト設定
' ============================================
Const BATCH_SIZE As Long = 50        ' バッチ処理サイズ（この件数ごとにコミット）
Const CONNECTION_TIMEOUT As Long = 30 ' 接続タイムアウト(秒)
Const COMMAND_TIMEOUT As Long = 60   ' コマンドタイムアウト(秒)

' DBパス設定（年別振り分け用）
Const DB_BASE_PATH As String = "Z:¥全社共有¥オート事業部¥日報¥不良集計¥不良集計表¥"
Const DB_FILE_PREFIX As String = "不良調査表DB-"

' ============================================
' 補助関数: BuildDBPath
' 役割: 年からDBファイルパスを動的生成
' 引数: yearValue - 年（Integer）
' 戻り値: Z:¥...¥{年}年¥不良調査表DB-{年}.accdb
' ============================================
Function BuildDBPath(yearValue As Integer) As String
    BuildDBPath = DB_BASE_PATH & yearValue & "年¥" & DB_FILE_PREFIX & yearValue & ".accdb"
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

Sub ゾーン別データ転送ADO()
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
    Dim successCount As Long
    Dim skippedCount As Long ' 空白行のスキップカウント用
    Dim transStarted As Boolean ' トランザクション開始フラグ
    Dim batchCounter As Long   ' バッチ処理用カウンター
    Dim startTime As Double    ' 処理時間計測用
    Dim errorLocation As String ' エラー発生箇所特定用

    ' 年別振り分け用変数
    Dim years As Object
    Dim yearGroups As Object
    Dim yearKey As Variant
    Dim dbPath As String
    Dim missingDBs As String
    Dim totalSuccess As Long
    Dim rowNumbers As Collection
    Dim rowNum As Variant
    Dim dateIndex As Integer

    ' 処理時間計測開始
    startTime = Timer

    ' トランザクションフラグを初期化
    transStarted = False
    batchCounter = 0
    totalSuccess = 0

    ' 年リストを初期化（月別分割マクロとの連携用）
    InitTransferredYears

    ' 転送対象のフィールドを明示的に指定（品番末尾・注番月は2026-01-20に廃止）
    Dim targetFields As Variant
    targetFields = Array("日付", "品番", "ロット", "発見", "ゾーン", "番号", "数量", "差戻し")

    ' ============================================
    ' 初期設定：画面更新抑制とエラーハンドリング設定
    ' ============================================
    Application.ScreenUpdating = False
    Application.StatusBar = "ADO転送処理を開始します..."

    ' エラー処理
    On Error GoTo ErrorHandler

    ' 処理位置を記録
    errorLocation = "テーブル取得"

    ' ============================================
    ' データ検証：ソーステーブルの取得と存在確認、行数チェック
    ' ============================================
    Set tbl = ActiveSheet.ListObjects("_不良集計ゾーン別S")
    If tbl Is Nothing Then
        Application.StatusBar = "テーブル「_不良集計ゾーン別S」が見つかりません。"
        Application.Wait Now + TimeValue("00:00:03") ' 3秒間表示
        GoTo CleanExit
    End If

    rowCount = tbl.ListRows.Count
    If rowCount = 0 Then
        Application.StatusBar = "転送するデータがありません。"
        Application.Wait Now + TimeValue("00:00:03") ' 3秒間表示
        GoTo CleanExit
    End If

    ' 処理位置を記録
    errorLocation = "フィールドマッピング"

    ' ============================================
    ' フィールドマッピング：Excelヘッダー行から列インデックスを取得
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

        ' アクセスのフィールド確認
        sqlCheck = "SELECT TOP 1 * FROM [_不良集計ゾーン別]"
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

        ' トランザクション開始
        conn.BeginTrans
        transStarted = True

        ' この年のレコードを転送
        Application.StatusBar = yearKey & "年: データを転送しています..."
        successCount = 0
        skippedCount = 0
        batchCounter = 0

        Set rowNumbers = yearGroups(yearKey)
        For Each rowNum In rowNumbers
            i = CLng(rowNum)

            ' 空白行チェック
            If IsRowEmpty(tbl, i, targetFields, fieldIndices) Then
                skippedCount = skippedCount + 1
                GoTo NextRow
            End If

            ' 無条件でINSERT（重複チェックなし）
            Dim sqlInsert As String
            sqlInsert = "INSERT INTO [_不良集計ゾーン別] (" & fieldList & ") VALUES (" & _
                        GetSelectedValues(tbl, i, targetFields, fieldIndices) & ");"

            conn.Execute sqlInsert

            successCount = successCount + 1
            batchCounter = batchCounter + 1

            If batchCounter >= BATCH_SIZE Then
                conn.CommitTrans
                transStarted = False
                conn.BeginTrans
                transStarted = True
                batchCounter = 0
            End If

NextRow:
        Next rowNum

        ' 最後のバッチをコミット
        If transStarted Then
            conn.CommitTrans
            transStarted = False
        End If

        totalSuccess = totalSuccess + successCount

        ' この年を処理済みリストに追加（月別分割マクロとの連携用）
        AddTransferredYear CInt(yearKey)

        ' 接続クローズ
        conn.Close
        Set conn = Nothing

        Application.StatusBar = yearKey & "年: " & successCount & "件転送完了"
        DoEvents

    Next yearKey

    ' ============================================
    ' Phase 3: 完了処理とクリア
    ' ============================================
    Dim elapsedTime As String
    elapsedTime = Format((Timer - startTime) / 86400, "hh:mm:ss")

    If totalSuccess > 0 Then
        Application.StatusBar = "全" & totalSuccess & "件のデータを転送しました。処理時間: " & elapsedTime
    Else
        Application.StatusBar = "新規データはありませんでした。処理時間: " & elapsedTime
    End If
    DoEvents
    Application.Wait Now + TimeValue("00:00:03")

    ' ソースデータを自動的にクリア（データの有無に関わらず実行）
    Application.StatusBar = "ソースデータをクリアしています..."
    ClearSourceTable tbl, targetFields, fieldIndices
    Application.StatusBar = "ソースデータをクリアしました。処理が完了しました。処理時間: " & elapsedTime
    Application.Wait Now + TimeValue("00:00:02")

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
    MsgBox errMsg, vbExclamation, "エラー - ゾーン別データ転送ADO"

    Resume CleanExit
End Sub

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

    If tbl Is Nothing Then Exit Sub
    If tbl.ListRows.Count = 0 Then Exit Sub

    On Error Resume Next

    For ii = 0 To UBound(targetFields)
        fName = targetFields(ii)

        If fieldIndices.Exists(fName) Then
            colIndex = fieldIndices(fName)
            If Not tbl.ListColumns(colIndex).DataBodyRange Is Nothing Then
                tbl.ListColumns(colIndex).DataBodyRange.ClearContents
                DoEvents
            End If
        End If
    Next ii

    On Error GoTo 0
End Sub
