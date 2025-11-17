Attribute VB_Name = "mゾーン別データ転送ADO"
Option Explicit

' 定数定義
Const BATCH_SIZE As Long = 50        ' バッチ処理サイズ
Const CONNECTION_TIMEOUT As Long = 30 ' 接続タイムアウト(秒)
Const COMMAND_TIMEOUT As Long = 60   ' コマンドタイムアウト(秒)
Const RECENT_DAYS As Long = 90       ' 直近何日分のデータを対象とするか（現在不使用。将来使うかも？）

Sub ゾーン別データ転送ADO()
    ' 変数宣言
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
    Dim skippedCount As Long ' 空白行のスキップカウント用
    Dim keyFields As String
    Dim transStarted As Boolean ' トランザクション開始フラグ
    Dim batchCounter As Long   ' バッチ処理用カウンター
    Dim startTime As Double    ' 処理時間計測用
    Dim recordCount As Long    ' レコード数カウント用
    Dim errorLocation As String ' エラー発生箇所特定用
    
    ' 処理時間計測開始
    startTime = Timer
    
    ' トランザクション開始フラグを初期化
    transStarted = False
    batchCounter = 0
    
    ' 転送対象のフィールドを明示的に指定（差戻しを末尾に追加）
    Dim targetFields As Variant
    targetFields = Array("日付", "品番", "品番末尾", "注番月", "ロット", "発見", "ゾーン", "番号", "数量", "差戻し")
    
    ' 進捗表示用
    Application.ScreenUpdating = False
    Application.StatusBar = "ADO転送処理を開始します..."
    
    ' エラー処理
    On Error GoTo ErrorHandler
    
    ' 処理位置を記録
    errorLocation = "テーブル取得"
    
    ' テーブル取得
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
        ' 新規データがなくても処理を続行（クリアするため）
        GoTo SkipTransfer
    End If
    
    ' 重複チェック用のキーフィールド設定
    keyFields = "日付,品番,品番末尾,注番月,ロット,発見,ゾーン,番号,差戻し"
    
    ' Dictionary オブジェクト作成
    Set existingDict = CreateObject("Scripting.Dictionary")
    
    ' 処理位置を記録
    errorLocation = "ADO接続"
    
    ' ADO接続
    Set conn = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")
    
    ' 接続文字列（タイムアウト設定追加）
    conn.ConnectionTimeout = CONNECTION_TIMEOUT
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
              "Data Source=Z:¥全社共有¥オート事業部¥日報¥不良集計¥不良集計表¥2025年¥不良調査表DB-2025.accdb;"
    
    ' コマンドオブジェクトの設定
    Set cmd.ActiveConnection = conn
    cmd.CommandType = 1  ' 1 = adCmdText
    cmd.CommandTimeout = COMMAND_TIMEOUT
    
    ' 処理位置を記録
    errorLocation = "フィールドマッピング"
    
    ' フィールド名と列インデックスの対応を取得
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
    
    ' 転送先テーブルのフィールド確認
    Application.StatusBar = "転送先テーブルの構造を確認しています..."
    
    ' 処理位置を記録
    errorLocation = "テーブル構造確認"
    
    ' テスト的にフィールドを取得してみる
    sqlCheck = "SELECT TOP 1 * FROM [_不良集計ゾーン別]"
    Set rs = conn.Execute(sqlCheck)
    
    ' アクセスのフィールド一覧
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
        Application.StatusBar = "以下のフィールドがAccess側テーブルに見つかりません: " & missingFields
        Application.Wait Now + TimeValue("00:00:05") ' 5秒間表示
        GoTo CleanExit
    End If
    
    rs.Close
    
    ' SQL用のフィールドリスト作成
    Dim fieldList As String
    fieldList = ""
    For Each fieldName In targetFields
        If fieldList <> "" Then fieldList = fieldList & ", "
        fieldList = fieldList & "[" & fieldName & "]"
    Next
    
    ' 処理位置を記録
    errorLocation = "既存データ確認"
    
    ' 既存データの確認（重複転送防止）- 日付による絞り込み追加
    Application.StatusBar = "既存データを確認しています...(0件)"
    
    ' 日付範囲を使って既存データを絞り込む
    ' Excel側の日付データの最小値と最大値を取得
    Dim minDate As Date
    Dim maxDate As Date
    Dim dateIndex As Integer
    Dim dateValue As Variant
    
    minDate = DateSerial(2100, 1, 1) ' 十分未来の日付
    maxDate = DateSerial(1900, 1, 1) ' 十分過去の日付
    
    dateIndex = fieldIndices("日付")
    
    ' データ有無チェック
    Dim hasData As Boolean
    hasData = False
    
    ' 日付範囲を計算
    For i = 1 To rowCount
        dateValue = tbl.ListRows(i).Range(1, dateIndex).Value
        If IsDate(dateValue) Then
            hasData = True
            If CDate(dateValue) < minDate Then minDate = CDate(dateValue)
            If CDate(dateValue) > maxDate Then maxDate = CDate(dateValue)
        End If
    Next i
    
    ' データがない場合は処理を終了（ただしクリアは実行）
    If Not hasData Then
        Application.StatusBar = "転送するデータに有効な日付がありません。"
        Application.Wait Now + TimeValue("00:00:03") ' 3秒間表示
        GoTo SkipTransfer ' クリア処理のために変更
    End If
    
    ' 安全マージンを追加（日付を±7日で広げる）
    minDate = minDate - 7
    maxDate = maxDate + 7
    
    ' 日付範囲を文字列に変換
    Dim dateFilter As String
    dateFilter = " WHERE [日付] BETWEEN #" & Format(minDate, "yyyy/mm/dd") & "# AND #" & Format(maxDate, "yyyy/mm/dd") & "#"
    
    sqlCheck = "SELECT " & Replace(keyFields, ",", ", ") & " FROM [_不良集計ゾーン別]" & dateFilter
    
    ' 処理時間が長い可能性があるため、進捗表示を更新
    Application.StatusBar = "既存データを確認しています...クエリを実行中"
    DoEvents ' UIの更新を許可
    
    Set rs = conn.Execute(sqlCheck)
    
    ' 既存データをDictionaryに格納
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
                existingDict.Add key, True
            End If
            
            rs.MoveNext
            
            ' 進捗表示の更新（100レコードごと）
            recordCount = recordCount + 1
            If recordCount Mod 100 = 0 Then
                Application.StatusBar = "既存データを確認しています...(" & recordCount & "件)"
                DoEvents ' UIの更新を許可
            End If
        Loop
    End If
    rs.Close
    
    ' 最終的な既存データ数を表示
    Application.StatusBar = "既存データを確認しました（" & recordCount & "件）。データ転送を開始します..."
    DoEvents ' UIを更新
    
    ' 処理位置を記録
    errorLocation = "データ転送"
    
    ' トランザクション開始
    conn.BeginTrans
    transStarted = True ' トランザクション開始フラグを設定
    
    ' データ転送実行
    Application.StatusBar = "データを転送しています (0/" & rowCount & " - 0%)"
    successCount = 0
    skippedCount = 0  ' スキップ行カウント初期化
    batchCounter = 0  ' バッチカウンター初期化
    
    For i = 1 To rowCount
        ' 空白行チェック
        If IsRowEmpty(tbl, i, targetFields, fieldIndices) Then
            skippedCount = skippedCount + 1
            ' 空白行はスキップ
            GoTo NextRow
        End If
    
        ' キー値を作成して重複チェック
        key = CreateKeyFromRow(tbl, i, keyFields, fieldIndices)

        ' 重複チェック（下記の2行のどちらかをコメントアウトして切り替え）
        If Not existingDict.Exists(key) Then     ' ← 重複チェック有効（デフォルト）
        'If True Then                             ' ← 重複チェック無効化（上をコメントアウトしてこちらを有効化）
            ' この行のINSERT SQL文を作成（指定フィールドのみ）
            Dim sqlInsert As String
            sqlInsert = "INSERT INTO [_不良集計ゾーン別] (" & fieldList & ") VALUES (" & _
                        GetSelectedValues(tbl, i, targetFields, fieldIndices) & ");"
            
            ' SQL実行
            conn.Execute sqlInsert
            
            ' 成功カウント増加
            successCount = successCount + 1

            ' Dictionaryに追加（重複チェック無効時のエラーを回避）
            On Error Resume Next
            existingDict.Add key, True
            On Error GoTo ErrorHandler

            ' バッチカウンターを増加
            batchCounter = batchCounter + 1
            
            ' バッチサイズに達したらコミットして新しいトランザクションを開始
            If batchCounter >= BATCH_SIZE Then
                conn.CommitTrans
                transStarted = False
                
                ' 進捗表示を更新
                Application.StatusBar = "データ転送中 - バッチ完了: " & i & "/" & rowCount & _
                                       " (" & Format(i / rowCount, "0%") & ") - 成功: " & successCount & "件"
                DoEvents ' UIの更新を許可
                
                ' 新しいトランザクションを開始
                conn.BeginTrans
                transStarted = True
                batchCounter = 0
            End If
        End If
        
NextRow:
        ' 進捗表示を更新（5行ごと、またはバッチの最後）
        If i Mod 5 = 0 Or i = rowCount Then
            Application.StatusBar = "データを転送しています (" & i & "/" & rowCount & " - " & _
                                    Format(i / rowCount, "0%") & ") - 成功: " & successCount & "件"
            DoEvents ' UIの更新を許可
        End If
    Next i
    
    ' 最後のバッチをコミット
    If transStarted Then
        conn.CommitTrans
        transStarted = False
    End If
    
SkipTransfer:
    ' 経過時間を計算
    Dim elapsedTime As String
    elapsedTime = Format((Timer - startTime) / 86400, "hh:mm:ss")
    
    ' 完了メッセージをステータスバーに表示
    If successCount > 0 Then
        Application.StatusBar = successCount & "件のデータを転送しました。" & skippedCount & "件の空白行をスキップしました。処理時間: " & elapsedTime
    Else
        Application.StatusBar = "新規データはありませんでした。" & skippedCount & "件の空白行をスキップしました。処理時間: " & elapsedTime
    End If
    DoEvents ' UIの更新を許可
    Application.Wait Now + TimeValue("00:00:03") ' 3秒間表示
    
    ' ソースデータを自動的にクリア（データの有無に関わらず実行）
    Application.StatusBar = "ソースデータをクリアしています..."
    ClearSourceTable tbl, targetFields, fieldIndices
    Application.StatusBar = "ソースデータをクリアしました。処理が完了しました。処理時間: " & elapsedTime
    Application.Wait Now + TimeValue("00:00:02") ' 2秒間表示
    
CleanExit:
    ' リソース解放
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.state = 1 Then rs.Close
    End If
    
    If Not conn Is Nothing Then
        If conn.state = 1 Then
            If transStarted Then ' トランザクションが開始されている場合のみロールバック
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
    
    ' 状態表示を戻す
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    Dim errMsg As String
    errMsg = "エラーが発生しました[" & errorLocation & "]: " & Err.Description & " (エラー番号: " & Err.Number & ")"
    
    ' トランザクションのロールバック
    On Error Resume Next
    If Not conn Is Nothing Then
        If conn.state = 1 Then
            If transStarted Then ' トランザクションが開始されている場合のみロールバック
                conn.RollbackTrans
            End If
        End If
    End If
    
    ' エラーメッセージをステータスバーに表示
    Application.StatusBar = errMsg
    MsgBox errMsg, vbExclamation, "エラー - ゾーン別データ転送ADO"
    
    ' エラー時はクリアを実行しない（元の処理どおり）
    Resume CleanExit
End Sub

' 行が空かどうかをチェックする関数
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

' テーブル行からキー値を作成する関数（列インデックス使用）
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

' 指定されたフィールドの値のみを取得する関数
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

' 指定フィールドのみをクリアする関数（より堅牢に修正）
Sub ClearSourceTable(tbl As ListObject, targetFields As Variant, fieldIndices As Object)
    Dim i As Integer
    Dim fieldName As String
    Dim colIndex As Integer
    
    ' テーブル有無と行数チェック
    If tbl Is Nothing Then Exit Sub
    If tbl.ListRows.Count = 0 Then Exit Sub
    
    ' 指定されたフィールドのみをクリア
    On Error Resume Next ' フィールドが見つからない場合のエラーを無視
    
    For i = 0 To UBound(targetFields)
        fieldName = targetFields(i)
        
        If fieldIndices.Exists(fieldName) Then
            colIndex = fieldIndices(fieldName)
            ' DataBodyRangeが存在することを確認
            If Not tbl.ListColumns(colIndex).DataBodyRange Is Nothing Then
                tbl.ListColumns(colIndex).DataBodyRange.ClearContents
                DoEvents ' UIの更新を許可
            End If
        End If
    Next i
    
    On Error GoTo 0 ' エラーハンドリングを元に戻す
End Sub

