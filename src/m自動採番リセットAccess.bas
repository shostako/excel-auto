Attribute VB_Name = "m自動採番リセットAccess"
Option Explicit

' ============================================
' マクロ名: リセット自動採番ID順序保持版
' 処理概要: Accessテーブルの自動採番IDを1から再開（元の順序を保持）
'
' 処理内容:
'   1. ユーザーに年とテーブル名入力を求める
'   2. 元のテーブル構造とデータを解析
'   3. バックアップテーブルと順序保持用テーブルを作成
'   4. 元のテーブルを削除して新規作成（AUTOINCREMENT付き）
'   5. 元の順序を保持したままデータを復元
'
' 接続先データベース:
'   動的構築（年を入力で指定）
'
' 警告:
'   - この処理はテーブルのリレーションシップを破壊する可能性があります
'   - 実行前に必ずデータベースのバックアップを取得してください
'   - 処理後、リレーションシップは手動で再設定が必要です
'
' バックアップ:
'   - 元テーブル: [テーブル名_バックアップ_yyyymmdd_hhnnss]
'   - 順序保持用: [テーブル名_順序_yyyymmdd_hhnnss]（処理後削除）
'
' 更新日: 2026-01-21（DBパス動的構築対応）
' ============================================

' DBパス設定（他マクロと統一）
Private Const DB_BASE_PATH As String = "Z:\全社共有\オート事業部\日報\不良集計\不良集計表\"
Private Const DB_FILE_PREFIX As String = "不良調査表DB-"

Sub リセット自動採番ID順序保持版()
    ' 変数宣言
    Dim conn As Object
    Dim rs As Object
    Dim fieldObj As Object
    Dim sqlSelect As String, sqlCreate As String, sqlInsert As String, sqlDrop As String
    Dim sqlExport As String
    Dim targetTable As String
    Dim backupTable As String
    Dim yearValue As Integer
    Dim dbPath As String
    Dim yearInput As String
    Dim result As VbMsgBoxResult
    Dim idColumnName As String
    Dim fieldList As String, createFieldList As String
    Dim i As Integer
    Dim fieldCount As Integer
    Dim fieldNames() As String
    Dim fieldTypes() As String
    Dim fieldSizes() As Long
    Dim fieldAttributes() As Boolean
    Dim fieldDecimals() As Integer
    Dim hasRelationships As Boolean
    Dim recordCount As Long
    Dim newCount As Long
    Dim tempOrderTable As String ' 順序保持用の一時テーブル
    
    ' エラー処理
    On Error GoTo ErrorHandler
    
    ' 画面更新の停止
    Application.ScreenUpdating = False
    Application.StatusBar = "処理を開始します..."

    ' ============================================
    ' 実行確認とテーブル名入力
    ' リレーションシップ破壊のリスクをユーザーに警告
    ' ============================================
    result = MsgBox("Accessテーブルの自動採番IDをリセットします。" & vbCrLf & _
                    "この処理はテーブル内のデータの順序を保持したまま、ID番号を1から振り直します。" & vbCrLf & _
                    "実行する前にデータベースのバックアップを取ることを強くお勧めします。" & vbCrLf & vbCrLf & _
                    "注意: この処理を実行すると、このテーブルと関連するリレーションシップが破壊される可能性があります。" & vbCrLf & vbCrLf & _
                    "続行しますか？", vbQuestion + vbYesNo, "自動採番リセット確認")
    
    If result <> vbYes Then
        MsgBox "処理をキャンセルしました。", vbInformation
        GoTo CleanExit
    End If
    
    ' 年の指定
    yearInput = InputBox("対象年を入力してください（例: 2026）:", "年の指定", Year(Date))

    If yearInput = "" Then
        MsgBox "処理をキャンセルしました。", vbInformation
        GoTo CleanExit
    End If

    If Not IsNumeric(yearInput) Then
        MsgBox "年は数値で入力してください。", vbExclamation
        GoTo CleanExit
    End If

    yearValue = CInt(yearInput)

    ' DBパス構築
    dbPath = DB_BASE_PATH & yearValue & "年\" & DB_FILE_PREFIX & yearValue & ".accdb"

    ' ファイル存在チェック
    If Dir(dbPath) = "" Then
        MsgBox yearValue & "年のデータベースが見つかりません。" & vbCrLf & dbPath, vbExclamation
        GoTo CleanExit
    End If

    ' テーブル名の指定
    targetTable = InputBox("リセット対象のテーブル名を入力してください:", "テーブル名", "_不良集計ゾーン別")

    If targetTable = "" Then
        MsgBox "処理をキャンセルしました。", vbInformation
        GoTo CleanExit
    End If

    ' バックアップテーブル名と順序保持用テーブル名の設定
    backupTable = targetTable & "_バックアップ_" & Format(Now, "yyyymmdd_hhnnss")
    tempOrderTable = targetTable & "_順序_" & Format(Now, "yyyymmdd_hhnnss")
    
    ' ADO接続
    Set conn = CreateObject("ADODB.Connection")
    
    ' 接続文字列（動的パス使用）
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    
    ' ターゲットテーブルが存在するか確認
    If Not TableExistsInAccess(conn, targetTable) Then
        MsgBox "指定されたテーブル「" & targetTable & "」が見つかりません。", vbExclamation
        GoTo CleanExit
    End If
    
    ' リレーションシップの警告
    hasRelationships = True ' 常にTrueと仮定して警告
    
    If hasRelationships Then
        result = MsgBox("このテーブルにはリレーションシップが設定されている可能性があります。" & vbCrLf & _
                      "処理を続行すると、リレーションシップが破壊される可能性があります。" & vbCrLf & vbCrLf & _
                      "続行しますか？", vbExclamation + vbYesNo, "リレーションシップ警告")
                      
        If result <> vbYes Then
            MsgBox "処理をキャンセルしました。", vbInformation
            GoTo CleanExit
        End If
    End If

    ' ============================================
    ' テーブル構造の解析
    ' フィールド名、データ型、サイズ、NULL許容を取得
    ' ============================================
    Application.StatusBar = "テーブル構造を取得しています..."
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM [" & targetTable & "] WHERE 1=0", conn, 1, 3 ' adOpenKeyset, adLockOptimistic
    
    ' フィールド情報を格納する配列の準備
    fieldCount = rs.Fields.Count
    ReDim fieldNames(fieldCount - 1)
    ReDim fieldTypes(fieldCount - 1)
    ReDim fieldSizes(fieldCount - 1)
    ReDim fieldAttributes(fieldCount - 1)
    ReDim fieldDecimals(fieldCount - 1)
    
    ' テーブル構造を解析してフィールド情報を取得
    For i = 0 To fieldCount - 1
        Set fieldObj = rs.Fields(i)
        fieldNames(i) = fieldObj.Name
        fieldTypes(i) = GetAccessDataType(fieldObj.Type)
        fieldSizes(i) = IIf(fieldObj.DefinedSize > 0, fieldObj.DefinedSize, 0)
        fieldAttributes(i) = fieldObj.Attributes And 1 ' adFldMayBeNull = 1
        fieldDecimals(i) = 0 ' ADOでは小数点以下の桁数が取得できないため0に設定
        
        ' ID列名を保存（最初のフィールドと仮定）
        If i = 0 Then
            idColumnName = fieldObj.Name
        End If
    Next i
    
    rs.Close

    ' ============================================
    ' バックアップと順序保持用テーブル作成
    ' ROW_NUMBER() OVER構文またはレコードセットで順序を保存
    ' ============================================
    Application.StatusBar = "データをバックアップしています..."
    sqlExport = "SELECT * INTO [" & backupTable & "] FROM [" & targetTable & "]"
    conn.Execute sqlExport
    
    ' 元データに連番フィールドを追加した順序保持用テーブルを作成
    Application.StatusBar = "順序保持用テーブルを作成しています..."
    sqlCreate = "SELECT *, ROW_NUMBER() OVER (ORDER BY [" & idColumnName & "]) AS OriginalOrder " & _
                "INTO [" & tempOrderTable & "] FROM [" & targetTable & "]"
    
    On Error Resume Next
    conn.Execute sqlCreate
    
    If Err.Number <> 0 Then
        ' ROW_NUMBER OVER構文がサポートされていない古いAccessバージョンの場合の代替処理
        Err.Clear
        On Error GoTo ErrorHandler
        
        ' 代替方法：順序列を持つテーブルを作成（二段階プロセス）
        ' 1. 一時テーブル作成
        sqlCreate = "SELECT * INTO [" & tempOrderTable & "] FROM [" & targetTable & "]"
        conn.Execute sqlCreate
        
        ' 2. 順序フィールドを追加
        conn.Execute "ALTER TABLE [" & tempOrderTable & "] ADD COLUMN OriginalOrder LONG"
        
        ' 3. レコードセットを使って順序を設定
        Dim updateRS As Object
        Set updateRS = CreateObject("ADODB.Recordset")
        updateRS.Open "SELECT * FROM [" & tempOrderTable & "] ORDER BY [" & idColumnName & "]", _
                    conn, 1, 3 ' adOpenKeyset, adLockOptimistic
        
        Dim orderNum As Long
        orderNum = 1
        
        If Not updateRS.EOF Then
            updateRS.MoveFirst
            Do Until updateRS.EOF
                updateRS("OriginalOrder") = orderNum
                orderNum = orderNum + 1
                updateRS.Update
                updateRS.MoveNext
            Loop
        End If
        
        updateRS.Close
        Set updateRS = Nothing
    Else
        On Error GoTo ErrorHandler
    End If
    
    ' テーブルをバックアップできたか確認
    If Not TableExistsInAccess(conn, tempOrderTable) Then
        MsgBox "順序保持用テーブルの作成に失敗しました。処理を中止します。", vbExclamation
        GoTo CleanExit
    End If
    
    ' レコード数を確認
    Set rs = conn.Execute("SELECT COUNT(*) FROM [" & tempOrderTable & "]")
    recordCount = rs(0)
    rs.Close
    
    If recordCount = 0 Then
        result = MsgBox("対象テーブルにはデータがありません。処理を続行しますか？", _
                       vbQuestion + vbYesNo, "確認")
        If result <> vbYes Then
            MsgBox "処理をキャンセルしました。", vbInformation
            GoTo CleanExit
        End If
    End If
    
    ' テーブル作成用のフィールドリスト作成
    createFieldList = ""
    
    ' ID列を含むすべてのフィールドの定義を構築
    For i = 0 To fieldCount - 1
        If i > 0 Then createFieldList = createFieldList & ", "
        
        ' ID列（最初のフィールド）は特別な処理
        If i = 0 Then
            createFieldList = createFieldList & "[" & fieldNames(i) & "] AUTOINCREMENT PRIMARY KEY"
        Else
            createFieldList = createFieldList & "[" & fieldNames(i) & "] " & fieldTypes(i)
            
            ' テキストフィールドの場合はサイズを指定
            If fieldTypes(i) = "TEXT" And fieldSizes(i) > 0 Then
                createFieldList = createFieldList & "(" & fieldSizes(i) & ")"
            End If
            
            ' NULL許容を設定
            If fieldAttributes(i) Then
                createFieldList = createFieldList & " NULL"
            End If
        End If
    Next i
    
    ' ID以外のフィールドリストを作成（データ復元用）
    fieldList = ""
    For i = 1 To fieldCount - 1 ' i=1から開始して最初のフィールド（ID）を除外
        If fieldList <> "" Then fieldList = fieldList & ", "
        fieldList = fieldList & "[" & fieldNames(i) & "]"
    Next i

    ' ============================================
    ' テーブル再作成とデータ復元
    ' 元のテーブルを削除し、AUTOINCREMENT付きで再作成
    ' 元の順序を保持したままデータを復元
    ' ============================================
    Application.StatusBar = "元のテーブルを削除しています..."
    sqlDrop = "DROP TABLE [" & targetTable & "]"
    conn.Execute sqlDrop
    
    ' 新しいテーブルを作成
    Application.StatusBar = "新しいテーブルを作成しています..."
    sqlCreate = "CREATE TABLE [" & targetTable & "] (" & createFieldList & ")"
    conn.Execute sqlCreate
    
    ' 作成できたか確認
    If Not TableExistsInAccess(conn, targetTable) Then
        MsgBox "新しいテーブルの作成に失敗しました。バックアップは「" & backupTable & "」にあります。", vbExclamation
        GoTo CleanExit
    End If
    
    ' データを元の順序で復元
    If recordCount > 0 Then
        Application.StatusBar = "データを元の順序で復元しています..."
        sqlInsert = "INSERT INTO [" & targetTable & "] (" & fieldList & ") " & _
                    "SELECT " & fieldList & " FROM [" & tempOrderTable & "] ORDER BY OriginalOrder"
        
        conn.Execute sqlInsert
        
        ' 復元できたか確認
        Set rs = conn.Execute("SELECT COUNT(*) FROM [" & targetTable & "]")
        newCount = rs(0)
        rs.Close
        
        If newCount <> recordCount Then
            MsgBox "データの復元に問題が発生しました。" & vbCrLf & _
                  "元のレコード数: " & recordCount & vbCrLf & _
                  "復元されたレコード数: " & newCount & vbCrLf & vbCrLf & _
                  "バックアップは「" & backupTable & "」と「" & tempOrderTable & "」にあります。", vbExclamation
            GoTo CleanExit
        End If
    End If
    
    ' 順序保持用テーブルの削除（オプション）
    Application.StatusBar = "一時テーブルをクリーンアップしています..."
    sqlDrop = "DROP TABLE [" & tempOrderTable & "]"
    On Error Resume Next
    conn.Execute sqlDrop
    On Error GoTo ErrorHandler
    
    ' 成功メッセージ
    MsgBox "テーブル「" & targetTable & "」の自動採番IDを正常にリセットしました。" & vbCrLf & _
           "レコード数: " & recordCount & vbCrLf & vbCrLf & _
           "元のデータ順序を保持したまま、ID番号が1から振り直されました。" & vbCrLf & vbCrLf & _
           "バックアップテーブル「" & backupTable & "」は保持されています。" & vbCrLf & _
           "問題がなければ後で手動で削除してください。" & vbCrLf & vbCrLf & _
           "注意: リレーションシップは自動的に復元されませんので、" & vbCrLf & _
           "必要に応じて手動で再設定してください。", vbInformation, "処理完了"
    
CleanExit:
    ' リソース解放
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.state > 0 Then rs.Close
    End If
    
    If Not conn Is Nothing Then
        If conn.state > 0 Then conn.Close
    End If
    
    Set fieldObj = Nothing
    Set rs = Nothing
    Set conn = Nothing
    
    ' 状態表示をリセット
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    Dim errMsg As String
    errMsg = "エラーが発生しました:" & vbCrLf & _
             "説明: " & Err.Description & vbCrLf & _
             "エラー番号: " & Err.Number
    
    ' 途中でエラーが発生した場合、バックアップの有無を確認
    If TableExistsInAccess(conn, backupTable) Then
        errMsg = errMsg & vbCrLf & vbCrLf & _
                "バックアップは「" & backupTable & "」に保存されています。"
    End If
    
    MsgBox errMsg, vbCritical, "自動採番リセットエラー"
    Resume CleanExit
End Sub

' ============================================
' 関数名: GetAccessDataType
' 処理概要: ADOデータ型をAccessデータ型文字列に変換
'
' 引数:
'   adoType - ADOデータ型コード（Integer）
'
' 戻り値:
'   Accessデータ型文字列（例: "TEXT", "LONG", "DATETIME"）
'
' 参照:
'   https://learn.microsoft.com/en-us/sql/ado/reference/ado-api/datatypeenum
' ============================================
Function GetAccessDataType(adoType As Integer) As String
    Select Case adoType
        Case 2  ' adSmallInt
            GetAccessDataType = "SHORT"
        Case 3  ' adInteger
            GetAccessDataType = "LONG"
        Case 4  ' adSingle
            GetAccessDataType = "SINGLE"
        Case 5  ' adDouble
            GetAccessDataType = "DOUBLE"
        Case 6  ' adCurrency
            GetAccessDataType = "CURRENCY"
        Case 7  ' adDate
            GetAccessDataType = "DATETIME"
        Case 11 ' adBoolean
            GetAccessDataType = "YESNO"
        Case 17 ' adUnsignedTinyInt
            GetAccessDataType = "BYTE"
        Case 20 ' adBigInt
            GetAccessDataType = "LONG"
        Case 72 ' adGUID
            GetAccessDataType = "GUID"
        Case 128, 200, 202, 203 ' adBinary, adVarBinary, adVarChar, adLongVarChar
            GetAccessDataType = "TEXT"
        Case 129, 130, 201 ' adChar, adWChar, adLongVarWChar
            GetAccessDataType = "TEXT"
        Case 131 ' adNumeric
            GetAccessDataType = "DECIMAL"
        Case 135 ' adDBTimeStamp
            GetAccessDataType = "DATETIME"
        Case Else
            ' デフォルトはテキスト
            GetAccessDataType = "TEXT"
    End Select
End Function

' ============================================
' 関数名: TableExistsInAccess
' 処理概要: Accessデータベースに指定テーブルが存在するか確認
'
' 引数:
'   conn - ADODB.Connectionオブジェクト
'   tableName - 確認対象のテーブル名
'
' 戻り値:
'   True: テーブルが存在する
'   False: テーブルが存在しない
'
' 処理方法:
'   SELECT TOP 1クエリを実行してエラー有無で判定
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
        If tempRS.state > 0 Then tempRS.Close
    End If
    
    Set tempRS = Nothing
    Err.Clear
    On Error GoTo 0
End Function
