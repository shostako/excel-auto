Attribute VB_Name = "mクエリ参照元変更"
' ========================================
' マクロ名: クエリ参照元変更
' 処理概要: ロット数量ADO、不良集計ゾーン別ADO、番号ADOの3つのクエリを複数月対応で更新
' 入力: 組合せシートのQ2（例：「2025-09」）
' 出力: 組合せシートのQ1に実行後の参照ファイル名
' 処理内容:
'   1. 組合せシートのQ2から年月を取得
'   2. 前月を自動計算
'   3. ロット数量ADOを前月・当月結合に更新
'   4. 不良集計ゾーン別ADOを前月・当月結合に更新
'   5. 番号ADOを固定参照（2025-09）に更新
'   6. Q1に結果メッセージを出力
' ========================================

Sub クエリ参照元変更()

    Application.StatusBar = "クエリ参照元の変更を開始します..."

    On Error GoTo ErrorHandler

    ' ============================================
    ' 変数宣言
    ' ============================================
    Dim ws As Worksheet
    Dim currentMonth As String
    Dim previousMonth As String
    Dim currentYear As String
    Dim currentFolder As String
    Dim previousFolder As String
    Dim qry As WorkbookQuery
    Dim updateCount As Integer
    Dim i As Integer
    Dim tempFormula As String
    Dim newFormula As String
    Dim resultMessage As String

    ' ============================================
    ' 組合せシートの確認と入力値取得
    ' ============================================
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("組合せ")
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        MsgBox "「組合せ」シートが見つかりません。", vbExclamation
        GoTo CleanExit
    End If

    ' Q2の値を取得
    currentMonth = Trim(ws.Range("Q2").Value)

    ' 入力値チェック
    If currentMonth = "" Then
        MsgBox "Q2の値が設定されていません。" & vbCrLf & _
               "yyyy-mm形式で年月を入力してください。", vbExclamation
        GoTo CleanExit
    End If

    Debug.Print "Q2の値: " & currentMonth

    ' ============================================
    ' 前月計算
    ' ============================================
    Dim yearPart As Integer
    Dim monthPart As Integer

    If Len(currentMonth) >= 7 And Mid(currentMonth, 5, 1) = "-" Then
        yearPart = CInt(Left(currentMonth, 4))
        monthPart = CInt(Mid(currentMonth, 6, 2))

        ' 前月計算（1月の場合は前年12月）
        If monthPart = 1 Then
            previousMonth = (yearPart - 1) & "-12"
            previousFolder = (yearPart - 1) & "年"
        Else
            previousMonth = yearPart & "-" & Format(monthPart - 1, "00")
            previousFolder = yearPart & "年"
        End If

        currentYear = CStr(yearPart)
        currentFolder = yearPart & "年"

    Else
        MsgBox "Q2の形式が正しくありません（yyyy-mm形式である必要があります）。", vbExclamation
        GoTo CleanExit
    End If

    Debug.Print "前月: " & previousMonth & " (" & previousFolder & ")"
    Debug.Print "当月: " & currentMonth & " (" & currentFolder & ")"

    ' ============================================
    ' クエリ更新処理
    ' ============================================
    Application.StatusBar = "クエリをスキャン中..."
    updateCount = 0
    i = 0

    For Each qry In ThisWorkbook.Queries
        i = i + 1
        Application.StatusBar = "クエリをチェック中... (" & i & "/" & ThisWorkbook.Queries.Count & ")"

        tempFormula = qry.Formula

        ' ロット数量ADOクエリの処理
        If qry.Name = "ロット数量ADO" Or InStr(qry.Name, "ロット数量ADO") > 0 Then
            newFormula = Createロット数量ADOQuery(currentMonth, previousMonth, currentFolder, previousFolder)
            If newFormula <> tempFormula Then
                qry.Formula = newFormula
                updateCount = updateCount + 1
                Debug.Print "更新: " & qry.Name & " [ロット数量ADO]"
            End If

        ' 不良集計ゾーン別ADOクエリの処理
        ElseIf qry.Name = "不良集計ゾーン別ADO" Or InStr(qry.Name, "不良集計ゾーン別ADO") > 0 Then
            newFormula = Create不良集計ゾーン別ADOQuery(currentMonth, previousMonth, currentFolder, previousFolder)
            If newFormula <> tempFormula Then
                qry.Formula = newFormula
                updateCount = updateCount + 1
                Debug.Print "更新: " & qry.Name & " [不良集計ゾーン別ADO]"
            End If

        ' 番号ADOクエリの処理
        ElseIf qry.Name = "番号ADO" Or InStr(qry.Name, "番号ADO") > 0 Then
            newFormula = Create番号ADOQuery()
            If newFormula <> tempFormula Then
                qry.Formula = newFormula
                updateCount = updateCount + 1
                Debug.Print "更新: " & qry.Name & " [番号ADO]"
            End If
        End If
    Next qry

    ' ============================================
    ' Q1への結果出力
    ' ============================================
    Application.StatusBar = "結果を記録中..."

    If updateCount > 0 Then
        resultMessage = "複数月結合: " & previousMonth & " + " & currentMonth & " (更新数: " & updateCount & ")"
    Else
        resultMessage = "更新対象クエリなし (設定: " & currentMonth & ")"
    End If

    ' Q1に結果を記録
    ws.Range("Q1").Value = resultMessage

    ' ============================================
    ' 全クエリの強制更新
    ' ============================================
    If updateCount > 0 Then
        Application.StatusBar = "クエリを更新中..."
        ThisWorkbook.RefreshAll
        Debug.Print "全クエリの更新を実行"
    End If

    ' ============================================
    ' 処理完了
    ' ============================================
    Debug.Print "========================================"
    Debug.Print "クエリ参照元変更 完了"
    Debug.Print "更新クエリ数: " & updateCount
    Debug.Print "Q1出力: " & resultMessage
    Debug.Print "========================================"

    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "クエリ参照元変更エラー"

CleanExit:
    Application.StatusBar = False

End Sub

' ============================================
' 補助関数：ロット数量ADOクエリ作成
' ============================================
Private Function Createロット数量ADOQuery(currentMonth As String, previousMonth As String, _
                                         currentFolder As String, previousFolder As String) As String

    Dim newFormula As String

    newFormula = "let" & vbCrLf & _
                "    // 前月データ" & vbCrLf & _
                "    前月ソース = Access.Database(File.Contents(""Z:¥全社共有¥オート事業部¥日報¥不良集計¥不良集計表¥" & previousFolder & "¥不良調査表DB-" & previousMonth & ".accdb""), [CreateNavigationProperties=true])," & vbCrLf & _
                "    前月__ロット数量 = 前月ソース{[Schema="""",Item=""_ロット数量""]}[Data]," & vbCrLf & _
                "    前月フィルターされた行 = Table.SelectRows(前月__ロット数量, each [日付] <> null and [日付] <> """")," & vbCrLf & _
                "    // 当月データ" & vbCrLf & _
                "    当月ソース = Access.Database(File.Contents(""Z:¥全社共有¥オート事業部¥日報¥不良集計¥不良集計表¥" & currentFolder & "¥不良調査表DB-" & currentMonth & ".accdb""), [CreateNavigationProperties=true])," & vbCrLf & _
                "    当月__ロット数量 = 当月ソース{[Schema="""",Item=""_ロット数量""]}[Data]," & vbCrLf & _
                "    当月フィルターされた行 = Table.SelectRows(当月__ロット数量, each [日付] <> null and [日付] <> """")," & vbCrLf & _
                "    // 結合" & vbCrLf & _
                "    結合データ = Table.Combine({前月フィルターされた行, 当月フィルターされた行})" & vbCrLf & _
                "in" & vbCrLf & _
                "    結合データ"

    Createロット数量ADOQuery = newFormula

End Function

' ============================================
' 補助関数：不良集計ゾーン別ADOクエリ作成
' ============================================
Private Function Create不良集計ゾーン別ADOQuery(currentMonth As String, previousMonth As String, _
                                               currentFolder As String, previousFolder As String) As String

    Dim newFormula As String

    newFormula = "let" & vbCrLf & _
                "    // 前月データ" & vbCrLf & _
                "    前月ソース = Access.Database(File.Contents(""Z:¥全社共有¥オート事業部¥日報¥不良集計¥不良集計表¥" & previousFolder & "¥不良調査表DB-" & previousMonth & ".accdb""), [CreateNavigationProperties=true])," & vbCrLf & _
                "    前月__不良集計ゾーン別 = 前月ソース{[Schema="""",Item=""_不良集計ゾーン別""]}[Data]," & vbCrLf & _
                "    前月フィルターされた行 = Table.SelectRows(前月__不良集計ゾーン別, each [日付] <> null and [日付] <> """")," & vbCrLf & _
                "    前月変更された型 = Table.TransformColumnTypes(前月フィルターされた行,{{""数量"", Int64.Type}, {""注番月"", Int64.Type}, {""ロット"", Int64.Type}, {""品番末尾"", Int64.Type}, {""差戻し"", Int64.Type}})," & vbCrLf & _
                "    // 当月データ" & vbCrLf & _
                "    当月ソース = Access.Database(File.Contents(""Z:¥全社共有¥オート事業部¥日報¥不良集計¥不良集計表¥" & currentFolder & "¥不良調査表DB-" & currentMonth & ".accdb""), [CreateNavigationProperties=true])," & vbCrLf & _
                "    当月__不良集計ゾーン別 = 当月ソース{[Schema="""",Item=""_不良集計ゾーン別""]}[Data]," & vbCrLf & _
                "    当月フィルターされた行 = Table.SelectRows(当月__不良集計ゾーン別, each [日付] <> null and [日付] <> """")," & vbCrLf & _
                "    当月変更された型 = Table.TransformColumnTypes(当月フィルターされた行,{{""数量"", Int64.Type}, {""注番月"", Int64.Type}, {""ロット"", Int64.Type}, {""品番末尾"", Int64.Type}, {""差戻し"", Int64.Type}})," & vbCrLf & _
                "    // 結合" & vbCrLf & _
                "    結合データ = Table.Combine({前月変更された型, 当月変更された型})" & vbCrLf & _
                "in" & vbCrLf & _
                "    結合データ"

    Create不良集計ゾーン別ADOQuery = newFormula

End Function

' ============================================
' 補助関数：番号ADO固定クエリ作成
' ============================================
Private Function Create番号ADOQuery() As String

    Dim newFormula As String

    ' 番号は固定で2025-09を参照
    newFormula = "let" & vbCrLf & _
                "    // 番号ADO（2025-09固定版）" & vbCrLf & _
                "    // Q2の値に関係なく常に2025年9月のデータを参照" & vbCrLf & _
                "    ソース = Access.Database(File.Contents(""Z:¥全社共有¥オート事業部¥日報¥不良集計¥不良集計表¥2025年¥不良調査表DB-2025-09.accdb""), [CreateNavigationProperties=true])," & vbCrLf & _
                "    __番号 = ソース{[Schema="""",Item=""_番号""]}[Data]" & vbCrLf & _
                "in" & vbCrLf & _
                "    __番号"

    Create番号ADOQuery = newFormula

End Function
