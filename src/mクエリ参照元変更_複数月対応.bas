Attribute VB_Name = "mクエリ参照元変更_複数月対応"
Option Explicit

' ========================================
' マクロ名: 統合クエリ参照元変更（複数月対応）
' 処理概要: 手直し、流出不良調査表、日報成形、日報塗装の4種のクエリを更新
' ソース: アクティブシートのG2（対象シート）
' 処理内容:
'   1. アクティブシートのG2から年月を取得
'   2. 前月を自動計算
'   3. 手直しクエリを年単位ファイル参照に変更
'   4. 流出不良調査表系クエリを複数月結合に変更
'   5. 日報成形系クエリを複数月結合に変更
'   6. 日報塗装系クエリを複数月結合に変更
'   7. 全シートのG1・G2を同期
' ========================================

Sub 統合クエリ参照元変更_複数月対応()

    Application.StatusBar = "統合クエリ参照元の変更を開始します..."

    On Error GoTo ErrorHandler

    ' ============================================
    ' 変数宣言
    ' ============================================
    Dim ws As Worksheet
    Dim activeSheetName As String
    Dim currentMonth As String
    Dim previousMonth As String
    Dim currentYear As String
    Dim currentFolder As String
    Dim previousFolder As String
    Dim qry As WorkbookQuery
    Dim updateCountDB As Integer
    Dim updateCountExcel As Integer
    Dim updateCountSeikei As Integer
    Dim updateCountToso As Integer
    Dim updateCountKako As Integer
    Dim i As Integer
    Dim tempFormula As String
    Dim newFormula As String
    Dim resultMessage As String

    ' 対象シート名配列（拡張版）※手直しシートは同期対象外
    Dim targetSheets As Variant
    targetSheets = Array("成形", "塗装", "加工", "成形N", "塗装N", "成形ND", _
                         "成形T", "成形H", "成形G", "成形NW", _
                         "塗装T", "塗装H", "塗装G", "塗装NW", _
                         "加工T", "加工H", "加工G", "加工NW", "グラフ")

    ' ============================================
    ' アクティブシート確認と入力値取得
    ' ============================================
    Set ws = ActiveSheet
    activeSheetName = ws.Name

    ' 対象シートかどうか確認
    Dim isValidSheet As Boolean
    isValidSheet = False
    Dim j As Integer
    For j = 0 To UBound(targetSheets)
        If activeSheetName = targetSheets(j) Then
            isValidSheet = True
            Exit For
        End If
    Next j

    If Not isValidSheet Then
        MsgBox "このマクロは以下のシートから実行してください：" & vbCrLf & _
               "・基本：手直し、成形、塗装、加工、成形N、塗装N、成形ND" & vbCrLf & _
               "・拡張：成形T、成形H、成形G、成形NW、塗装T、塗装H、塗装G、塗装NW、加工T、加工H、加工G、加工NW" & vbCrLf & _
               "・その他：グラフ" & vbCrLf & vbCrLf & _
               "現在のシート: " & activeSheetName, vbExclamation
        GoTo CleanExit
    End If

    ' G2の値を取得
    currentMonth = Trim(ws.Range("G2").Value)

    ' 入力値チェック
    If currentMonth = "" Then
        MsgBox "G2のDB末尾が設定されていません。", vbExclamation
        GoTo CleanExit
    End If

    Debug.Print "アクティブシート: " & activeSheetName
    Debug.Print "G2の値: " & currentMonth

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
        MsgBox "DB末尾の形式が正しくありません（yyyy-mm形式である必要があります）。", vbExclamation
        GoTo CleanExit
    End If

    Debug.Print "前月: " & previousMonth & " (" & previousFolder & ")"
    Debug.Print "当月: " & currentMonth & " (" & currentFolder & ")"

    ' ============================================
    ' クエリ更新処理
    ' ============================================
    Application.StatusBar = "クエリをスキャン中..."
    updateCountDB = 0
    updateCountExcel = 0
    updateCountSeikei = 0
    updateCountToso = 0
    updateCountKako = 0
    i = 0

    For Each qry In ThisWorkbook.Queries
        i = i + 1
        Application.StatusBar = "クエリをチェック中... (" & i & "/" & ThisWorkbook.Queries.Count & ")"

        tempFormula = qry.Formula

        ' 手直しクエリの処理（_不良集計ゾーン別テーブルを参照）
        If InStr(tempFormula, "_不良集計ゾーン別") > 0 Then
            newFormula = Create手直しQuery(currentYear)
            If newFormula <> tempFormula Then
                qry.Formula = newFormula
                updateCountDB = updateCountDB + 1
                Debug.Print "更新: " & qry.Name & " [手直し]"
            End If

        ' 流出不良調査表系クエリの処理
        ElseIf InStr(tempFormula, "流出不良調査表-") > 0 Then
            newFormula = Create流出不良複数月結合Query(currentMonth, previousMonth, currentFolder, previousFolder)
            If newFormula <> tempFormula Then
                qry.Formula = newFormula
                updateCountExcel = updateCountExcel + 1
                Debug.Print "更新: " & qry.Name & " [流出不良]"
            End If

        ' 日報成形系クエリの処理（クエリ名で判定）
        ElseIf qry.Name = "日報成形" Then
            newFormula = Create日報成形複数月結合Query(currentMonth, previousMonth, currentFolder, previousFolder)
            If newFormula <> tempFormula Then
                qry.Formula = newFormula
                updateCountSeikei = updateCountSeikei + 1
                Debug.Print "更新: " & qry.Name & " [日報成形]"
            End If

        ' 日報塗装系クエリの処理（クエリ名で判定）
        ElseIf qry.Name = "日報塗装" Then
            newFormula = Create日報塗装複数月結合Query(currentMonth, previousMonth, currentFolder, previousFolder)
            If newFormula <> tempFormula Then
                qry.Formula = newFormula
                updateCountToso = updateCountToso + 1
                Debug.Print "更新: " & qry.Name & " [日報塗装]"
            End If

        ' 日報加工系クエリの処理（クエリ名で判定）
        ElseIf qry.Name = "日報加工" Then
            newFormula = Create日報加工複数月結合Query(currentMonth, previousMonth, currentFolder, previousFolder)
            If newFormula <> tempFormula Then
                qry.Formula = newFormula
                updateCountKako = updateCountKako + 1
                Debug.Print "更新: " & qry.Name & " [日報加工]"
            End If
        End If
    Next qry

    ' ============================================
    ' 全シートへのG1・G2同期処理
    ' ============================================
    Application.StatusBar = "全シートを同期中..."

    resultMessage = "複数月結合: " & previousMonth & " + " & currentMonth
    Dim targetSheet As Worksheet
    Dim syncCount As Integer
    syncCount = 0

    For j = 0 To UBound(targetSheets)
        On Error Resume Next
        Set targetSheet = ThisWorkbook.Sheets(targetSheets(j))
        If Not targetSheet Is Nothing Then
            ' G1に結果を記録
            targetSheet.Range("G1").Value = resultMessage
            ' G2に元の値を同期
            targetSheet.Range("G2").Value = currentMonth
            syncCount = syncCount + 1
            Debug.Print "同期完了: " & targetSheets(j)
        End If
        On Error GoTo ErrorHandler
    Next j

    ' ============================================
    ' 処理完了
    ' ============================================
    Debug.Print "========================================"
    Debug.Print "統合クエリ参照元変更 完了"
    Debug.Print "手直しクエリ更新数: " & updateCountDB
    Debug.Print "流出不良調査表クエリ更新数: " & updateCountExcel
    Debug.Print "日報成形クエリ更新数: " & updateCountSeikei
    Debug.Print "日報塗装クエリ更新数: " & updateCountToso
    Debug.Print "日報加工クエリ更新数: " & updateCountKako
    Debug.Print "シート同期数: " & syncCount & "/" & (UBound(targetSheets) + 1)
    Debug.Print "========================================"

    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "統合クエリ参照元変更エラー"

CleanExit:
    Application.StatusBar = False

End Sub

' ============================================
' 補助関数：手直しクエリ作成（年単位ファイル単独参照）
' ============================================
Private Function Create手直しQuery(currentYear As String) As String

    Dim newFormula As String
    Dim yearFolder As String

    yearFolder = currentYear & "年"

    newFormula = "let" & vbCrLf & _
                "    ソース = Access.Database(File.Contents(""Z:¥全社共有¥オート事業部¥日報¥不良集計¥不良集計表¥" & yearFolder & "¥不良調査表DB-" & currentYear & ".accdb""), [CreateNavigationProperties=true])," & vbCrLf & _
                "    テーブル = ソース{[Schema="""",Item=""_不良集計ゾーン別""]}[Data]," & vbCrLf & _
                "    削除された他の列 = Table.SelectColumns(テーブル,{""ID"", ""日付"", ""品番"", ""ロット"", ""発見"", ""ゾーン"", ""番号"", ""数量"", ""差戻し""})," & vbCrLf & _
                "    変更された型 = Table.TransformColumnTypes(削除された他の列,{{""数量"", Int64.Type}, {""差戻し"", Int64.Type}})" & vbCrLf & _
                "in" & vbCrLf & _
                "    変更された型"

    Create手直しQuery = newFormula

End Function

' ============================================
' 補助関数：流出不良複数月結合クエリ作成
' ============================================
Private Function Create流出不良複数月結合Query(currentMonth As String, previousMonth As String, _
                                              currentFolder As String, previousFolder As String) As String

    Dim newFormula As String

    newFormula = "let" & vbCrLf & _
                "    // 前月データ" & vbCrLf & _
                "    前月ソース = Excel.Workbook(File.Contents(""Z:¥全社共有¥オート事業部¥日報¥不良集計¥流出不良調査表¥" & previousFolder & "¥流出不良調査表-" & previousMonth & ".xlsm""), null, true)," & vbCrLf & _
                "    前月_集計_Table = 前月ソース{[Item=""_集計"",Kind=""Table""]}[Data]," & vbCrLf & _
                "    前月変更された型 = Table.TransformColumnTypes(前月_集計_Table,{{""日付"", type date}, {""品番"", type text}, {""Fr/Rr"", type text}, {""ロット"", Int64.Type}, {""テープ加工"", type text}, {""工程"", type text}, {""不良内容"", type text}, {""R/L"", type text}, {""件数"", Int64.Type}, {""担当"", type text}})," & vbCrLf & _
                "    前月フィルターされた行 = Table.SelectRows(前月変更された型, each ([日付] <> null))," & vbCrLf & _
                "    // 当月データ" & vbCrLf & _
                "    当月ソース = Excel.Workbook(File.Contents(""Z:¥全社共有¥オート事業部¥日報¥不良集計¥流出不良調査表¥" & currentFolder & "¥流出不良調査表-" & currentMonth & ".xlsm""), null, true)," & vbCrLf & _
                "    当月_集計_Table = 当月ソース{[Item=""_集計"",Kind=""Table""]}[Data]," & vbCrLf & _
                "    当月変更された型 = Table.TransformColumnTypes(当月_集計_Table,{{""日付"", type date}, {""品番"", type text}, {""Fr/Rr"", type text}, {""ロット"", Int64.Type}, {""テープ加工"", type text}, {""工程"", type text}, {""不良内容"", type text}, {""R/L"", type text}, {""件数"", Int64.Type}, {""担当"", type text}})," & vbCrLf & _
                "    当月フィルターされた行 = Table.SelectRows(当月変更された型, each ([日付] <> null))," & vbCrLf & _
                "    // 結合" & vbCrLf & _
                "    結合データ = Table.Combine({前月フィルターされた行, 当月フィルターされた行})" & vbCrLf & _
                "in" & vbCrLf & _
                "    結合データ"

    Create流出不良複数月結合Query = newFormula

End Function

' ============================================
' 補助関数：日報成形複数月結合クエリ作成
' ============================================
Private Function Create日報成形複数月結合Query(currentMonth As String, previousMonth As String, _
                                              currentFolder As String, previousFolder As String) As String

    Dim newFormula As String
    Dim typeList As String
    Dim columnList As String
    Dim replaceList1 As String
    Dim replaceList2 As String

    ' 型定義リスト（前月・当月共通）
    typeList = "{{""日付"", type date}, {""開始時間"", type datetime}, {""終了時間"", type datetime}, "
    typeList = typeList & "{""所要時間"", type number}, {""型替"", type number}, {""号機"", Int64.Type}, "
    typeList = typeList & "{""品番"", type text}, {""材料"", type text}, {""サイクル"", type number}, "
    typeList = typeList & "{""秒/ショット"", type number}, {""出来高率"", type number}, {""ショット数"", Int64.Type}, "
    typeList = typeList & "{""不良数"", Int64.Type}, {""不良率"", type number}, {""打出し"", Int64.Type}, "
    typeList = typeList & "{""ショート"", type any}, {""ウエルド"", Int64.Type}, {""シワ"", type any}, "
    typeList = typeList & "{""異物"", Int64.Type}, {""シルバー"", Int64.Type}, {""フローマーク"", Int64.Type}, "
    typeList = typeList & "{""ゴミ押し"", Int64.Type}, {""GCカス"", Int64.Type}, {""キズ"", Int64.Type}, "
    typeList = typeList & "{""ヒケ"", Int64.Type}, {""糸引き"", Int64.Type}, {""型汚れ"", Int64.Type}, "
    typeList = typeList & "{""マクレ"", Int64.Type}, {""取出不良"", Int64.Type}, {""割れ白化"", Int64.Type}, "
    typeList = typeList & "{""コアカス"", type any}, {""その他"", Int64.Type}, {""チョコ停打出し"", Int64.Type}, "
    typeList = typeList & "{""検査"", Int64.Type}, {""流出不良"", Int64.Type}, {""理論数"", Int64.Type}, "
    typeList = typeList & "{""コメント"", type text}, {""コメント２"", type any}, {""補助1"", Int64.Type}}"

    ' 選択列リスト
    columnList = "{""日付"", ""所要時間"", ""型替"", ""号機"", ""品番"", ""ショット数"", ""不良数"", "
    columnList = columnList & """打出し"", ""ショート"", ""ウエルド"", ""シワ"", ""異物"", ""シルバー"", "
    columnList = columnList & """フローマーク"", ""ゴミ押し"", ""GCカス"", ""キズ"", ""ヒケ"", ""糸引き"", "
    columnList = columnList & """型汚れ"", ""マクレ"", ""取出不良"", ""割れ白化"", ""コアカス"", ""その他"", "
    columnList = columnList & """チョコ停打出し"", ""検査"", ""流出不良"", ""コメント""}"

    ' 置換列リスト1
    replaceList1 = "{""不良数"", ""打出し"", ""ショート"", ""ウエルド"", ""シワ"", ""異物"", ""シルバー"", "
    replaceList1 = replaceList1 & """フローマーク"", ""ゴミ押し"", ""GCカス"", ""キズ"", ""ヒケ"", ""糸引き"", "
    replaceList1 = replaceList1 & """型汚れ"", ""マクレ"", ""取出不良"", ""割れ白化"", ""コアカス"", ""その他"", "
    replaceList1 = replaceList1 & """チョコ停打出し"", ""検査"", ""流出不良""}"

    ' 置換列リスト2
    replaceList2 = "{""所要時間"", ""型替"", ""ショット数""}"

    ' クエリ本体を構築
    newFormula = "let" & vbCrLf
    newFormula = newFormula & "    // 前月データ" & vbCrLf
    newFormula = newFormula & "    前月ソース = Excel.Workbook(File.Contents(""Z:¥全社共有¥オート事業部¥日報¥成形日報¥" & previousFolder
    newFormula = newFormula & "¥SEIKEI MES-" & previousMonth & ".xlsm""), null, true)," & vbCrLf
    newFormula = newFormula & "    前月集計1_Table = 前月ソース{[Item=""集計1"",Kind=""Table""]}[Data]," & vbCrLf
    newFormula = newFormula & "    前月変更された型 = Table.TransformColumnTypes(前月集計1_Table," & typeList & ")," & vbCrLf
    newFormula = newFormula & "    前月削除された下の行 = Table.RemoveLastN(前月変更された型,1)," & vbCrLf
    newFormula = newFormula & "    前月フィルターされた行 = Table.SelectRows(前月削除された下の行, each ([日付] <> null))," & vbCrLf
    newFormula = newFormula & "    前月削除された他の列 = Table.SelectColumns(前月フィルターされた行," & columnList & ")," & vbCrLf
    newFormula = newFormula & "    前月フィルターされた行1 = Table.SelectRows(前月削除された他の列, each ([品番] <> ""CHOUREI（朝礼）"" "
    newFormula = newFormula & "and [品番] <> ""KEIKAKU（故障）"" and [品番] <> ""KEIKAKU（計画）"" "
    newFormula = newFormula & "and [品番] <> ""RYUUSYUTU（流出）"" and [品番] <> ""TRY（トライ）"" "
    newFormula = newFormula & "and [品番] <> ""QC"" and [品番] <> ""700B"" and [品番] <> ""670B"" and [品番] <> ""032D""))," & vbCrLf
    newFormula = newFormula & "    前月置き換えられた値 = Table.ReplaceValue(前月フィルターされた行1,null,0,Replacer.ReplaceValue," & replaceList1 & ")," & vbCrLf
    newFormula = newFormula & "    前月置き換えられた値1 = Table.ReplaceValue(前月置き換えられた値,null,0,Replacer.ReplaceValue," & replaceList2 & ")," & vbCrLf
    newFormula = newFormula & "    前月列名変更 = Table.RenameColumns(前月置き換えられた値1, {{""その他"",""その他O""}})," & vbCrLf
    newFormula = newFormula & vbCrLf
    ' 当月データ部分
    newFormula = newFormula & "    // 当月データ" & vbCrLf
    newFormula = newFormula & "    当月ソース = Excel.Workbook(File.Contents(""Z:¥全社共有¥オート事業部¥日報¥成形日報¥" & currentFolder
    newFormula = newFormula & "¥SEIKEI MES-" & currentMonth & ".xlsm""), null, true)," & vbCrLf
    newFormula = newFormula & "    当月集計1_Table = 当月ソース{[Item=""集計1"",Kind=""Table""]}[Data]," & vbCrLf
    newFormula = newFormula & "    当月変更された型 = Table.TransformColumnTypes(当月集計1_Table," & typeList & ")," & vbCrLf
    newFormula = newFormula & "    当月削除された下の行 = Table.RemoveLastN(当月変更された型,1)," & vbCrLf
    newFormula = newFormula & "    当月フィルターされた行 = Table.SelectRows(当月削除された下の行, each ([日付] <> null))," & vbCrLf
    newFormula = newFormula & "    当月削除された他の列 = Table.SelectColumns(当月フィルターされた行," & columnList & ")," & vbCrLf
    newFormula = newFormula & "    当月フィルターされた行1 = Table.SelectRows(当月削除された他の列, each ([品番] <> ""CHOUREI（朝礼）"" "
    newFormula = newFormula & "and [品番] <> ""KEIKAKU（故障）"" and [品番] <> ""KEIKAKU（計画）"" "
    newFormula = newFormula & "and [品番] <> ""RYUUSYUTU（流出）"" and [品番] <> ""TRY（トライ）"" "
    newFormula = newFormula & "and [品番] <> ""QC"" and [品番] <> ""700B"" and [品番] <> ""670B"" and [品番] <> ""032D""))," & vbCrLf
    newFormula = newFormula & "    当月置き換えられた値 = Table.ReplaceValue(当月フィルターされた行1,null,0,Replacer.ReplaceValue," & replaceList1 & ")," & vbCrLf
    newFormula = newFormula & "    当月置き換えられた値1 = Table.ReplaceValue(当月置き換えられた値,null,0,Replacer.ReplaceValue," & replaceList2 & ")," & vbCrLf
    newFormula = newFormula & "    当月列名変更 = Table.RenameColumns(当月置き換えられた値1, {{""その他"",""その他O""}})," & vbCrLf
    newFormula = newFormula & vbCrLf
    ' 結合部分
    newFormula = newFormula & "    // 結合" & vbCrLf
    newFormula = newFormula & "    結合データ = Table.Combine({前月列名変更, 当月列名変更})" & vbCrLf
    newFormula = newFormula & "in" & vbCrLf
    newFormula = newFormula & "    結合データ"

    Create日報成形複数月結合Query = newFormula

End Function

' ============================================
' 補助関数：日報塗装複数月結合クエリ作成
' ============================================
Private Function Create日報塗装複数月結合Query(currentMonth As String, previousMonth As String, _
                                              currentFolder As String, previousFolder As String) As String

    Dim newFormula As String

    newFormula = "let" & vbCrLf
    newFormula = newFormula & "    // 前月データ" & vbCrLf
    newFormula = newFormula & "    前月ソース = Excel.Workbook(File.Contents(""Z:¥全社共有¥オート事業部¥日報¥塗装日報¥" & previousFolder
    newFormula = newFormula & "¥塗装日報まとめTOSO_" & previousMonth & ".xlsm""), null, true)," & vbCrLf
    newFormula = newFormula & "    前月塗装集計_Table = 前月ソース{[Item=""塗装集計"",Kind=""Table""]}[Data]," & vbCrLf
    newFormula = newFormula & "    前月削除された下の行 = Table.RemoveLastN(前月塗装集計_Table, 1)," & vbCrLf
    newFormula = newFormula & vbCrLf
    newFormula = newFormula & "    // 当月データ" & vbCrLf
    newFormula = newFormula & "    当月ソース = Excel.Workbook(File.Contents(""Z:¥全社共有¥オート事業部¥日報¥塗装日報¥" & currentFolder
    newFormula = newFormula & "¥塗装日報まとめTOSO_" & currentMonth & ".xlsm""), null, true)," & vbCrLf
    newFormula = newFormula & "    当月塗装集計_Table = 当月ソース{[Item=""塗装集計"",Kind=""Table""]}[Data]," & vbCrLf
    newFormula = newFormula & "    当月削除された下の行 = Table.RemoveLastN(当月塗装集計_Table, 1)," & vbCrLf
    newFormula = newFormula & vbCrLf
    newFormula = newFormula & "    // 結合" & vbCrLf
    newFormula = newFormula & "    連結 = Table.Combine({前月削除された下の行, 当月削除された下の行})," & vbCrLf
    newFormula = newFormula & vbCrLf
    newFormula = newFormula & "    // データ加工処理" & vbCrLf
    newFormula = newFormula & "    フィルターされた行1 = Table.SelectRows(連結, each ([日付] <> null))," & vbCrLf
    newFormula = newFormula & vbCrLf
    newFormula = newFormula & "    // ===== 修正1: 列選択を型変換の前に実施 =====" & vbCrLf
    newFormula = newFormula & "    削除された他の列 = Table.SelectColumns(フィルターされた行1,{" & vbCrLf
    newFormula = newFormula & "        ""日付"",""品番"",""L/R"",""ショット数"",""リコート"",""廃棄"",""ヒゲ"",""ミスト"",""ライン""," & vbCrLf
    newFormula = newFormula & "        ""ゴミ"",""スケ"",""ピンホール"",""マット"",""その他"",""ゴミ2"",""タレ"",""キズ"",""再塗装"",""成形"",""その他2""" & vbCrLf
    newFormula = newFormula & "    })," & vbCrLf
    newFormula = newFormula & vbCrLf
    newFormula = newFormula & "    // ===== 修正2: null値を0に置換（型変換の前に実施）=====" & vbCrLf
    newFormula = newFormula & "    null置き換え = Table.ReplaceValue(削除された他の列, null, 0, Replacer.ReplaceValue," & vbCrLf
    newFormula = newFormula & "        {""ショット数"",""リコート"",""廃棄"",""ヒゲ"",""ミスト"",""ライン"",""ゴミ"",""スケ"",""ピンホール"",""マット"",""その他"",""ゴミ2"",""タレ"",""キズ"",""再塗装"",""成形"",""その他2""})," & vbCrLf
    newFormula = newFormula & vbCrLf
    newFormula = newFormula & "    // ===== 修正3: 型変換（type any を Int64.Type に変更）=====" & vbCrLf
    newFormula = newFormula & "    変更された型 = Table.TransformColumnTypes(null置き換え,{" & vbCrLf
    newFormula = newFormula & "        {""日付"", type date}," & vbCrLf
    newFormula = newFormula & "        {""品番"", type text}," & vbCrLf
    newFormula = newFormula & "        {""L/R"", type text}," & vbCrLf
    newFormula = newFormula & "        {""ショット数"", Int64.Type}," & vbCrLf
    newFormula = newFormula & "        {""リコート"", Int64.Type}," & vbCrLf
    newFormula = newFormula & "        {""廃棄"", Int64.Type}," & vbCrLf
    newFormula = newFormula & "        {""ヒゲ"", Int64.Type}," & vbCrLf
    newFormula = newFormula & "        {""ミスト"", Int64.Type}," & vbCrLf
    newFormula = newFormula & "        {""ライン"", Int64.Type}," & vbCrLf
    newFormula = newFormula & "        {""ゴミ"", Int64.Type}," & vbCrLf
    newFormula = newFormula & "        {""スケ"", Int64.Type}," & vbCrLf
    newFormula = newFormula & "        {""ピンホール"", Int64.Type}," & vbCrLf
    newFormula = newFormula & "        {""マット"", Int64.Type}," & vbCrLf
    newFormula = newFormula & "        {""その他"", Int64.Type}," & vbCrLf
    newFormula = newFormula & "        {""ゴミ2"", Int64.Type}," & vbCrLf
    newFormula = newFormula & "        {""タレ"", Int64.Type}," & vbCrLf
    newFormula = newFormula & "        {""キズ"", Int64.Type}," & vbCrLf
    newFormula = newFormula & "        {""再塗装"", Int64.Type}," & vbCrLf
    newFormula = newFormula & "        {""成形"", Int64.Type}," & vbCrLf
    newFormula = newFormula & "        {""その他2"", Int64.Type}" & vbCrLf
    newFormula = newFormula & "    })," & vbCrLf
    newFormula = newFormula & vbCrLf
    newFormula = newFormula & "    // リコート + 廃棄 -> 不良数" & vbCrLf
    newFormula = newFormula & "    追加_不良数 = Table.AddColumn(変更された型, ""不良数"", each [リコート] + [廃棄], type number)," & vbCrLf
    newFormula = newFormula & vbCrLf
    newFormula = newFormula & "    // ゴミ + ゴミ2 -> ゴミ" & vbCrLf
    newFormula = newFormula & "    追加_ゴミ_new = Table.AddColumn(追加_不良数, ""ゴミ_new"", each [ゴミ] + [ゴミ2], type number)," & vbCrLf
    newFormula = newFormula & "    削除_ゴミ元 = Table.RemoveColumns(追加_ゴミ_new, {""ゴミ"",""ゴミ2""})," & vbCrLf
    newFormula = newFormula & "    名前変更_ゴミ = Table.RenameColumns(削除_ゴミ元, {{""ゴミ_new"",""ゴミ""}})," & vbCrLf
    newFormula = newFormula & vbCrLf
    newFormula = newFormula & "    // その他 + その他2 -> その他O" & vbCrLf
    newFormula = newFormula & "    追加_その他_new = Table.AddColumn(名前変更_ゴミ, ""その他O"", each [その他] + [その他2], type number)," & vbCrLf
    newFormula = newFormula & "    削除_その他元 = Table.RemoveColumns(追加_その他_new, {""その他"",""その他2""})," & vbCrLf
    newFormula = newFormula & vbCrLf
    newFormula = newFormula & "    フィルターされた行 = Table.SelectRows(削除_その他元, each ([日付] <> null))," & vbCrLf
    newFormula = newFormula & vbCrLf
    newFormula = newFormula & "    並べ替えられた列 = Table.ReorderColumns(フィルターされた行," & vbCrLf
    newFormula = newFormula & "        {""日付"",""品番"",""L/R"",""ショット数"",""不良数"",""リコート"",""廃棄"",""ヒゲ"",""ミスト"",""ライン"",""スケ"",""ピンホール"",""マット"",""タレ"",""キズ"",""再塗装"",""成形"",""ゴミ"",""その他O""})," & vbCrLf
    newFormula = newFormula & vbCrLf
    newFormula = newFormula & "    フィルターされた行2 = Table.SelectRows(並べ替えられた列, each ([ショット数] <> 0))" & vbCrLf
    newFormula = newFormula & "in" & vbCrLf
    newFormula = newFormula & "    フィルターされた行2"

    Create日報塗装複数月結合Query = newFormula

End Function

' ============================================
' 補助関数：日報加工複数月結合クエリ作成
' ============================================
Private Function Create日報加工複数月結合Query(currentMonth As String, previousMonth As String, _
                                              currentFolder As String, previousFolder As String) As String

    Dim newFormula As String
    Dim typeList As String
    Dim columnList As String
    Dim replaceList As String

    ' 型定義リスト（元クエリの変更された型ベース）
    typeList = "{{""生産日"", type date}, {""ライン"", Int64.Type}, {""作業者名"", type text}, "
    typeList = typeList & "{""品番"", type text}, {""ロット"", type text}, {""数量"", type number}, "
    typeList = typeList & "{""開始時間"", type datetime}, {""終了時間"", type datetime}, {""休憩時間"", Int64.Type}, "
    typeList = typeList & "{""段取時間"", Int64.Type}, {""成形不良"", type any}, {""プライマー付着"", type any}, "
    typeList = typeList & "{""テープ貼り失敗"", type any}, {""テープ蛇行"", type any}, {""テープ内異物"", type any}, "
    typeList = typeList & "{""キズ付け"", type any}, {""その他"", type any}, {""廃棄合計"", Int64.Type}, "
    typeList = typeList & "{""稼働時間(分)"", type number}, {""稼働時間"", type number}, {""出来高"", Int64.Type}, "
    typeList = typeList & "{""備考"", type text}, {""ヘルパー"", type any}, {""名称"", type text}, {""セット/単品"", type text}}"

    ' 選択列リスト（出力に必要な列）
    columnList = "{""生産日"", ""品番"", ""ロット"", ""数量"", ""成形不良"", ""プライマー付着"", "
    columnList = columnList & """テープ貼り失敗"", ""テープ蛇行"", ""テープ内異物"", ""キズ付け"", ""その他"", ""廃棄合計""}"

    ' null置換対象列リスト
    replaceList = "{""成形不良"", ""プライマー付着"", ""テープ貼り失敗"", ""テープ蛇行"", ""テープ内異物"", ""キズ付け"", ""その他""}"

    ' クエリ本体を構築
    newFormula = "let" & vbCrLf
    newFormula = newFormula & "    // 前月データ" & vbCrLf
    newFormula = newFormula & "    前月ソース = Excel.Workbook(File.Contents(""Z:¥全社共有¥オート事業部¥日報¥加工日報¥" & previousFolder
    newFormula = newFormula & "¥加工日報まとめ-" & previousMonth & ".xlsm""), null, true)," & vbCrLf
    newFormula = newFormula & "    前月加工日報_Table = 前月ソース{[Item=""加工日報"",Kind=""Table""]}[Data]," & vbCrLf
    newFormula = newFormula & "    前月変更された型 = Table.TransformColumnTypes(前月加工日報_Table," & typeList & ")," & vbCrLf
    newFormula = newFormula & "    前月削除された下の行 = Table.RemoveLastN(前月変更された型,1)," & vbCrLf
    newFormula = newFormula & "    前月フィルターされた行 = Table.SelectRows(前月削除された下の行, each ([生産日] <> null))," & vbCrLf
    newFormula = newFormula & "    前月削除された他の列 = Table.SelectColumns(前月フィルターされた行," & columnList & ")," & vbCrLf
    newFormula = newFormula & "    前月名前が変更された列 = Table.RenameColumns(前月削除された他の列,{{""生産日"", ""日付""}, {""数量"", ""ショット数""}, {""廃棄合計"", ""不良数""}})," & vbCrLf
    newFormula = newFormula & "    前月変更された型1 = Table.TransformColumnTypes(前月名前が変更された列,{{""成形不良"", type number}, {""プライマー付着"", type number}, {""テープ貼り失敗"", type number}, {""テープ蛇行"", type number}, {""テープ内異物"", type number}, {""キズ付け"", type number}, {""その他"", type number}})," & vbCrLf
    newFormula = newFormula & "    前月置き換えられた値 = Table.ReplaceValue(前月変更された型1,null,0,Replacer.ReplaceValue," & replaceList & ")," & vbCrLf
    newFormula = newFormula & "    前月並べ替えられた列 = Table.ReorderColumns(前月置き換えられた値,{""日付"", ""品番"", ""ロット"", ""ショット数"", ""不良数"", ""成形不良"", ""プライマー付着"", ""テープ貼り失敗"", ""テープ蛇行"", ""テープ内異物"", ""キズ付け"", ""その他""})," & vbCrLf
    newFormula = newFormula & vbCrLf
    ' 当月データ部分
    newFormula = newFormula & "    // 当月データ" & vbCrLf
    newFormula = newFormula & "    当月ソース = Excel.Workbook(File.Contents(""Z:¥全社共有¥オート事業部¥日報¥加工日報¥" & currentFolder
    newFormula = newFormula & "¥加工日報まとめ-" & currentMonth & ".xlsm""), null, true)," & vbCrLf
    newFormula = newFormula & "    当月加工日報_Table = 当月ソース{[Item=""加工日報"",Kind=""Table""]}[Data]," & vbCrLf
    newFormula = newFormula & "    当月変更された型 = Table.TransformColumnTypes(当月加工日報_Table," & typeList & ")," & vbCrLf
    newFormula = newFormula & "    当月削除された下の行 = Table.RemoveLastN(当月変更された型,1)," & vbCrLf
    newFormula = newFormula & "    当月フィルターされた行 = Table.SelectRows(当月削除された下の行, each ([生産日] <> null))," & vbCrLf
    newFormula = newFormula & "    当月削除された他の列 = Table.SelectColumns(当月フィルターされた行," & columnList & ")," & vbCrLf
    newFormula = newFormula & "    当月名前が変更された列 = Table.RenameColumns(当月削除された他の列,{{""生産日"", ""日付""}, {""数量"", ""ショット数""}, {""廃棄合計"", ""不良数""}})," & vbCrLf
    newFormula = newFormula & "    当月変更された型1 = Table.TransformColumnTypes(当月名前が変更された列,{{""成形不良"", type number}, {""プライマー付着"", type number}, {""テープ貼り失敗"", type number}, {""テープ蛇行"", type number}, {""テープ内異物"", type number}, {""キズ付け"", type number}, {""その他"", type number}})," & vbCrLf
    newFormula = newFormula & "    当月置き換えられた値 = Table.ReplaceValue(当月変更された型1,null,0,Replacer.ReplaceValue," & replaceList & ")," & vbCrLf
    newFormula = newFormula & "    当月並べ替えられた列 = Table.ReorderColumns(当月置き換えられた値,{""日付"", ""品番"", ""ロット"", ""ショット数"", ""不良数"", ""成形不良"", ""プライマー付着"", ""テープ貼り失敗"", ""テープ蛇行"", ""テープ内異物"", ""キズ付け"", ""その他""})," & vbCrLf
    newFormula = newFormula & vbCrLf
    ' 結合部分
    newFormula = newFormula & "    // 結合" & vbCrLf
    newFormula = newFormula & "    結合データ = Table.Combine({前月並べ替えられた列, 当月並べ替えられた列})" & vbCrLf
    newFormula = newFormula & "in" & vbCrLf
    newFormula = newFormula & "    結合データ"

    Create日報加工複数月結合Query = newFormula

End Function
