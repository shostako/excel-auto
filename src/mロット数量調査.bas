Attribute VB_Name = "mロット数量調査"
' ========================================
' マクロ名: ロット数量調査_高速化
' 処理概要: 調査テーブルの品番・工程・日付に対応するロット数量を高速検索して転記
' ソーステーブル: シート「（複数）」テーブル「_ロット数量ADO」
'                 シート「ActiveSheet等」テーブル「_調査」「_調査3」「_調査4」
' ターゲットテーブル: 各調査テーブルの「ロット数量」列
' 高速化手法: Dictionary（二段階キー構造）+ 配列処理
' ========================================

Sub ロット数量調査_高速化(Optional テーブル番号 As Integer = 0)
    ' ----- 変数宣言 -----
    ' ワークシート・テーブル関連
    Dim ws調査 As Worksheet                 ' 処理対象の調査シート
    Dim lo調査 As ListObject                ' 処理対象の調査テーブル
    Dim loロット数量 As ListObject          ' ロット数量ADOテーブル

    ' 配列関連
    Dim surveyData As Variant               ' _調査テーブルのデータ配列
    Dim lotData As Variant                  ' _ロット数量ADOテーブルのデータ配列
    Dim resultData() As Variant             ' 結果（ロット数量）を格納する配列

    ' Dictionary関連（高速検索用）
    Dim lookupDict As Object                ' 外側Dictionary (Key: 品番&工程, Value: innerDict)
    Dim innerDict As Object                 ' 内側Dictionary (Key: 日付シリアル値, Value: ロット数量文字列)

    ' 列インデックス（調査テーブル）
    Dim 品番調査Col As Long
    Dim 工程調査Col As Long
    Dim 日付調査Col As Long
    Dim ロット数量調査Col As Long

    ' 列インデックス（ロット数量ADOテーブル）
    Dim 品番ロットCol As Long
    Dim 工程ロットCol As Long
    Dim 日付ロットCol As Long
    Dim ロット数量ロットCol As Long

    ' ループカウンター・カウント
    Dim i As Long, j As Long                ' ループカウンター
    Dim 処理行数 As Long, 全行数 As Long    ' 行数カウント
    Dim errorCount As Long                  ' 日付変換エラーカウント

    ' 検索キー関連
    Dim key As String                       ' 複合キー（品番&工程）
    Dim arr日付() As String                 ' パイプ区切り日付の分割配列
    Dim currentDateStr As String            ' 現在処理中の日付文字列
    Dim currentDateNum As Double            ' 日付のシリアル値（整数部分）

    ' 結果関連
    Dim resultLot数量 As String             ' 検索結果のロット数量（パイプ区切り）

    ' セル値取得用
    Dim 品番 As Variant
    Dim 工程 As Variant
    Dim 日付 As Variant
    Dim ロット数量値 As Variant

    ' テーブル・シート指定関連
    Dim テーブル名 As String                ' 処理対象の調査テーブル名
    Dim シート名 As String                  ' 処理対象のシート名

    ' 進捗・パフォーマンス関連
    Dim 進捗率 As String                    ' 進捗率（%表示用）
    Dim startTime As Double                 ' 処理開始時刻

    ' 処理開始時刻を記録
    startTime = Timer

    ' ============================================
    ' アプリケーション設定：高速化モード
    ' ============================================
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual ' 手動計算モード
    Application.EnableEvents = False ' イベント発生を抑止
    errorCount = 0

    ' 初期メッセージの表示
    Application.StatusBar = "処理の準備をしています..."

    ' ============================================
    ' テーブル番号に基づく対象テーブル・シートの決定
    ' ============================================
    Select Case テーブル番号
        Case 0
            テーブル名 = "_調査"
            シート名 = "" ' ActiveSheetを使用 (既存コードとの互換性維持)
        Case 3
            テーブル名 = "_調査3"
            シート名 = "セット品"
        Case 4
            テーブル名 = "_調査4"
            シート名 = "単品"
        Case Else
            ' 不明なテーブル番号
            Application.StatusBar = "エラー: テーブル番号 " & テーブル番号 & " に対応するシートが定義されていません。"
            GoTo CleanExit
    End Select

    ' ============================================
    ' 調査シートの取得
    ' ============================================
    ' シート名が指定されている場合はそのシートを使用、そうでなければActiveSheetを使用
    If シート名 <> "" Then
        On Error Resume Next
        Set ws調査 = ThisWorkbook.Sheets(シート名)
        On Error GoTo 0

        If ws調査 Is Nothing Then
            Application.StatusBar = "エラー: シート「" & シート名 & "」が見つかりません。"
            GoTo CleanExit
        End If
    Else
        Set ws調査 = ActiveSheet
    End If

    ' ============================================
    ' 調査テーブルの取得
    ' ============================================
    ' 指定されたシート内でテーブルを検索
    On Error Resume Next
    Set lo調査 = ws調査.ListObjects(テーブル名)
    On Error GoTo 0

    If lo調査 Is Nothing Then
        Application.StatusBar = "エラー: シート「" & ws調査.Name & "」に「" & テーブル名 & "」テーブルが見つかりません。"
        GoTo CleanExit
    End If

    ' ============================================
    ' ロット数量ADOテーブルの取得（ブック内全シート検索）
    ' ============================================
    Set loロット数量 = GetListObjectByName("_ロット数量ADO")
    If loロット数量 Is Nothing Then
        Application.StatusBar = "エラー: 「_ロット数量ADO」テーブルが見つかりません。"
        GoTo CleanExit ' 終了処理へ
    End If

    ' ============================================
    ' ロット数量ADOテーブルのデータをDictionaryに読み込む
    ' ============================================
    Application.StatusBar = "ロット数量データを読み込んでいます..."
    Set lookupDict = CreateObject("Scripting.Dictionary")

    ' 列インデックスを取得（エラーチェック付き）
    Err.Clear ' エラー状態をクリア
    On Error Resume Next ' 列が存在しない場合のエラーを捕捉
    品番ロットCol = loロット数量.ListColumns("品番").Index
    工程ロットCol = loロット数量.ListColumns("工程").Index
    日付ロットCol = loロット数量.ListColumns("日付").Index
    ロット数量ロットCol = loロット数量.ListColumns("ロット数量").Index
    If Err.Number <> 0 Then
        Application.StatusBar = "エラー: 「_ロット数量ADO」テーブルに必要な列（品番, 工程, 日付, ロット数量）が見つかりません。"
        On Error GoTo 0 ' エラーハンドリングを通常に戻す
        GoTo CleanExit ' 終了処理へ
    End If
    On Error GoTo 0 ' エラーハンドリングを通常に戻す

    ' データ範囲を配列に読み込み (.Value2で数値や日付シリアル値を取得)
    Dim lotDataRange As Range
    Dim lotRowCount As Long

    If loロット数量.ListRows.Count > 0 Then
        Set lotDataRange = loロット数量.DataBodyRange
        lotData = lotDataRange.Value2 ' データが1行でも配列になるはず
        lotRowCount = loロット数量.ListRows.Count ' ListRows.Countを使う方が確実
    Else
        lotRowCount = 0
        lotData = Array() ' データ行が0件の場合は空の配列
    End If

    ' ============================================
    ' Dictionary構築（品番・工程・日付の三重キー構造）
    ' ============================================
    For i = 1 To lotRowCount
        ' 配列から値を取得
        If lotRowCount = 1 Then
            ' データが1行の場合、lotDataが1次元配列か2次元配列(1, N)か判定が必要
            If IsArray(lotData) Then
                Dim isOneDim As Boolean
                isOneDim = False
                On Error Resume Next ' UBoundの2次元アクセスでエラーが出るかで判定
                Dim checkDim As Long
                checkDim = UBound(lotData, 2) ' 2番目の次元の上限を取得試行
                If Err.Number <> 0 Then
                    isOneDim = True ' 2次元目のUBoundでエラー -> 1次元配列
                    Err.Clear
                End If
                On Error GoTo 0

                If isOneDim Then
                    ' 1次元配列の場合 (lotDataの要素は1から始まる)
                    If UBound(lotData) >= 品番ロットCol Then 品番 = lotData(品番ロットCol) Else 品番 = Empty
                    If UBound(lotData) >= 工程ロットCol Then 工程 = lotData(工程ロットCol) Else 工程 = Empty
                    If UBound(lotData) >= 日付ロットCol Then 日付 = lotData(日付ロットCol) Else 日付 = Empty
                    If UBound(lotData) >= ロット数量ロットCol Then ロット数量値 = lotData(ロット数量ロットCol) Else ロット数量値 = Empty
                Else
                    ' 2次元配列(1行)の場合 (lotData(1, N))
                    品番 = lotData(1, 品番ロットCol)
                    工程 = lotData(1, 工程ロットCol)
                    日付 = lotData(1, 日付ロットCol)
                    ロット数量値 = lotData(1, ロット数量ロットCol)
                End If
            Else
                 ' lotDataが配列ではない異常ケース（通常発生しないはず）
                 GoTo SkipLotRow
            End If
        Else ' lotRowCount > 1 (通常の2次元配列)
             品番 = lotData(i, 品番ロットCol)
             工程 = lotData(i, 工程ロットCol)
             日付 = lotData(i, 日付ロットCol) ' Value2なので日付シリアル値(Double)のはず
             ロット数量値 = lotData(i, ロット数量ロットCol)
        End If

        ' 品番、工程、日付、ロット数量が有効な値かチェック
        If Not IsEmpty(品番) And Not IsEmpty(工程) And IsNumeric(日付) And Not IsEmpty(ロット数量値) Then
            key = CStr(品番) & vbTab & CStr(工程) ' 複合キーを作成

            ' 日付をDouble型に変換し、整数部分（日付のみ）をキーとする
            Dim currentLotDateNum As Double
            ' 日付が数値でない場合のエラーを考慮
            On Error Resume Next
            currentLotDateNum = Int(CDbl(日付))
            If Err.Number <> 0 Then
                ' 日付変換エラーの場合は、この行をスキップ
                Err.Clear
                GoTo SkipLotRow
            End If
            On Error GoTo 0

            If lookupDict.Exists(key) Then
                ' 既に品番・工程のキーが存在する場合
                Set innerDict = lookupDict(key)
            Else
                ' 新しい品番・工程のキーの場合、内側のDictionaryを作成して追加
                Set innerDict = CreateObject("Scripting.Dictionary")
                lookupDict.Add key, innerDict
            End If

            ' 内側のDictionaryに日付とロット数量を追加/更新
            If innerDict.Exists(currentLotDateNum) Then
                ' 既に同じ日付が存在する場合、ロット数量をパイプで連結
                innerDict(currentLotDateNum) = innerDict(currentLotDateNum) & "|" & CStr(ロット数量値)
            Else
                ' 新しい日付の場合、追加
                innerDict.Add currentLotDateNum, CStr(ロット数量値)
            End If
        End If
SkipLotRow:
    Next i
    If IsArray(lotData) Then Erase lotData ' メモリ解放 (配列の場合のみ)

    ' ============================================
    ' 調査テーブルのデータを配列に読み込み
    ' ============================================
    Application.StatusBar = "「" & テーブル名 & "」の処理を開始します..."

    ' 列インデックスを取得（エラーチェック付き）
    Err.Clear
    On Error Resume Next
    品番調査Col = lo調査.ListColumns("品番").Index
    工程調査Col = lo調査.ListColumns("工程").Index
    日付調査Col = lo調査.ListColumns("日付").Index
    ロット数量調査Col = lo調査.ListColumns("ロット数量").Index
    If Err.Number <> 0 Then
        Application.StatusBar = "エラー: 「" & テーブル名 & "」に必要な列（品番, 工程, 日付, ロット数量）が見つかりません。"
        On Error GoTo 0
        GoTo CleanExit
    End If
    On Error GoTo 0

    ' データ範囲を配列に読み込み
    Dim surveyRowCount As Long
    If lo調査.ListRows.Count > 0 Then
        ' .Value を使って表示されている文字列を取得（特に日付列）
        surveyData = lo調査.DataBodyRange.Value
        surveyRowCount = lo調査.ListRows.Count ' 行数を取得
        全行数 = surveyRowCount
    Else
        全行数 = 0
        surveyData = Array()
    End If

    If 全行数 = 0 Then
        Application.StatusBar = "「" & テーブル名 & "」には処理対象データがありません。"
         GoTo CleanExit
    End If

    ' 結果を格納する配列を準備 (N行1列)
    ReDim resultData(1 To 全行数, 1 To 1)

    Application.StatusBar = "「" & テーブル名 & "」の処理を開始します。全 " & 全行数 & " 行..."

    ' ============================================
    ' メインループ：調査テーブルの各行に対してロット数量を検索
    ' ============================================
    For i = 1 To 全行数

        ' 進捗状況をステータスバーに表示 (100行ごと、最初、最後)
        If i Mod 100 = 0 Or i = 1 Or i = 全行数 Then
            進捗率 = Format(i / 全行数, "0%")
            Application.StatusBar = "「" & テーブル名 & "」処理中... " & i & "/" & 全行数 & " 行 (" & 進捗率 & ")"
            DoEvents ' 応答性を維持
        End If

        ' 配列から値を取得 (1行の場合も考慮)
        If 全行数 = 1 Then
            ' データが1行の場合、surveyDataが1次元配列か2次元配列(1, N)か判定
             If IsArray(surveyData) Then
                Dim isOneDimSurvey As Boolean
                isOneDimSurvey = False
                On Error Resume Next
                Dim checkDimSurvey As Long
                checkDimSurvey = UBound(surveyData, 2) ' 2番目の次元の上限を取得試行
                If Err.Number <> 0 Then
                    isOneDimSurvey = True ' エラー -> 1次元配列
                    Err.Clear
                End If
                On Error GoTo 0

                If isOneDimSurvey Then
                    ' 1次元配列の場合
                    If UBound(surveyData) >= 品番調査Col Then 品番 = surveyData(品番調査Col) Else 品番 = Empty
                    If UBound(surveyData) >= 工程調査Col Then 工程 = surveyData(工程調査Col) Else 工程 = Empty
                    If UBound(surveyData) >= 日付調査Col Then 日付 = surveyData(日付調査Col) Else 日付 = Empty
                Else
                    ' 2次元配列(1行)の場合
                    品番 = surveyData(1, 品番調査Col)
                    工程 = surveyData(1, 工程調査Col)
                    日付 = surveyData(1, 日付調査Col) ' .Value なので表示されている文字列のはず
                End If
            Else
                 品番 = Empty
                 工程 = Empty
                 日付 = Empty
            End If
        Else ' 全行数 > 1 (通常の2次元配列)
             品番 = surveyData(i, 品番調査Col)
             工程 = surveyData(i, 工程調査Col)
             日付 = surveyData(i, 日付調査Col) ' .Value なので表示されている文字列のはず
        End If

        resultLot数量 = "" ' 結果を初期化

        ' 品番、工程、日付が有効かチェック
        If Not IsEmpty(品番) And Not IsEmpty(工程) And Not IsEmpty(日付) Then
            key = CStr(品番) & vbTab & CStr(工程)

            ' Dictionaryに品番・工程のキーが存在するか確認
            If lookupDict.Exists(key) Then
                Set innerDict = lookupDict(key) ' 対応する内側Dictionaryを取得

                ' パイプ区切りで複数の日付を分割
                arr日付 = Split(CStr(日付), "|")

                ' 分割した各日付について処理
                For j = LBound(arr日付) To UBound(arr日付)
                    currentDateStr = Trim(arr日付(j))

                    ' 日付文字列を数値(Double)に変換試行し、整数部分を取得
                    On Error Resume Next ' 日付変換エラーを捕捉
                    currentDateNum = 0 ' 初期化
                    ' CDateで日付として認識させ、CDblでシリアル値にし、Intで整数部分（日付のみ）を取得
                    currentDateNum = Int(CDbl(CDate(currentDateStr)))
                    If Err.Number <> 0 Then
                        ' 日付変換エラーの場合
                        errorCount = errorCount + 1
                        Err.Clear
                        ' この日付はスキップ（ロット数量は見つからない）
                    Else
                        ' 日付変換成功、内側Dictionaryで日付キーを検索
                        If innerDict.Exists(currentDateNum) Then
                            ロット数量値 = innerDict(currentDateNum)
                            ' 結果文字列にロット数量を連結（初回か追記かで分岐）
                            If resultLot数量 = "" Then
                                resultLot数量 = CStr(ロット数量値)
                            Else
                                resultLot数量 = resultLot数量 & "|" & CStr(ロット数量値)
                            End If
                        End If
                    End If
                    On Error GoTo 0 ' エラーハンドリングを元に戻す
                Next j
            End If ' lookupDict.Exists(key)
        End If ' Not IsEmpty(品番) And ...

        ' 結果配列に格納
        resultData(i, 1) = resultLot数量

    Next i

    ' ============================================
    ' 結果を調査テーブルの「ロット数量」列に一括書き込み
    ' ============================================
    If 全行数 > 0 Then
        Application.StatusBar = "結果を書き込んでいます..."
        ' 書き込み前に列の保護などを解除する必要がある場合がある
        On Error Resume Next ' 書き込みエラーが発生する可能性（シート保護など）
        lo調査.ListColumns("ロット数量").DataBodyRange.Value = resultData
        If Err.Number <> 0 Then
             Application.StatusBar = "エラー: 結果の書き込みに失敗しました。シートが保護されているか、他の問題が発生した可能性があります。"
             On Error GoTo 0
             GoTo CleanExit ' エラー発生時は終了処理へ
        End If
        On Error GoTo 0
    End If

    ' ============================================
    ' 処理完了メッセージ表示（ステータスバー）
    ' ============================================
    Dim endTime As Double
    endTime = Timer
    Dim elapsedTime As Double
    elapsedTime = endTime - startTime

    If errorCount > 0 Then
        Application.StatusBar = "「" & テーブル名 & "」の処理が完了しました。" & 全行数 & "行の処理完了。（" & errorCount & "件の日付変換エラー）処理時間: " & Format(elapsedTime, "0.00") & "秒"
    Else
        Application.StatusBar = "「" & テーブル名 & "」の処理が完了しました。" & 全行数 & "行の処理が完了。処理時間: " & Format(elapsedTime, "0.00") & "秒"
    End If
    ' メッセージ表示のために少し待機
    Dim waitTime As Date
    waitTime = Now + TimeValue("00:00:01")
    Do While Now < waitTime
        DoEvents
    Loop

CleanExit: ' 終了処理共通部
    ' ============================================
    ' アプリケーション設定を元に戻す
    ' ============================================
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic ' 自動計算モードに戻す
    Application.ScreenUpdating = True
    Application.EnableEvents = True ' イベント発生を元に戻す

    ' オブジェクト変数の解放
    Set lo調査 = Nothing
    Set loロット数量 = Nothing
    Set lookupDict = Nothing
    Set innerDict = Nothing
    ' 配列変数は自動的に解放されるが、明示的に解放する場合
    ' If IsArray(surveyData) Then Erase surveyData
    ' If IsArray(resultData) Then Erase resultData

End Sub

' ============================================
' 関数: GetListObjectByName
' 機能: 指定のテーブル名を持つListObjectをブック内全シートから検索
' 引数: targetTableName - 検索対象のテーブル名
' 戻り値: ListObject（見つかった場合）、Nothing（見つからない場合）
' ============================================
Function GetListObjectByName(targetTableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    ' ブック内の全シートを順に検索
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next ' 特定のシートでListObjectが見つからない場合のエラーを無視
        Set lo = ws.ListObjects(targetTableName)
        On Error GoTo 0 ' エラーハンドリングをリセット

        If Not lo Is Nothing Then
            ' テーブルが見つかった場合、関数の戻り値として返して終了
            Set GetListObjectByName = lo
            Exit Function
        End If
    Next ws

    ' 見つからなかった場合はNothingを返す
    Set GetListObjectByName = Nothing
End Function

' ============================================
' 以下、調査テーブル番号ごとの呼び出し用サブルーチン
' 各サブルーチンは対応する番号でロット数量調査_高速化を呼び出す
' ============================================

' 調査1：削除済みテーブル（情報表示のみ）
Sub ロット数量調査1()
    MsgBox "テーブル「_調査1」は削除されました。", vbInformation, "情報"
End Sub

' 調査2：削除済みテーブル（情報表示のみ）
Sub ロット数量調査2()
    MsgBox "テーブル「_調査2」は削除されました。", vbInformation, "情報"
End Sub

' 調査3：セット品シートの_調査3テーブルを処理
Sub ロット数量調査3()
    Call ロット数量調査_高速化(3)
End Sub

' 調査4：単品シートの_調査4テーブルを処理
Sub ロット数量調査4()
    Call ロット数量調査_高速化(4)
End Sub

' 調査5：（定義のみ、テーブル番号5での処理）
Sub ロット数量調査5()
    Call ロット数量調査_高速化(5)
End Sub
