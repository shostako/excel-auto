Attribute VB_Name = "m転記_日報_加工NW"
Option Explicit

' ========================================
' マクロ名: 転記_日報_加工NW
' 処理概要: 日報データを期間別に9分類で集計して加工シートに転記（ワースト順対応版）
'
' 【処理の特徴】
' 1. 空白期間スキップ：集計期間テーブルに行があっても、該当期間内にデータがなければテーブルを作らない
' 2. 動的期間対応：集計期間テーブルの行数が変わっても自動的に対応（増減どちらもOK）
' 3. 高速化：配列処理による大量データの高速集計
' 4. ワースト順機能：項目テーブルの「ワースト」設定に応じて動的に出力順序を変更
'
' 【テーブル構成】
' 期間テーブル : シート「加工NW」、テーブル「_集計期間日報加工W」
' ソーステーブル : シート「日報加工」、テーブル「_日報加工」
' 項目テーブル : シート「加工NW」、テーブル「_日報項目加工W」
' 出力テーブル : シート「加工NW」、複数テーブル「_日報W_加工_{期間名}」
'
' 【処理フロー】
' 1. 既存出力テーブルとデータを完全削除
' 2. ワースト設定（全項目 or 数値N）を読み込み
' 3. 各期間ごとに日付フィルター + 品番による9分類集計
' 4. 集計結果を降順ソートしてワースト順出力
' 5. データがある期間のみテーブル出力（空白期間はスキップ）
'
' 【出力形式】
' - 1行目：ショット数
' - 2行目：不良数（各項目の合計値：不良がゼロでも表示）
' - 3行目以降：ワースト順で項目別集計
'   - 「全項目」設定：0でない項目を降順で全て出力
'   - 数値N設定：上位N件 + 「その他」行（N+1行、ただし0でない項目数<=Nなら「その他」なし）
' - 最終列：合計
'
' 【集計方法（加工特有・2パターン）】
' ■パターンA（62-xxxxx形式）: LH/RH両方にカウント（×1）
'   - 62-58040Fr, 62-58050Fr, 62-58060Fr → 58050FrLH/RH両方
'   - 62-58040Rr, 62-58050Rr, 62-58060Rr → 58050RrLH/RH両方
'   - 62-28030Fr〜62-28060Fr → 28050FrLH/RH両方
'   - 62-28030Rr〜62-28060Rr → 28050RrLH/RH両方
' ■パターンB（ロット=「単」かつ品番に特定数字含む）: 特定1グループのみ（×2）
'   - 58050FrLH: 品番に{58042,58052,58062}含む
'   - 58050FrRH: 品番に{58041,58051,58061}含む
'   - 58050RrLH: 品番に{58056,58066}含む
'   - 58050RrRH: 品番に{58055,58065}含む
'   - 28050FrLH: 品番に{28032,28042,28052,28062}含む
'   - 28050FrRH: 品番に{28031,28041,28051,28061}含む
'   - 28050RrLH: 品番に{28036,28046,28056,28066}含む
'   - 28050RrRH: 品番に{28035,28045,28055,28065}含む
' ■補給品（上記以外）:
'   - ロット=「単」 → ×2
'   - ロット≠「単」 → 末尾LH/RHあり×1、なし×2
' ========================================

Sub 転記_日報_加工NW()
    ' ============================================
    ' 最適化設定の保存と適用
    ' 理由：画面更新・再計算・イベントを止めて処理を高速化
    ' ============================================
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    Dim origEnableEvents As Boolean
    Dim origDisplayAlerts As Boolean

    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation
    origEnableEvents = Application.EnableEvents
    origDisplayAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    On Error GoTo ErrorHandler
    Application.StatusBar = "日報加工W転記処理を開始します..."

    ' ============================================
    ' シートとテーブルの参照取得
    ' 理由：オブジェクト参照で直接操作するため（Activateは使わない）
    ' ============================================
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Set wsSource = ThisWorkbook.Worksheets("日報加工")
    Set wsTarget = ThisWorkbook.Worksheets("加工NW")

    ' テーブル参照を取得（存在チェックはOn Error Resume Nextで安全に）
    Dim tblSource As ListObject, tblItems As ListObject, tblPeriod As ListObject
    On Error Resume Next
    Set tblSource = wsSource.ListObjects("_日報加工")
    Set tblItems = wsTarget.ListObjects("_日報項目加工W")
    Set tblPeriod = wsTarget.ListObjects("_集計期間日報加工W")
    On Error GoTo ErrorHandler

    ' ソーステーブルは必須
    If tblSource Is Nothing Then
        MsgBox "シート「日報加工」にテーブル「_日報加工」が見つかりません。", vbCritical
        GoTo Cleanup
    End If

    ' ============================================
    ' ワースト設定の読み込み
    ' 理由：項目テーブルの「ワースト」列から「全項目」or数値Nを取得
    ' ポイント：テーブルが空またはNothingでもエラーにならない
    ' ============================================
    Dim worstSetting As String
    Dim worstNum As Long
    Dim isAllItems As Boolean

    worstSetting = ""
    worstNum = 0
    isAllItems = False

    If Not tblItems Is Nothing Then
        If Not tblItems.DataBodyRange Is Nothing Then
            ' ワースト列の1行目を取得
            Dim colWorstIdx As Long
            On Error Resume Next
            colWorstIdx = tblItems.ListColumns("ワースト").Index
            On Error GoTo ErrorHandler

            If colWorstIdx > 0 Then
                worstSetting = Trim(CStr(tblItems.DataBodyRange.Cells(1, colWorstIdx).Value))

                If worstSetting = "全項目" Then
                    isAllItems = True
                ElseIf IsNumeric(worstSetting) Then
                    worstNum = CLng(worstSetting)
                    If worstNum <= 0 Then
                        MsgBox "「ワースト」列の数値は1以上を指定してください。", vbCritical
                        GoTo Cleanup
                    End If
                Else
                    MsgBox "「ワースト」列の値が不正です。「全項目」または数値を指定してください。", vbCritical
                    GoTo Cleanup
                End If
            End If
        End If
    End If

    ' ワースト設定が取得できない場合はデフォルト（全項目）
    If worstSetting = "" Then
        isAllItems = True
        Application.StatusBar = "ワースト設定が見つかりません。全項目モードで実行します..."
    End If

    ' ============================================
    ' 期間テーブルの読み込み
    ' 理由：Nothing対策を多段階で実施し、安全に期間情報を配列化
    ' ポイント：テーブルが空でもエラーにならないよう慎重にチェック
    ' ============================================
    Dim periodCount As Long
    periodCount = 0
    Dim periodInfo() As Variant

    If Not tblPeriod Is Nothing Then
        If Not tblPeriod.DataBodyRange Is Nothing Then
            periodCount = tblPeriod.DataBodyRange.Rows.Count
            If periodCount > 0 Then
                ReDim periodInfo(1 To periodCount, 1 To 3)
                Dim p As Long
                For p = 1 To periodCount
                    periodInfo(p, 1) = CStr(tblPeriod.DataBodyRange.Cells(p, 1).Value) ' 期間名
                    periodInfo(p, 2) = tblPeriod.DataBodyRange.Cells(p, 2).Value       ' 開始日
                    periodInfo(p, 3) = tblPeriod.DataBodyRange.Cells(p, 3).Value       ' 終了日
                Next p
            End If
        End If
    End If

    ' 集計期間が1つもなければ処理中止
    If periodCount = 0 Then
        MsgBox "「_集計期間日報加工W」に有効な集計期間がありません。処理を中止します。", vbExclamation
        GoTo Cleanup
    End If

    ' ============================================
    ' ソーステーブルのデータ範囲取得
    ' 理由：後で配列化して高速処理するための準備
    ' ============================================
    Dim srcData As Range
    Set srcData = tblSource.DataBodyRange
    If srcData Is Nothing Then
        Application.StatusBar = "ソーステーブルにデータがありません"
        GoTo Cleanup
    End If

    ' ============================================
    ' 列インデックスの取得
    ' 理由：テーブル内の相対位置を事前に取得して配列アクセスに使用
    ' ============================================
    Dim colHizuke As Long, colHinban As Long, colLot As Long
    colHizuke = tblSource.ListColumns("日付").Index
    colHinban = tblSource.ListColumns("品番").Index
    colLot = tblSource.ListColumns("ロット").Index

    ' 集計対象列のインデックスをDictionaryで管理
    ' 理由：列が存在しない場合でもエラーにならず、柔軟に対応
    Dim colIndexes As Object
    Set colIndexes = CreateObject("Scripting.Dictionary")

    Dim targetColumns As Variant
    targetColumns = Array("ショット数", "不良数", "成形不良", "プライマー付着", "テープ貼り失敗", _
                          "テープ蛇行", "テープ内異物", "キズ付け", "その他")

    ' ワースト順集計対象項目（7項目）
    Dim worstTargetItems As Variant
    worstTargetItems = Array("成形不良", "プライマー付着", "テープ貼り失敗", _
                             "テープ蛇行", "テープ内異物", "キズ付け", "その他")

    Dim colName As Variant
    Dim colIdx As Long
    On Error Resume Next
    For Each colName In targetColumns
        colIdx = 0
        colIdx = tblSource.ListColumns(CStr(colName)).Index
        If Err.Number = 0 And colIdx > 0 Then
            colIndexes(CStr(colName)) = colIdx
        End If
        Err.Clear
    Next colName
    On Error GoTo ErrorHandler

    ' ============================================
    ' 品番パターンの定義（加工特有・2パターン）
    ' 理由：品番で8グループ+補給品に振り分け
    ' ============================================

    ' ■パターンA（62-xxxxx形式）→ LH/RH両方にカウント
    ' 58050Fr系パターン → 58050FrLH/RH両方
    Dim patternA_58050Fr As Object
    Set patternA_58050Fr = CreateObject("Scripting.Dictionary")
    patternA_58050Fr("62-58040Fr") = True
    patternA_58050Fr("62-58050Fr") = True
    patternA_58050Fr("62-58060Fr") = True

    ' 58050Rr系パターン → 58050RrLH/RH両方
    Dim patternA_58050Rr As Object
    Set patternA_58050Rr = CreateObject("Scripting.Dictionary")
    patternA_58050Rr("62-58040Rr") = True
    patternA_58050Rr("62-58050Rr") = True
    patternA_58050Rr("62-58060Rr") = True

    ' 28050Fr系パターン → 28050FrLH/RH両方
    Dim patternA_28050Fr As Object
    Set patternA_28050Fr = CreateObject("Scripting.Dictionary")
    patternA_28050Fr("62-28030Fr") = True
    patternA_28050Fr("62-28040Fr") = True
    patternA_28050Fr("62-28050Fr") = True
    patternA_28050Fr("62-28060Fr") = True

    ' 28050Rr系パターン → 28050RrLH/RH両方
    Dim patternA_28050Rr As Object
    Set patternA_28050Rr = CreateObject("Scripting.Dictionary")
    patternA_28050Rr("62-28030Rr") = True
    patternA_28050Rr("62-28040Rr") = True
    patternA_28050Rr("62-28050Rr") = True
    patternA_28050Rr("62-28060Rr") = True

    ' ■パターンB（ロット=「単」かつ品番に特定数字含む）→ 特定1グループのみ
    ' 品番に含まれるべき数字リスト（部分一致判定用）
    Dim patternB_58050FrLH As Variant
    patternB_58050FrLH = Array("58042", "58052", "58062")

    Dim patternB_58050FrRH As Variant
    patternB_58050FrRH = Array("58041", "58051", "58061")

    Dim patternB_58050RrLH As Variant
    patternB_58050RrLH = Array("58056", "58066")

    Dim patternB_58050RrRH As Variant
    patternB_58050RrRH = Array("58055", "58065")

    Dim patternB_28050FrLH As Variant
    patternB_28050FrLH = Array("28032", "28042", "28052", "28062")

    Dim patternB_28050FrRH As Variant
    patternB_28050FrRH = Array("28031", "28041", "28051", "28061")

    Dim patternB_28050RrLH As Variant
    patternB_28050RrLH = Array("28036", "28046", "28056", "28066")

    Dim patternB_28050RrRH As Variant
    patternB_28050RrRH = Array("28035", "28045", "28055", "28065")

    ' ============================================
    ' 既存の出力テーブルオブジェクトを削除
    ' 理由：期間数が減った場合、古いテーブルが残るとエラーになる
    ' ポイント：逆順でループすることで削除中のインデックスずれを防止
    ' ============================================
    Dim idxLO As Long
    For idxLO = wsTarget.ListObjects.Count To 1 Step -1
        Dim loTemp As ListObject
        Set loTemp = wsTarget.ListObjects(idxLO)
        If loTemp.Name Like "_日報W_加工_*" Then
            loTemp.Delete  ' 直接削除（名前での再参照は不要）
        End If
    Next idxLO

    ' ============================================
    ' 既存出力範囲の行削除
    ' 理由：テーブルオブジェクト削除後もセルの値は残るため、
    '       参照テーブルより下の行を全削除してクリーンアップ
    ' ============================================
    Dim itemsTableLastRow As Long, periodTableLastRow As Long
    itemsTableLastRow = 0
    If Not tblItems Is Nothing Then
        itemsTableLastRow = tblItems.Range.Row + tblItems.Range.Rows.Count - 1
    End If

    periodTableLastRow = 0
    If Not tblPeriod Is Nothing Then
        periodTableLastRow = tblPeriod.Range.Row + tblPeriod.Range.Rows.Count - 1
    End If

    ' 2つのテーブルで下にある方を基準行とする
    Dim baseRow As Long
    If itemsTableLastRow > periodTableLastRow Then
        baseRow = itemsTableLastRow
    Else
        baseRow = periodTableLastRow
    End If
    If baseRow < 1 Then baseRow = 1

    ' 基準行より下を全削除
    Dim lastUsedRow As Long
    lastUsedRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
    If lastUsedRow >= baseRow + 1 Then
        wsTarget.Rows((baseRow + 1) & ":" & lastUsedRow).Delete
    End If

    ' ============================================
    ' 出力開始位置の決定
    ' 理由：参照テーブルの下に2行空けてから出力開始
    ' ============================================
    Dim currentRow As Long
    currentRow = baseRow + 3

    ' ============================================
    ' 全グループ配列の定義
    ' 理由：出力時の列順序を固定（LH→RHの順）
    ' ============================================
    Dim allGroups As Variant
    allGroups = Array("58050FrLH", "58050FrRH", "58050RrLH", "58050RrRH", _
                      "28050FrLH", "28050FrRH", "28050RrLH", "28050RrRH", "補給品")

    ' ============================================
    ' ソースデータを配列に取り込み
    ' 理由：Range.Cellsへの繰り返しアクセスは遅いため、
    '       一度配列化することで大幅に高速化
    ' ポイント：配列は1-based (rows, cols)
    ' ============================================
    Dim srcArr As Variant
    srcArr = srcData.Value

    ' ============================================
    ' 印刷範囲の記録用変数
    ' 理由：出力したテーブル全体を印刷範囲に設定するため、
    '       最初のテーブルの開始位置と最後のテーブルの終了位置を記録
    ' ============================================
    Dim printRangeStart As Long
    Dim printRangeEnd As Long
    printRangeStart = 0  ' 0なら未設定（データが1つもない場合）
    printRangeEnd = 0

    ' ============================================
    ' 各期間の処理ループ
    ' ============================================
    Dim periodIdx As Long
    For periodIdx = 1 To periodCount
        Application.StatusBar = "期間 " & periodIdx & "/" & periodCount & " を処理中..."

        ' 期間情報の取り出し
        Dim periodName As String, startDate As Date, endDate As Date
        periodName = CStr(periodInfo(periodIdx, 1))
        startDate = CDate(periodInfo(periodIdx, 2))
        endDate = CDate(periodInfo(periodIdx, 3))

        ' ============================================
        ' グループ別集計用Dictionaryの初期化
        ' 理由：各期間ごとに集計をゼロからやり直すため
        ' 構造：
        '   aggShot: グループ名 → ショット数合計
        '   aggFuryo: グループ名 → 不良数合計
        '   aggItems: グループ名 → Dictionary(項目名 → 合計値)
        ' ============================================
        Dim aggShot As Object, aggFuryo As Object, aggItems As Object
        Set aggShot = CreateObject("Scripting.Dictionary")
        Set aggFuryo = CreateObject("Scripting.Dictionary")
        Set aggItems = CreateObject("Scripting.Dictionary")

        Dim grp As Variant
        For Each grp In allGroups
            aggShot(CStr(grp)) = 0
            aggFuryo(CStr(grp)) = 0

            ' 項目別集計用のネストDictionary（7項目）
            Set aggItems(CStr(grp)) = CreateObject("Scripting.Dictionary")
            Dim wItem As Variant
            For Each wItem In worstTargetItems
                aggItems(CStr(grp))(CStr(wItem)) = 0
            Next wItem
        Next grp

        ' ============================================
        ' 空白期間判定フラグ
        ' 理由：この期間内に実際のデータ（空白でない値）が1つでもあるか
        '       を判定し、完全に空白ならテーブルを作らない
        ' ============================================
        Dim hasData As Boolean
        hasData = False

        ' ============================================
        ' ソース配列の走査と集計
        ' 理由：日付フィルターで該当期間のデータのみを集計
        ' ============================================
        Dim r As Long
        Dim totalRows As Long
        totalRows = UBound(srcArr, 1)

        For r = 1 To totalRows
            Dim cellDate As Variant
            cellDate = srcArr(r, colHizuke)

            ' 日付フィルタ
            If IsDate(cellDate) Then
                Dim dt As Date
                dt = CDate(cellDate)

                If dt >= startDate And dt <= endDate Then
                    ' 品番列の値で9分類を判定
                    Dim hinbanVal As String
                    hinbanVal = Trim(CStr(srcArr(r, colHinban)))

                    ' ロット列の値を取得
                    Dim lotVal As String
                    lotVal = Trim(CStr(srcArr(r, colLot)))

                    ' 振り分け先グループを決定（パターンA/B/補給品）
                    Dim targetGroups() As String
                    Dim targetGroupCount As Long
                    Dim multiplier As Double
                    Dim matched As Boolean
                    Dim pIdx As Long

                    targetGroupCount = 0
                    multiplier = 1
                    matched = False

                    ' ■パターンA判定（62-xxxxx形式）→ LH/RH両方にカウント
                    If patternA_58050Fr.Exists(hinbanVal) Then
                        ' 58050Fr系 → FrLH/RH両方
                        ReDim targetGroups(1 To 2)
                        targetGroups(1) = "58050FrLH"
                        targetGroups(2) = "58050FrRH"
                        targetGroupCount = 2
                        matched = True
                    ElseIf patternA_58050Rr.Exists(hinbanVal) Then
                        ' 58050Rr系 → RrLH/RH両方
                        ReDim targetGroups(1 To 2)
                        targetGroups(1) = "58050RrLH"
                        targetGroups(2) = "58050RrRH"
                        targetGroupCount = 2
                        matched = True
                    ElseIf patternA_28050Fr.Exists(hinbanVal) Then
                        ' 28050Fr系 → FrLH/RH両方
                        ReDim targetGroups(1 To 2)
                        targetGroups(1) = "28050FrLH"
                        targetGroups(2) = "28050FrRH"
                        targetGroupCount = 2
                        matched = True
                    ElseIf patternA_28050Rr.Exists(hinbanVal) Then
                        ' 28050Rr系 → RrLH/RH両方
                        ReDim targetGroups(1 To 2)
                        targetGroups(1) = "28050RrLH"
                        targetGroups(2) = "28050RrRH"
                        targetGroupCount = 2
                        matched = True
                    End If

                    ' ■パターンB判定（ロット=「単」かつ品番に特定数字含む）→ 特定1グループのみ、×2
                    If Not matched And lotVal = "単" Then
                        ' 58050FrLH: 品番に{58042,58052,58062}含む
                        For pIdx = LBound(patternB_58050FrLH) To UBound(patternB_58050FrLH)
                            If InStr(hinbanVal, CStr(patternB_58050FrLH(pIdx))) > 0 Then
                                ReDim targetGroups(1 To 1)
                                targetGroups(1) = "58050FrLH"
                                targetGroupCount = 1
                                multiplier = 2  ' パターンBは×2
                                matched = True
                                Exit For
                            End If
                        Next pIdx

                        ' 58050FrRH: 品番に{58041,58051,58061}含む
                        If Not matched Then
                            For pIdx = LBound(patternB_58050FrRH) To UBound(patternB_58050FrRH)
                                If InStr(hinbanVal, CStr(patternB_58050FrRH(pIdx))) > 0 Then
                                    ReDim targetGroups(1 To 1)
                                    targetGroups(1) = "58050FrRH"
                                    targetGroupCount = 1
                                    multiplier = 2  ' パターンBは×2
                                    matched = True
                                    Exit For
                                End If
                            Next pIdx
                        End If

                        ' 58050RrLH: 品番に{58056,58066}含む
                        If Not matched Then
                            For pIdx = LBound(patternB_58050RrLH) To UBound(patternB_58050RrLH)
                                If InStr(hinbanVal, CStr(patternB_58050RrLH(pIdx))) > 0 Then
                                    ReDim targetGroups(1 To 1)
                                    targetGroups(1) = "58050RrLH"
                                    targetGroupCount = 1
                                    multiplier = 2  ' パターンBは×2
                                    matched = True
                                    Exit For
                                End If
                            Next pIdx
                        End If

                        ' 58050RrRH: 品番に{58055,58065}含む
                        If Not matched Then
                            For pIdx = LBound(patternB_58050RrRH) To UBound(patternB_58050RrRH)
                                If InStr(hinbanVal, CStr(patternB_58050RrRH(pIdx))) > 0 Then
                                    ReDim targetGroups(1 To 1)
                                    targetGroups(1) = "58050RrRH"
                                    targetGroupCount = 1
                                    multiplier = 2  ' パターンBは×2
                                    matched = True
                                    Exit For
                                End If
                            Next pIdx
                        End If

                        ' 28050FrLH: 品番に{28032,28042,28052,28062}含む
                        If Not matched Then
                            For pIdx = LBound(patternB_28050FrLH) To UBound(patternB_28050FrLH)
                                If InStr(hinbanVal, CStr(patternB_28050FrLH(pIdx))) > 0 Then
                                    ReDim targetGroups(1 To 1)
                                    targetGroups(1) = "28050FrLH"
                                    targetGroupCount = 1
                                    multiplier = 2  ' パターンBは×2
                                    matched = True
                                    Exit For
                                End If
                            Next pIdx
                        End If

                        ' 28050FrRH: 品番に{28031,28041,28051,28061}含む
                        If Not matched Then
                            For pIdx = LBound(patternB_28050FrRH) To UBound(patternB_28050FrRH)
                                If InStr(hinbanVal, CStr(patternB_28050FrRH(pIdx))) > 0 Then
                                    ReDim targetGroups(1 To 1)
                                    targetGroups(1) = "28050FrRH"
                                    targetGroupCount = 1
                                    multiplier = 2  ' パターンBは×2
                                    matched = True
                                    Exit For
                                End If
                            Next pIdx
                        End If

                        ' 28050RrLH: 品番に{28036,28046,28056,28066}含む
                        If Not matched Then
                            For pIdx = LBound(patternB_28050RrLH) To UBound(patternB_28050RrLH)
                                If InStr(hinbanVal, CStr(patternB_28050RrLH(pIdx))) > 0 Then
                                    ReDim targetGroups(1 To 1)
                                    targetGroups(1) = "28050RrLH"
                                    targetGroupCount = 1
                                    multiplier = 2  ' パターンBは×2
                                    matched = True
                                    Exit For
                                End If
                            Next pIdx
                        End If

                        ' 28050RrRH: 品番に{28035,28045,28055,28065}含む
                        If Not matched Then
                            For pIdx = LBound(patternB_28050RrRH) To UBound(patternB_28050RrRH)
                                If InStr(hinbanVal, CStr(patternB_28050RrRH(pIdx))) > 0 Then
                                    ReDim targetGroups(1 To 1)
                                    targetGroups(1) = "28050RrRH"
                                    targetGroupCount = 1
                                    multiplier = 2  ' パターンBは×2
                                    matched = True
                                    Exit For
                                End If
                            Next pIdx
                        End If
                    End If

                    ' ■補給品判定（上記以外）
                    If Not matched Then
                        ReDim targetGroups(1 To 1)
                        targetGroups(1) = "補給品"
                        targetGroupCount = 1

                        If lotVal = "単" Then
                            ' ロット=「単」でパターンBに該当しない → ×2
                            multiplier = 2
                        Else
                            ' ロット≠「単」 → 末尾LH/RH判定で倍率を決定
                            If Right(hinbanVal, 2) = "LH" Or Right(hinbanVal, 2) = "RH" Then
                                multiplier = 1
                            Else
                                multiplier = 2
                            End If
                        End If
                    End If

                    ' 各列の値を集計
                    Dim keyName As Variant
                    For Each keyName In colIndexes.Keys
                        Dim colIdxSrc As Long
                        colIdxSrc = colIndexes(keyName)

                        If colIdxSrc >= 1 And colIdxSrc <= UBound(srcArr, 2) Then
                            Dim colValue As Variant
                            colValue = srcArr(r, colIdxSrc)

                            ' 空白チェック（空白でなければデータありと判定）
                            If Not IsError(colValue) Then
                                If Len(Trim(CStr(colValue))) > 0 Then
                                    hasData = True
                                End If
                            End If

                            ' 数値なら集計に加算
                            If IsNumeric(colValue) Then
                                Dim numValue As Double
                                numValue = CDbl(colValue) * multiplier

                                ' 各ターゲットグループに加算
                                Dim tgIdx As Long
                                For tgIdx = 1 To targetGroupCount
                                    Dim tg As String
                                    tg = targetGroups(tgIdx)

                                    ' 列名による振り分け
                                    If CStr(keyName) = "ショット数" Then
                                        aggShot(tg) = aggShot(tg) + numValue
                                    ElseIf CStr(keyName) = "不良数" Then
                                        aggFuryo(tg) = aggFuryo(tg) + numValue
                                    Else
                                        ' 7項目のいずれか
                                        Dim isWorstItem As Boolean
                                        isWorstItem = False
                                        Dim checkItem As Variant
                                        For Each checkItem In worstTargetItems
                                            If CStr(keyName) = CStr(checkItem) Then
                                                isWorstItem = True
                                                Exit For
                                            End If
                                        Next checkItem

                                        If isWorstItem Then
                                            aggItems(tg)(CStr(keyName)) = aggItems(tg)(CStr(keyName)) + numValue
                                        End If
                                    End If
                                Next tgIdx
                            End If
                        End If
                    Next keyName
                End If
            End If

            ' 進捗表示（200行ごと）
            If (r Mod 200) = 0 Then
                Application.StatusBar = "期間 " & periodIdx & "/" & periodCount & " - " & r & "/" & totalRows & " 行処理中..."
            End If
        Next r

        ' ============================================
        ' 空白期間スキップ処理
        ' 理由：データが1つもない期間はテーブルを作らない
        '       （出力位置を進めずに次の期間へ）
        ' ============================================
        If Not hasData Then
            Application.StatusBar = "期間 " & periodIdx & " はデータ無しのためスキップします..."
            GoTo NextPeriod
        End If

        ' ============================================
        ' 印刷範囲の開始位置を記録（最初のテーブルのみ）
        ' 理由：複数テーブルの最初のタイトル行を記録
        ' ============================================
        If printRangeStart = 0 Then
            printRangeStart = currentRow  ' 最初のタイトル行
        End If

        ' ============================================
        ' テーブル出力：タイトル行
        ' ============================================
        Dim titleText As String
        titleText = "日報W_加工_" & periodName & "_" & Format(startDate, "m/d") & "‾" & Format(endDate, "m/d")

        With wsTarget.Cells(currentRow, 1)
            .Value = titleText
            .ShrinkToFit = False
            .WrapText = False
            .Font.Bold = True
            .Font.Size = 12
        End With

        ' ============================================
        ' テーブル出力：ヘッダー行
        ' ============================================
        Dim outputStartRow As Long
        outputStartRow = currentRow + 1

        wsTarget.Cells(outputStartRow, 1).Value = "項目"

        Dim colOffset As Long
        colOffset = 2
        For Each grp In allGroups
            With wsTarget.Cells(outputStartRow, colOffset)
                .Value = CStr(grp)
                .ShrinkToFit = True
            End With
            colOffset = colOffset + 1
        Next grp

        ' 合計列のヘッダー
        With wsTarget.Cells(outputStartRow, colOffset)
            .Value = "合計"
            .ShrinkToFit = True
        End With

        ' ============================================
        ' テーブル出力：データ行
        ' ============================================
        Dim dataStartRow As Long
        dataStartRow = outputStartRow + 1
        Dim rowIdx As Long
        rowIdx = dataStartRow

        ' 1行目：ショット数
        With wsTarget.Cells(rowIdx, 1)
            .Value = "ショット数"
            .ShrinkToFit = True
        End With
        Dim rowTotal As Double
        rowTotal = 0
        colOffset = 2
        For Each grp In allGroups
            Dim cellValue As Double
            cellValue = aggShot(CStr(grp))
            wsTarget.Cells(rowIdx, colOffset).Value = cellValue
            rowTotal = rowTotal + cellValue
            colOffset = colOffset + 1
        Next grp
        ' 合計列
        wsTarget.Cells(rowIdx, colOffset).Value = rowTotal
        rowIdx = rowIdx + 1

        ' 2行目：不良数（各項目の合計）
        With wsTarget.Cells(rowIdx, 1)
            .Value = "不良数"
            .ShrinkToFit = True
        End With
        rowTotal = 0
        colOffset = 2
        For Each grp In allGroups
            cellValue = aggFuryo(CStr(grp))
            wsTarget.Cells(rowIdx, colOffset).Value = cellValue
            rowTotal = rowTotal + cellValue
            colOffset = colOffset + 1
        Next grp
        ' 合計列
        wsTarget.Cells(rowIdx, colOffset).Value = rowTotal
        rowIdx = rowIdx + 1

        ' ============================================
        ' 3行目以降：ワースト順で項目別集計
        ' 理由：全グループの合計値でワースト順を決定する（7項目）
        ' ============================================

        ' 全グループの項目別合計を計算（7項目）
        Dim totalItems As Object
        Set totalItems = CreateObject("Scripting.Dictionary")

        ' 7項目すべてを初期化
        Dim wItem2 As Variant
        For Each wItem2 In worstTargetItems
            totalItems(CStr(wItem2)) = 0
        Next wItem2

        ' 全グループの合計を計算
        For Each grp In allGroups
            Dim itemKey3 As Variant
            For Each itemKey3 In aggItems(CStr(grp)).Keys
                totalItems(CStr(itemKey3)) = totalItems(CStr(itemKey3)) + CDbl(aggItems(CStr(grp))(itemKey3))
            Next itemKey3
        Next grp

        ' 全グループ合計を配列化して降順ソート
        Dim totalArr() As Variant
        Dim totalCount As Long
        totalCount = totalItems.Count

        ReDim totalArr(1 To totalCount, 1 To 2)
        Dim idx As Long
        idx = 1
        Dim totalKey As Variant
        For Each totalKey In totalItems.Keys
            totalArr(idx, 1) = CStr(totalKey)  ' 項目名
            totalArr(idx, 2) = CDbl(totalItems(totalKey))  ' 全グループ合計値
            idx = idx + 1
        Next totalKey

        ' 降順ソート（全グループ合計値の大きい順）
        Call QuickSortDesc(totalArr, 1, totalCount)

        ' ============================================
        ' ワースト順出力の実行
        ' ============================================

        ' 出力する項目リストを作成（全グループ合計値のワースト順）
        Dim outputItemList() As String
        Dim outputItemCount As Long
        Dim hasSonotaRow As Boolean

        hasSonotaRow = False
        outputItemCount = 0

        ' 0でない項目だけをフィルタリング
        Dim nonZeroCount As Long
        nonZeroCount = 0
        Dim i2 As Long
        For i2 = 1 To UBound(totalArr, 1)
            If CDbl(totalArr(i2, 2)) <> 0 Then
                nonZeroCount = nonZeroCount + 1
            End If
        Next i2

        ' ワースト設定に応じて出力項目を決定
        If isAllItems Then
            ' 「全項目」モード：0でない項目を全て出力
            outputItemCount = nonZeroCount
            If outputItemCount > 0 Then
                ReDim outputItemList(1 To outputItemCount)
                Dim outIdx As Long
                outIdx = 1
                For i2 = 1 To UBound(totalArr, 1)
                    If CDbl(totalArr(i2, 2)) <> 0 Then
                        outputItemList(outIdx) = CStr(totalArr(i2, 1))
                        outIdx = outIdx + 1
                    End If
                Next i2
            End If
        Else
            ' 数値Nモード：上位N件 + その他
            If nonZeroCount > worstNum Then
                ' 0でない項目数 > N → 上位N件 + 「その他」
                outputItemCount = worstNum
                ReDim outputItemList(1 To outputItemCount)
                For i2 = 1 To worstNum
                    outputItemList(i2) = CStr(totalArr(i2, 1))
                Next i2
                hasSonotaRow = True  ' 必ず「その他」行を出力
            Else
                ' 0でない項目数 <= N → 0でない項目のみ
                outputItemCount = nonZeroCount
                If outputItemCount > 0 Then
                    ReDim outputItemList(1 To outputItemCount)
                    outIdx = 1
                    For i2 = 1 To UBound(totalArr, 1)
                        If CDbl(totalArr(i2, 2)) <> 0 Then
                            outputItemList(outIdx) = CStr(totalArr(i2, 1))
                            outIdx = outIdx + 1
                        End If
                    Next i2
                End If
            End If
        End If

        ' ============================================
        ' 項目行の出力（ワースト順）
        ' ============================================
        Dim outItem As Long
        For outItem = 1 To outputItemCount
            Dim currentItemName As String
            currentItemName = outputItemList(outItem)

            With wsTarget.Cells(rowIdx, 1)
                .Value = currentItemName
                .ShrinkToFit = True
            End With

            rowTotal = 0
            colOffset = 2
            For Each grp In allGroups
                ' このグループの集計値から該当項目の値を取得
                Dim itemValue As Double
                itemValue = 0

                If aggItems(CStr(grp)).Exists(currentItemName) Then
                    itemValue = CDbl(aggItems(CStr(grp))(currentItemName))
                End If

                wsTarget.Cells(rowIdx, colOffset).Value = itemValue
                rowTotal = rowTotal + itemValue
                colOffset = colOffset + 1
            Next grp

            ' 合計列
            wsTarget.Cells(rowIdx, colOffset).Value = rowTotal

            rowIdx = rowIdx + 1
        Next outItem

        ' ============================================
        ' 「その他」行の出力（必要な場合のみ）
        ' 理由：上位N件以外の項目の合計
        ' ============================================
        If hasSonotaRow Then
            With wsTarget.Cells(rowIdx, 1)
                .Value = "その他計"
                .ShrinkToFit = True
            End With

            rowTotal = 0
            colOffset = 2
            For Each grp In allGroups
                ' このグループの「その他」合計を計算
                Dim sonotaSum As Double
                sonotaSum = 0

                ' 上位N件以外を加算
                Dim k As Long
                For k = worstNum + 1 To UBound(totalArr, 1)
                    Dim sonotaItemName As String
                    sonotaItemName = CStr(totalArr(k, 1))

                    ' このグループにその項目があれば加算
                    If aggItems(CStr(grp)).Exists(sonotaItemName) Then
                        sonotaSum = sonotaSum + CDbl(aggItems(CStr(grp))(sonotaItemName))
                    End If
                Next k

                wsTarget.Cells(rowIdx, colOffset).Value = sonotaSum
                rowTotal = rowTotal + sonotaSum
                colOffset = colOffset + 1
            Next grp

            ' 合計列
            wsTarget.Cells(rowIdx, colOffset).Value = rowTotal

            rowIdx = rowIdx + 1
        End If

        ' ============================================
        ' テーブル化
        ' 理由：範囲をテーブル化してフィルタ機能と書式を適用
        ' ポイント：範囲取得時のエラーに備えてOn Error Resume Next
        ' ============================================
        Dim lastCol As Long
        lastCol = UBound(allGroups) + 3  ' 項目列 + グループ数 + 合計列

        Dim tableRange As Range
        On Error Resume Next
        Set tableRange = wsTarget.Range(wsTarget.Cells(outputStartRow, 1), wsTarget.Cells(rowIdx - 1, lastCol))
        On Error GoTo ErrorHandler

        If Not tableRange Is Nothing Then
            ' テーブル名の重複回避
            ' 理由：同じ期間名で複数回実行した場合のエラー防止
            Dim baseName As String, tryName As String, tryIdx As Long
            baseName = "_日報W_加工_" & Replace(periodName, " ", "_")
            tryName = baseName
            tryIdx = 1
            Do While TableExists(wsTarget, tryName)
                tryIdx = tryIdx + 1
                tryName = baseName & "_" & tryIdx
            Loop

            ' テーブル作成と書式設定
            Dim newTable As ListObject
            Set newTable = wsTarget.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
            newTable.Name = tryName

            On Error Resume Next
            newTable.TableStyle = "TableStyleLight16"
            newTable.ShowAutoFilter = False  ' フィルターボタンは非表示
            On Error GoTo ErrorHandler

            ' 列幅を統一
            Dim cIdx As Long
            For cIdx = 1 To newTable.Range.Columns.Count
                newTable.Range.Columns(cIdx).ColumnWidth = 8
            Next cIdx
        End If

        ' ============================================
        ' 印刷範囲の終了位置を更新
        ' 理由：このテーブルの最終行を記録（次のテーブルで上書きされる）
        ' ============================================
        printRangeEnd = rowIdx - 1

        ' 次のテーブルの開始位置（2行空ける）
        currentRow = rowIdx + 2

NextPeriod:
        ' 次の期間へ
    Next periodIdx

    ' ============================================
    ' 印刷範囲の設定
    ' 理由：出力した全テーブルを印刷範囲として設定
    ' 条件：データが1つでもあった場合のみ（printRangeStart > 0）
    ' ============================================
    If printRangeStart > 0 And printRangeEnd > 0 Then
        Dim printLastCol As Long
        printLastCol = UBound(allGroups) + 3  ' 項目列 + グループ数 + 合計列

        On Error Resume Next
        wsTarget.PageSetup.PrintArea = wsTarget.Range( _
            wsTarget.Cells(printRangeStart, 1), _
            wsTarget.Cells(printRangeEnd, printLastCol)).Address
        On Error GoTo ErrorHandler

        Application.StatusBar = "印刷範囲を設定しました"
    End If

Cleanup:
    ' ============================================
    ' 最適化設定の復元
    ' 理由：処理後は元の設定に戻す
    ' ============================================
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.DisplayAlerts = origDisplayAlerts
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    ' エラー時も設定を復元してから終了
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.DisplayAlerts = origDisplayAlerts
    Application.StatusBar = False

    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "転記_日報_加工W"
End Sub

' ============================================
' Private関数: TableExists
' 目的：シート上に指定テーブル名が存在するか安全に判定
' 引数：ws - 検索対象シート、tblName - テーブル名
' 戻り値：存在すればTrue、しなければFalse
' ============================================
Private Function TableExists(ws As Worksheet, tblName As String) As Boolean
    Dim lo As ListObject
    TableExists = False

    If ws Is Nothing Then Exit Function

    For Each lo In ws.ListObjects
        If lo.Name = tblName Then
            TableExists = True
            Exit Function
        End If
    Next lo
End Function

' ============================================
' Private関数: QuickSortDesc
' 目的：2次元配列を2列目（値）の降順でソート
' 引数：arr - ソート対象配列（1列目：項目名、2列目：値）
'       left - ソート開始位置、right - ソート終了位置
' 注意：配列は参照渡しなので直接書き換わる
' ============================================
Private Sub QuickSortDesc(ByRef arr() As Variant, ByVal left As Long, ByVal right As Long)
    Dim i As Long, j As Long, pivot As Double
    Dim tempName As String, tempValue As Double

    If left >= right Then Exit Sub

    i = left: j = right
    pivot = CDbl(arr((left + right) \ 2, 2))

    Do While i <= j
        Do While CDbl(arr(i, 2)) > pivot: i = i + 1: Loop
        Do While CDbl(arr(j, 2)) < pivot: j = j - 1: Loop
        If i <= j Then
            tempName = arr(i, 1): tempValue = arr(i, 2)
            arr(i, 1) = arr(j, 1): arr(i, 2) = arr(j, 2)
            arr(j, 1) = tempName: arr(j, 2) = tempValue
            i = i + 1: j = j - 1
        End If
    Loop

    If left < j Then Call QuickSortDesc(arr, left, j)
    If i < right Then Call QuickSortDesc(arr, i, right)
End Sub
