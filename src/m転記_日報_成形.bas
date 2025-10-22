Attribute VB_Name = "m転記_日報_成形"
Option Explicit

' ========================================
' マクロ名: 転記_日報_成形
' 処理概要: 日報データを期間別に9分類で集計して成形シートに転記
'
' 【修正履歴】
' - 出力の「不良数」列の集計方法を変更
' - ソースの「不良数」列は無視
' - 各不良項目列（打出し、ショート、ウエルド等19項目）の合計値を「不良数」として出力
'
' 【処理の特徴】
' 1. 空白期間スキップ：集計期間テーブルに行があっても、該当期間内にデータがなければテーブルを作らない
' 2. 動的期間対応：集計期間テーブルの行数が変わっても自動的に対応（増減どちらもOK）
' 3. 高速化：配列処理による大量データの高速集計
'
' 【テーブル構成】
' 期間テーブル : シート「成形N」、テーブル「_集計期間日報成形」
' ソーステーブル : シート「日報成形」、テーブル「_日報成形」
' 項目テーブル : シート「成形N」、テーブル「_日報項目成形」
' 出力テーブル : シート「成形N」、複数テーブル「_日報_成形_{期間名}」
'
' 【処理フロー】
' 1. 既存出力テーブルとデータを完全削除
' 2. 各期間ごとに日付フィルター + 品番による9分類集計
' 3. データがある期間のみテーブル出力（空白期間はスキップ）
'
' 【出力形式】
' - 1行目：ショット数
' - 2行目：不良数（各項目の合計値）
' - 3行目以降：項目別集計（項目テーブルの順序）
' - 最終行：その他
'
' 【集計方法】
' 品番2列の文字列別で9分類の集計をする。
' 「58050FrLH」：「58050FrSET」「58050FrLH」の行を集計する
' 「58050FrRH」：「58050FrSET」「58050FrRH」の行を集計する
' 「58050RrLH」：「58050RrSET」「58050RrLH」の行を集計する
' 「58050RrRH」：「58050RrSET」「58050RrRH」の行を集計する
' 「28050FrLH」：「28050FrSET」「28050FrLH」の行を集計する
' 「28050FrRH」：「28050FrSET」「28050FrRH」の行を集計する
' 「28050RrLH」：「28050RrSET」「28050RrLH」の行を集計する
' 「28050RrRH」：「28050RrSET」「28050RrRH」の行を集計する
' 「補給品」：「補給品FrLH」「補給品FrRH」「補給品RrLH」「補給品RrRH」の行
'            および、「補給品FrSET」「補給品RrSET」の行を2倍したものを集計する。
' ========================================

Sub 転記_日報_成形()
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
    Application.StatusBar = "日報成形転記処理を開始します..."

    ' ============================================
    ' シートとテーブルの参照取得
    ' 理由：オブジェクト参照で直接操作するため（Activateは使わない）
    ' ============================================
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Set wsSource = ThisWorkbook.Worksheets("日報成形")
    Set wsTarget = ThisWorkbook.Worksheets("成形N")

    ' テーブル参照を取得（存在チェックはOn Error Resume Nextで安全に）
    Dim tblSource As ListObject, tblItems As ListObject, tblPeriod As ListObject
    On Error Resume Next
    Set tblSource = wsSource.ListObjects("_日報成形")
    Set tblItems = wsTarget.ListObjects("_日報項目成形")
    Set tblPeriod = wsTarget.ListObjects("_集計期間日報成形")
    On Error GoTo ErrorHandler

    ' ソーステーブルは必須
    If tblSource Is Nothing Then
        MsgBox "シート「日報成形」にテーブル「_日報成形」が見つかりません。", vbCritical
        GoTo Cleanup
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
        MsgBox "「_集計期間日報成形」に有効な集計期間がありません。処理を中止します。", vbExclamation
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
    Dim colHizuke As Long, colHinban As Long
    colHizuke = tblSource.ListColumns("日付").Index
    colHinban = tblSource.ListColumns("品番2").Index

    ' 集計対象列のインデックスをDictionaryで管理
    ' 理由：列が存在しない場合でもエラーにならず、柔軟に対応
    Dim colIndexes As Object
    Set colIndexes = CreateObject("Scripting.Dictionary")

    Dim targetColumns As Variant
    targetColumns = Array("ショット数", "不良数", "打出し", "ショート", "ウエルド", "シワ", "異物", _
                          "シルバー", "フローマーク", "ゴミ押し", "GCカス", "キズ", "ヒケ", "糸引き", _
                          "型汚れ", "マクレ", "取出不良", "割れ白化", "コアカス", "その他O", "チョコ停打出し")

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
    ' 項目テーブルから項目リスト取得
    ' 理由：出力時の項目順序を項目テーブルの定義通りにするため
    ' ポイント：項目テーブルが存在しない、または空でもエラーにならない
    ' ============================================
    Dim itemsList As Object
    Set itemsList = CreateObject("Scripting.Dictionary")

    If Not tblItems Is Nothing Then
        If Not tblItems.DataBodyRange Is Nothing Then
            Dim itemsData As Range
            Set itemsData = tblItems.DataBodyRange
            Dim i As Long
            For i = 1 To itemsData.Rows.Count
                Dim itemName As String
                itemName = CStr(itemsData.Cells(i, 1).Value)
                ' 「その他」は除外（別途処理するため）
                If Len(itemName) > 0 And itemName <> "その他" Then
                    itemsList(itemName) = i
                End If
            Next i
        End If
    End If

    ' ============================================
    ' 不良項目列のリスト定義
    ' 理由：これらの列の合計値を「不良数」として集計する
    ' ============================================
    Dim furyoItemColumns As Object
    Set furyoItemColumns = CreateObject("Scripting.Dictionary")
    furyoItemColumns("打出し") = 1
    furyoItemColumns("ショート") = 1
    furyoItemColumns("ウエルド") = 1
    furyoItemColumns("シワ") = 1
    furyoItemColumns("異物") = 1
    furyoItemColumns("シルバー") = 1
    furyoItemColumns("フローマーク") = 1
    furyoItemColumns("ゴミ押し") = 1
    furyoItemColumns("GCカス") = 1
    furyoItemColumns("キズ") = 1
    furyoItemColumns("ヒケ") = 1
    furyoItemColumns("糸引き") = 1
    furyoItemColumns("型汚れ") = 1
    furyoItemColumns("マクレ") = 1
    furyoItemColumns("取出不良") = 1
    furyoItemColumns("割れ白化") = 1
    furyoItemColumns("コアカス") = 1
    furyoItemColumns("その他O") = 1
    furyoItemColumns("チョコ停打出し") = 1

    ' ============================================
    ' 9分類グループの定義と品番マッピング
    ' 理由：品番2の値をどのグループに振り分けるか、
    '       また補給品の場合は倍率をかける必要があるため
    ' ============================================

    ' グループごとに該当する品番のリスト（Dictionary）を持つ
    Dim groupMapping As Object
    Set groupMapping = CreateObject("Scripting.Dictionary")

    ' 各グループに振り分けロジック用のDictionaryを作成
    ' 構造: groupMapping("グループ名") = Dictionary("品番2文字列" -> 倍率)

    Dim g As Variant
    Dim grpDetail As Object

    ' 58050FrLH グループ
    Set grpDetail = CreateObject("Scripting.Dictionary")
    grpDetail("58050FrSET") = 1
    grpDetail("58050FrLH") = 1
    Set groupMapping("58050FrLH") = grpDetail

    ' 58050FrRH グループ
    Set grpDetail = CreateObject("Scripting.Dictionary")
    grpDetail("58050FrSET") = 1
    grpDetail("58050FrRH") = 1
    Set groupMapping("58050FrRH") = grpDetail

    ' 58050RrLH グループ
    Set grpDetail = CreateObject("Scripting.Dictionary")
    grpDetail("58050RrSET") = 1
    grpDetail("58050RrLH") = 1
    Set groupMapping("58050RrLH") = grpDetail

    ' 58050RrRH グループ
    Set grpDetail = CreateObject("Scripting.Dictionary")
    grpDetail("58050RrSET") = 1
    grpDetail("58050RrRH") = 1
    Set groupMapping("58050RrRH") = grpDetail

    ' 28050FrLH グループ
    Set grpDetail = CreateObject("Scripting.Dictionary")
    grpDetail("28050FrSET") = 1
    grpDetail("28050FrLH") = 1
    Set groupMapping("28050FrLH") = grpDetail

    ' 28050FrRH グループ
    Set grpDetail = CreateObject("Scripting.Dictionary")
    grpDetail("28050FrSET") = 1
    grpDetail("28050FrRH") = 1
    Set groupMapping("28050FrRH") = grpDetail

    ' 28050RrLH グループ
    Set grpDetail = CreateObject("Scripting.Dictionary")
    grpDetail("28050RrSET") = 1
    grpDetail("28050RrLH") = 1
    Set groupMapping("28050RrLH") = grpDetail

    ' 28050RrRH グループ
    Set grpDetail = CreateObject("Scripting.Dictionary")
    grpDetail("28050RrSET") = 1
    grpDetail("28050RrRH") = 1
    Set groupMapping("28050RrRH") = grpDetail

    ' 補給品グループ（SETは2倍）
    Set grpDetail = CreateObject("Scripting.Dictionary")
    grpDetail("補給品FrLH") = 1
    grpDetail("補給品FrRH") = 1
    grpDetail("補給品RrLH") = 1
    grpDetail("補給品RrRH") = 1
    grpDetail("補給品FrSET") = 2  ' 2倍
    grpDetail("補給品RrSET") = 2  ' 2倍
    Set groupMapping("補給品") = grpDetail

    ' ============================================
    ' 既存の出力テーブルオブジェクトを削除
    ' 理由：期間数が減った場合、古いテーブルが残るとエラーになる
    ' ポイント：逆順でループすることで削除中のインデックスずれを防止
    ' ============================================
    Dim idxLO As Long
    For idxLO = wsTarget.ListObjects.Count To 1 Step -1
        Dim loTemp As ListObject
        Set loTemp = wsTarget.ListObjects(idxLO)
        If loTemp.Name Like "_日報_成形_*" Then
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
        '   aggSonota: グループ名 → その他合計
        ' ============================================
        Dim aggShot As Object, aggFuryo As Object, aggItems As Object, aggSonota As Object
        Set aggShot = CreateObject("Scripting.Dictionary")
        Set aggFuryo = CreateObject("Scripting.Dictionary")
        Set aggItems = CreateObject("Scripting.Dictionary")
        Set aggSonota = CreateObject("Scripting.Dictionary")

        Dim grp As Variant
        For Each grp In allGroups
            aggShot(CStr(grp)) = 0
            aggFuryo(CStr(grp)) = 0
            aggSonota(CStr(grp)) = 0

            ' 項目別集計用のネストDictionary
            Set aggItems(CStr(grp)) = CreateObject("Scripting.Dictionary")
            Dim itemKey As Variant
            For Each itemKey In itemsList.Keys
                aggItems(CStr(grp))(CStr(itemKey)) = 0
            Next itemKey
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
                    ' 品番2列の値
                    Dim hinbanVal As String
                    hinbanVal = Trim(CStr(srcArr(r, colHinban)))

                    ' この品番がどのグループに該当するかを全て列挙
                    ' 理由：SETは複数グループに同時加算されるため
                    Dim matchedGroups As Object
                    Set matchedGroups = CreateObject("Scripting.Dictionary")

                    Dim grpKey As Variant
                    For Each grpKey In groupMapping.Keys
                        Dim grpDic As Object
                        Set grpDic = groupMapping(CStr(grpKey))

                        If grpDic.Exists(hinbanVal) Then
                            ' グループ名と倍率をセットで保存
                            matchedGroups(CStr(grpKey)) = CDbl(grpDic(hinbanVal))
                        End If
                    Next grpKey

                    ' どのグループにも該当しない場合はスキップ
                    If matchedGroups.Count = 0 Then
                        GoTo NextRow
                    End If

                    ' 各列の値を集計（該当する全グループに加算）
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
                                Dim baseValue As Double
                                baseValue = CDbl(colValue)

                                ' マッチした全グループに加算（倍率を適用）
                                Dim targetGroup As Variant
                                For Each targetGroup In matchedGroups.Keys
                                    Dim numValue As Double
                                    numValue = baseValue * CDbl(matchedGroups(targetGroup))

                                    ' 列名による振り分け
                                    If CStr(keyName) = "ショット数" Then
                                        aggShot(CStr(targetGroup)) = aggShot(CStr(targetGroup)) + numValue
                                    ElseIf CStr(keyName) = "不良数" Then
                                        ' ソースの「不良数」列は無視（何もしない）
                                        ' 理由：各不良項目列の合計値を不良数として使用するため
                                    ElseIf itemsList.Exists(CStr(keyName)) Then
                                        ' 項目テーブルに定義されている項目
                                        aggItems(CStr(targetGroup))(CStr(keyName)) = aggItems(CStr(targetGroup))(CStr(keyName)) + numValue
                                        ' 不良項目の場合は aggFuryo にも加算
                                        If furyoItemColumns.Exists(CStr(keyName)) Then
                                            aggFuryo(CStr(targetGroup)) = aggFuryo(CStr(targetGroup)) + numValue
                                        End If
                                    Else
                                        ' その他（項目テーブルにない列）
                                        aggSonota(CStr(targetGroup)) = aggSonota(CStr(targetGroup)) + numValue
                                        ' 不良項目の場合は aggFuryo にも加算
                                        If furyoItemColumns.Exists(CStr(keyName)) Then
                                            aggFuryo(CStr(targetGroup)) = aggFuryo(CStr(targetGroup)) + numValue
                                        End If
                                    End If
                                Next targetGroup
                            End If
                        End If
                    Next keyName
                End If
            End If

NextRow:
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
        titleText = "日報_成形_" & periodName & "_" & Format(startDate, "m/d") & "‾" & Format(endDate, "m/d")

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

        ' 3行目以降：項目別集計（項目テーブルの順序）
        Dim itemKey2 As Variant
        For Each itemKey2 In itemsList.Keys
            With wsTarget.Cells(rowIdx, 1)
                .Value = CStr(itemKey2)
                .ShrinkToFit = True
            End With
            rowTotal = 0
            colOffset = 2
            For Each grp In allGroups
                cellValue = aggItems(CStr(grp))(CStr(itemKey2))
                wsTarget.Cells(rowIdx, colOffset).Value = cellValue
                rowTotal = rowTotal + cellValue
                colOffset = colOffset + 1
            Next grp
            ' 合計列
            wsTarget.Cells(rowIdx, colOffset).Value = rowTotal
            rowIdx = rowIdx + 1
        Next itemKey2

        ' 最終行：その他
        With wsTarget.Cells(rowIdx, 1)
            .Value = "その他"
            .ShrinkToFit = True
        End With
        rowTotal = 0
        colOffset = 2
        For Each grp In allGroups
            cellValue = aggSonota(CStr(grp))
            wsTarget.Cells(rowIdx, colOffset).Value = cellValue
            rowTotal = rowTotal + cellValue
            colOffset = colOffset + 1
        Next grp
        ' 合計列
        wsTarget.Cells(rowIdx, colOffset).Value = rowTotal

        ' ============================================
        ' テーブル化
        ' 理由：範囲をテーブル化してフィルタ機能と書式を適用
        ' ポイント：範囲取得時のエラーに備えてOn Error Resume Next
        ' ============================================
        Dim lastCol As Long
        lastCol = UBound(allGroups) + 3  ' 項目列 + グループ数 + 合計列

        Dim tableRange As Range
        On Error Resume Next
        Set tableRange = wsTarget.Range(wsTarget.Cells(outputStartRow, 1), wsTarget.Cells(rowIdx, lastCol))
        On Error GoTo ErrorHandler

        If Not tableRange Is Nothing Then
            ' テーブル名の重複回避
            ' 理由：同じ期間名で複数回実行した場合のエラー防止
            Dim baseName As String, tryName As String, tryIdx As Long
            baseName = "_日報_成形_" & Replace(periodName, " ", "_")
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
            newTable.TableStyle = "TableStyleLight21"
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
        printRangeEnd = rowIdx

        ' 次のテーブルの開始位置（2行空ける）
        currentRow = rowIdx + 3

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
           "エラー番号: " & Err.Number, vbCritical, "転記_日報_成形"
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
