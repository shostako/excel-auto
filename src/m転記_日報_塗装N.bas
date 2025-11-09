Attribute VB_Name = "m転記_日報_塗装N"
Option Explicit

' ========================================
' マクロ名: 転記_日報_塗装
' 処理概要: 日報データを期間別に9分類で集計して塗装シートに転記
'
' 【処理の特徴】
' 1. 空白期間スキップ：集計期間テーブルに行があっても、該当期間内にデータがなければテーブルを作らない
' 2. 動的期間対応：集計期間テーブルの行数が変わっても自動的に対応（増減どちらもOK）
' 3. 高速化：配列処理による大量データの高速集計
'
' 【テーブル構成】
' 期間テーブル : シート「塗装N」、テーブル「_集計期間日報塗装」
' ソーステーブル : シート「日報塗装」、テーブル「_日報塗装」
' 項目テーブル : シート「塗装N」、テーブル「_日報項目塗装」
' 出力テーブル : シート「塗装N」、複数テーブル「_日報_塗装_{期間名}」
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
' - 最終列：合計
' ========================================

Sub 転記_日報_塗装N()
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
    Application.StatusBar = "日報塗装転記処理を開始します..."

    ' ============================================
    ' シートとテーブルの参照取得
    ' 理由：オブジェクト参照で直接操作するため（Activateは使わない）
    ' ============================================
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Set wsSource = ThisWorkbook.Worksheets("日報塗装")
    Set wsTarget = ThisWorkbook.Worksheets("塗装N")

    ' テーブル参照を取得（存在チェックはOn Error Resume Nextで安全に）
    Dim tblSource As ListObject, tblItems As ListObject, tblPeriod As ListObject
    On Error Resume Next
    Set tblSource = wsSource.ListObjects("_日報塗装")
    Set tblItems = wsTarget.ListObjects("_日報項目塗装")
    Set tblPeriod = wsTarget.ListObjects("_集計期間日報塗装")
    On Error GoTo ErrorHandler

    ' ソーステーブルは必須
    If tblSource Is Nothing Then
        MsgBox "シート「日報塗装」にテーブル「_日報塗装」が見つかりません。", vbCritical
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
        MsgBox "「_集計期間日報塗装」に有効な集計期間がありません。処理を中止します。", vbExclamation
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

    ' リコート、廃棄列の存在チェック
    ' 理由：これらの列が存在しない場合はエラーで処理中断
    On Error Resume Next
    Dim colRecoat As Long, colHaiki As Long
    colRecoat = 0: colHaiki = 0
    colRecoat = tblSource.ListColumns("リコート").Index
    colHaiki = tblSource.ListColumns("廃棄").Index
    On Error GoTo ErrorHandler

    If colRecoat = 0 Or colHaiki = 0 Then
        MsgBox "ソーステーブルに「リコート」列または「廃棄」列が見つかりません。" & vbCrLf & _
               "処理を中止します。", vbCritical
        GoTo Cleanup
    End If

    ' 集計対象列のインデックスをDictionaryで管理
    ' 理由：列が存在しない場合でもエラーにならず、柔軟に対応
    Dim colIndexes As Object
    Set colIndexes = CreateObject("Scripting.Dictionary")

    Dim targetColumns As Variant
    targetColumns = Array("ショット数", "不良数", "リコート", "廃棄", "ヒゲ", "ミスト", "ライン", "ゴミ", "スケ", _
                          "ピンホール", "マット", "その他O", "タレ", "キズ", "再塗装", "成形")

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
    ' 9分類グループの定義
    ' 理由：品番2列の値がこれらに該当すればそのグループで集計、
    '       該当しなければ「補給品」に振り分ける
    ' ============================================
    Dim validGroups As Object
    Set validGroups = CreateObject("Scripting.Dictionary")

    validGroups("58050FrLH") = True: validGroups("58050FrRH") = True
    validGroups("58050RrLH") = True: validGroups("58050RrRH") = True
    validGroups("28050FrLH") = True: validGroups("28050FrRH") = True
    validGroups("28050RrLH") = True: validGroups("28050RrRH") = True
    validGroups("補給品") = True

    ' ============================================
    ' 既存の出力テーブルオブジェクトを削除
    ' 理由：期間数が減った場合、古いテーブルが残るとエラーになる
    ' ポイント：逆順でループすることで削除中のインデックスずれを防止
    ' ============================================
    Dim idxLO As Long
    For idxLO = wsTarget.ListObjects.Count To 1 Step -1
        Dim loTemp As ListObject
        Set loTemp = wsTarget.ListObjects(idxLO)
        If loTemp.Name Like "_日報_塗装_*" Then
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
        '   aggRecoat: グループ名 → リコート合計
        '   aggHaiki: グループ名 → 廃棄合計
        '   aggItems: グループ名 → Dictionary(項目名 → 合計値)
        '   aggSonota: グループ名 → その他合計
        ' ============================================
        Dim aggShot As Object, aggFuryo As Object, aggRecoat As Object, aggHaiki As Object
        Dim aggItems As Object, aggSonota As Object
        Set aggShot = CreateObject("Scripting.Dictionary")
        Set aggFuryo = CreateObject("Scripting.Dictionary")
        Set aggRecoat = CreateObject("Scripting.Dictionary")
        Set aggHaiki = CreateObject("Scripting.Dictionary")
        Set aggItems = CreateObject("Scripting.Dictionary")
        Set aggSonota = CreateObject("Scripting.Dictionary")

        Dim grp As Variant
        For Each grp In allGroups
            aggShot(CStr(grp)) = 0
            aggFuryo(CStr(grp)) = 0
            aggRecoat(CStr(grp)) = 0
            aggHaiki(CStr(grp)) = 0
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
                    ' 品番2列の値で9分類を判定
                    Dim hinbanVal As String
                    hinbanVal = Trim(CStr(srcArr(r, colHinban)))

                    Dim targetGroup As String
                    If validGroups.Exists(hinbanVal) Then
                        targetGroup = hinbanVal
                    Else
                        targetGroup = "補給品"  ' 9分類に該当しない場合
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
                                numValue = CDbl(colValue)

                                ' 列名による振り分け
                                If CStr(keyName) = "ショット数" Then
                                    aggShot(targetGroup) = aggShot(targetGroup) + numValue
                                ElseIf CStr(keyName) = "不良数" Then
                                    aggFuryo(targetGroup) = aggFuryo(targetGroup) + numValue
                                ElseIf CStr(keyName) = "リコート" Then
                                    aggRecoat(targetGroup) = aggRecoat(targetGroup) + numValue
                                ElseIf CStr(keyName) = "廃棄" Then
                                    aggHaiki(targetGroup) = aggHaiki(targetGroup) + numValue
                                ElseIf itemsList.Exists(CStr(keyName)) Then
                                    ' 項目テーブルに定義されている項目
                                    aggItems(targetGroup)(CStr(keyName)) = aggItems(targetGroup)(CStr(keyName)) + numValue
                                Else
                                    ' その他（項目テーブルにない列）
                                    aggSonota(targetGroup) = aggSonota(targetGroup) + numValue
                                End If
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
        titleText = "日報_塗装_" & periodName & "_" & Format(startDate, "m/d") & "‾" & Format(endDate, "m/d")

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

        ' 3行目：リコート
        With wsTarget.Cells(rowIdx, 1)
            .Value = "リコート"
            .ShrinkToFit = True
        End With
        rowTotal = 0
        colOffset = 2
        For Each grp In allGroups
            cellValue = aggRecoat(CStr(grp))
            wsTarget.Cells(rowIdx, colOffset).Value = cellValue
            rowTotal = rowTotal + cellValue
            colOffset = colOffset + 1
        Next grp
        ' 合計列
        wsTarget.Cells(rowIdx, colOffset).Value = rowTotal
        rowIdx = rowIdx + 1

        ' 4行目：廃棄
        With wsTarget.Cells(rowIdx, 1)
            .Value = "廃棄"
            .ShrinkToFit = True
        End With
        rowTotal = 0
        colOffset = 2
        For Each grp In allGroups
            cellValue = aggHaiki(CStr(grp))
            wsTarget.Cells(rowIdx, colOffset).Value = cellValue
            rowTotal = rowTotal + cellValue
            colOffset = colOffset + 1
        Next grp
        ' 合計列
        wsTarget.Cells(rowIdx, colOffset).Value = rowTotal
        rowIdx = rowIdx + 1

        ' 5行目以降：項目別集計（項目テーブルの順序）
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
            baseName = "_日報_塗装_" & Replace(periodName, " ", "_")
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
            newTable.TableStyle = "TableStyleLight17"
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
           "エラー番号: " & Err.Number, vbCritical, "転記_日報_塗装"
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
