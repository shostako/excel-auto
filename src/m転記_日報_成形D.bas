Attribute VB_Name = "m転記_日報_成形D"
Option Explicit

' ========================================
' マクロ名: 転記_日報_成形D
' 処理概要: 日報データから期間別に時間当り出来高を集計する
'
' 【処理の特徴】
' 1. 空白期間スキップ：集計期間テーブルに行があっても、該当期間内にデータがなければテーブルを作らない
' 2. 動的期間対応：集計期間テーブルの行数が変わっても自動的に対応（増減どちらもOK）
'
' 【テーブル構成】
' 期間テーブル : シート「成形ND」、テーブル「_集計期間日報成形D」
' ソーステーブル : シート「日報成形」、テーブル「_日報成形」
' 出力テーブル : シート「成形ND」、複数テーブル「_日報D_成形_{期間名}」
'
' 【処理フロー】
' 1. 既存出力テーブルとデータを完全削除
' 2. 各期間ごとに日付フィルター + 時間当出来高集計
' 3. 集計結果を出力
' 4. データがある期間のみテーブル出力（空白期間はスキップ）
'
' 【出力形式】
' - 見出し：「項目」「全品番」の2列のみ
' - 1行目：ショット数；全品番の「ショット数」の合計
' - 2行目：良品数；全品番の「良品数」の合計
' - 3行目：不良数；全品番の「不良数」の合計
' - 4行目：稼働時間；全品番の「稼働時間」の合計；小数点以下1桁表記
' - 5行目：時間当出来高；「良品数」/(「稼働時間」/60)；小数点以下1桁表記
' - 6行目：出来高サイクル；「稼働時間」*60/「良品数」；小数点以下1桁表記
' ========================================

Sub 転記_日報_成形D()
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
    Application.StatusBar = "日報成形D転記処理を開始します..."

    ' ============================================
    ' シートとテーブルの参照取得
    ' 理由：オブジェクト参照で直接操作するため（Activateは使わない）
    ' ============================================
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Set wsSource = ThisWorkbook.Worksheets("日報成形")
    Set wsTarget = ThisWorkbook.Worksheets("成形ND")

    ' テーブル参照を取得（存在チェックはOn Error Resume Nextで安全に）
    Dim tblSource As ListObject, tblPeriod As ListObject
    On Error Resume Next
    Set tblSource = wsSource.ListObjects("_日報成形")
    Set tblPeriod = wsTarget.ListObjects("_集計期間日報成形D")
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
        MsgBox "「_集計期間日報成形D」に有効な集計期間がありません。処理を中止します。", vbExclamation
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
    Dim colHizuke As Long
    Dim colShot As Long, colRyohin As Long, colFuryo As Long, colKadoJikan As Long

    On Error Resume Next
    colHizuke = tblSource.ListColumns("日付").Index
    colShot = tblSource.ListColumns("ショット数").Index
    colRyohin = tblSource.ListColumns("良品数").Index
    colFuryo = tblSource.ListColumns("不良数").Index
    colKadoJikan = tblSource.ListColumns("稼働時間").Index
    On Error GoTo ErrorHandler

    ' 必須列のチェック
    If colHizuke = 0 Or colShot = 0 Or colRyohin = 0 Or colFuryo = 0 Or colKadoJikan = 0 Then
        MsgBox "必須列（日付、ショット数、良品数、不良数、稼働時間）が見つかりません。", vbCritical
        GoTo Cleanup
    End If

    ' ============================================
    ' 既存の出力テーブルオブジェクトを削除
    ' 理由：期間数が減った場合、古いテーブルが残るとエラーになる
    ' ポイント：逆順でループすることで削除中のインデックスずれを防止
    ' ============================================
    Dim idxLO As Long
    For idxLO = wsTarget.ListObjects.Count To 1 Step -1
        Dim loTemp As ListObject
        Set loTemp = wsTarget.ListObjects(idxLO)
        If loTemp.Name Like "_日報D_成形_*" Then
            loTemp.Delete
        End If
    Next idxLO

    ' ============================================
    ' 既存出力範囲の行削除
    ' 理由：テーブルオブジェクト削除後もセルの値は残るため、
    '       期間テーブルより下の行を全削除してクリーンアップ
    ' ============================================
    Dim periodTableLastRow As Long
    periodTableLastRow = 0
    If Not tblPeriod Is Nothing Then
        periodTableLastRow = tblPeriod.Range.Row + tblPeriod.Range.Rows.Count - 1
    End If

    Dim baseRow As Long
    baseRow = periodTableLastRow
    If baseRow < 1 Then baseRow = 1

    ' 基準行より下を全削除
    Dim lastUsedRow As Long
    lastUsedRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
    If lastUsedRow >= baseRow + 1 Then
        wsTarget.Rows((baseRow + 1) & ":" & lastUsedRow).Delete
    End If

    ' ============================================
    ' 出力開始位置の決定
    ' 理由：期間テーブルの下に2行空けてから出力開始
    ' ============================================
    Dim currentRow As Long
    currentRow = baseRow + 3

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
        ' 期間内データの集計
        ' 理由：該当期間のデータのみを抽出して集計
        ' ============================================
        Dim totalShot As Double, totalRyohin As Double, totalFuryo As Double, totalKadoJikan As Double
        totalShot = 0
        totalRyohin = 0
        totalFuryo = 0
        totalKadoJikan = 0

        Dim hasData As Boolean
        hasData = False

        Dim rowIdx As Long
        For rowIdx = 1 To UBound(srcArr, 1)
            Dim hizukeVal As Variant
            hizukeVal = srcArr(rowIdx, colHizuke)

            ' 日付チェック
            If IsDate(hizukeVal) Then
                Dim hizuke As Date
                hizuke = CDate(hizukeVal)

                ' 期間内かチェック
                If hizuke >= startDate And hizuke <= endDate Then
                    hasData = True

                    ' 集計
                    Dim shotVal As Double, ryohinVal As Double, furyoVal As Double, kadoJikanVal As Double

                    shotVal = 0
                    If IsNumeric(srcArr(rowIdx, colShot)) Then
                        shotVal = CDbl(srcArr(rowIdx, colShot))
                    End If

                    ryohinVal = 0
                    If IsNumeric(srcArr(rowIdx, colRyohin)) Then
                        ryohinVal = CDbl(srcArr(rowIdx, colRyohin))
                    End If

                    furyoVal = 0
                    If IsNumeric(srcArr(rowIdx, colFuryo)) Then
                        furyoVal = CDbl(srcArr(rowIdx, colFuryo))
                    End If

                    kadoJikanVal = 0
                    If IsNumeric(srcArr(rowIdx, colKadoJikan)) Then
                        kadoJikanVal = CDbl(srcArr(rowIdx, colKadoJikan))
                    End If

                    totalShot = totalShot + shotVal
                    totalRyohin = totalRyohin + ryohinVal
                    totalFuryo = totalFuryo + furyoVal
                    totalKadoJikan = totalKadoJikan + kadoJikanVal
                End If
            End If
        Next rowIdx

        ' ============================================
        ' 空白期間はスキップ
        ' 理由：データがない期間はテーブルを作らない
        ' ============================================
        If Not hasData Then
            Application.StatusBar = "期間「" & periodName & "」はデータなしのためスキップ"
            GoTo NextPeriod
        End If

        ' ============================================
        ' 印刷範囲の開始位置を記録（最初のテーブルのみ）
        ' 理由：全期間処理後に印刷範囲を一括設定するため
        ' ============================================
        If printRangeStart = 0 Then
            printRangeStart = currentRow
        End If

        ' ============================================
        ' タイトル行の出力
        ' 理由：各テーブルの上に期間名と日付範囲を表示
        ' 形式：日報D_成形_{期間名}_{開始日}～{終了日}
        ' ============================================
        Dim titleText As String
        titleText = "日報D_成形_" & periodName & "_" & Format(startDate, "m/d") & "～" & Format(endDate, "m/d")

        With wsTarget.Cells(currentRow, 1)
            .Value = titleText
            .Font.Bold = True
            .ShrinkToFit = False
        End With
        currentRow = currentRow + 1

        ' ============================================
        ' ヘッダー行の出力
        ' 理由：テーブルの列見出し
        ' ============================================
        wsTarget.Cells(currentRow, 1).Value = "項目"
        wsTarget.Cells(currentRow, 1).ShrinkToFit = True
        wsTarget.Cells(currentRow, 2).Value = "全品番"
        wsTarget.Cells(currentRow, 2).ShrinkToFit = True

        Dim headerRow As Long
        headerRow = currentRow
        currentRow = currentRow + 1

        ' ============================================
        ' データ行の出力
        ' 理由：集計結果を6行分出力
        ' ============================================
        Dim dataStartRow As Long
        dataStartRow = currentRow

        ' 1行目: ショット数
        wsTarget.Cells(currentRow, 1).Value = "ショット数"
        wsTarget.Cells(currentRow, 1).ShrinkToFit = True
        wsTarget.Cells(currentRow, 2).Value = totalShot
        currentRow = currentRow + 1

        ' 2行目: 良品数
        wsTarget.Cells(currentRow, 1).Value = "良品数"
        wsTarget.Cells(currentRow, 1).ShrinkToFit = True
        wsTarget.Cells(currentRow, 2).Value = totalRyohin
        currentRow = currentRow + 1

        ' 3行目: 不良数
        wsTarget.Cells(currentRow, 1).Value = "不良数"
        wsTarget.Cells(currentRow, 1).ShrinkToFit = True
        wsTarget.Cells(currentRow, 2).Value = totalFuryo
        currentRow = currentRow + 1

        ' 4行目: 稼働時間（小数点以下1桁）
        wsTarget.Cells(currentRow, 1).Value = "稼働時間"
        wsTarget.Cells(currentRow, 1).ShrinkToFit = True
        wsTarget.Cells(currentRow, 2).Value = totalKadoJikan
        wsTarget.Cells(currentRow, 2).NumberFormat = "0.0"
        currentRow = currentRow + 1

        ' 5行目: 時間当出来高 = 良品数 / (稼働時間 / 60)（小数点以下1桁）
        Dim jikanDekidaka As Double
        jikanDekidaka = 0
        If totalKadoJikan > 0 Then
            jikanDekidaka = totalRyohin / (totalKadoJikan / 60)
        End If
        wsTarget.Cells(currentRow, 1).Value = "時間当出来高"
        wsTarget.Cells(currentRow, 1).ShrinkToFit = True
        wsTarget.Cells(currentRow, 2).Value = jikanDekidaka
        wsTarget.Cells(currentRow, 2).NumberFormat = "0.0"
        currentRow = currentRow + 1

        ' 6行目: 出来高サイクル = 稼働時間 * 60 / 良品数（小数点以下1桁）
        Dim dekidakaCycle As Double
        dekidakaCycle = 0
        If totalRyohin > 0 Then
            dekidakaCycle = totalKadoJikan * 60 / totalRyohin
        End If
        wsTarget.Cells(currentRow, 1).Value = "出来高サイクル"
        wsTarget.Cells(currentRow, 1).ShrinkToFit = True
        wsTarget.Cells(currentRow, 2).Value = dekidakaCycle
        wsTarget.Cells(currentRow, 2).NumberFormat = "0.0"
        currentRow = currentRow + 1

        Dim dataEndRow As Long
        dataEndRow = currentRow - 1

        ' ============================================
        ' テーブル作成
        ' 理由：Excel標準のテーブル機能で見やすく管理
        ' ============================================
        Dim tblRange As Range
        Set tblRange = wsTarget.Range(wsTarget.Cells(headerRow, 1), wsTarget.Cells(dataEndRow, 2))

        Dim newTable As ListObject
        Dim tableName As String
        tableName = "_日報D_成形_" & periodName

        Set newTable = wsTarget.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
        newTable.Name = tableName
        newTable.ShowAutoFilter = False
        newTable.TableStyle = "TableStyleLight21"

        ' ============================================
        ' 列幅の設定
        ' 理由：項目列（1列目）を14に固定、データ列（2列目）は自動調整
        ' ============================================
        wsTarget.Columns(1).ColumnWidth = 14
        newTable.ListColumns(2).Range.EntireColumn.AutoFit

        ' ============================================
        ' 印刷範囲の終了位置を更新
        ' 理由：各テーブル出力後に最新の終了行を記録
        ' ============================================
        printRangeEnd = dataEndRow

        ' ============================================
        ' 次のテーブルとの間隔
        ' 理由：見やすさのため2行空ける
        ' ============================================
        currentRow = currentRow + 2

NextPeriod:
    Next periodIdx

    ' ============================================
    ' 印刷範囲の設定
    ' 理由：全テーブルを含む範囲を印刷範囲として設定
    ' ============================================
    If printRangeStart > 0 And printRangeEnd > 0 Then
        wsTarget.PageSetup.PrintArea = wsTarget.Range( _
            wsTarget.Cells(printRangeStart, 1), _
            wsTarget.Cells(printRangeEnd, 2)).Address

        Application.StatusBar = "印刷範囲を設定しました"
    End If

    Application.StatusBar = "処理が完了しました"

Cleanup:
    ' ============================================
    ' 最適化設定の復元
    ' 理由：元の設定に戻してExcelの動作を通常に戻す
    ' ============================================
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.DisplayAlerts = origDisplayAlerts
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
    Resume Cleanup
End Sub
