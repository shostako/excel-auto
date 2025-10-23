Attribute VB_Name = "m合計_流出_加工_期間対応"
Option Explicit

' ========================================
' マクロ名: 合計_流出_加工_期間対応
' 処理概要: 手直しと廃棄のテーブルを読み込んで流出合計を計算
' 手直しテーブル: シート「加工」複数テーブル「_手直し加工_{期間}」
' 廃棄テーブル: シート「加工」複数テーブル「_廃棄_加工_{期間}」
' 出力テーブル: シート「加工」複数テーブル「_流出_加工_{期間}」
' 処理方式: 手直しと廃棄の同一位置（項目×品番）の値を合計
' 出力位置: V列から開始、手直し・廃棄テーブルと水平配置
' 空白期間: 手直しと廃棄の両方のテーブルがない期間はスキップ
' 改善点:
'   - 期間テーブルから動的に期間名を取得してテーブル名・タイトルを生成
'   - 手直しテーブルを動的検索して水平配置基準位置を決定
' ========================================

Sub 合計_流出_加工_期間対応()
    ' 最適化設定の保存
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    Dim origEnableEvents As Boolean
    Dim origDisplayAlerts As Boolean
    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation
    origEnableEvents = Application.EnableEvents
    origDisplayAlerts = Application.DisplayAlerts

    ' 最適化設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ' エラーハンドリング設定
    On Error GoTo ErrorHandler

    ' ステータスバー初期化
    Application.StatusBar = "流出合計処理（加工）を開始します..."

    ' ============================================
    ' シート参照取得
    ' ============================================
    Dim wsTarget As Worksheet
    Set wsTarget = ThisWorkbook.Worksheets("加工")

    ' ============================================
    ' 既存の流出テーブルとタイトル行の削除
    ' ============================================
    On Error Resume Next

    ' 既存の流出テーブルを削除（"_流出_加工_"で始まるテーブル全て）
    Dim tbl As ListObject
    Dim tblsToDelete As Collection
    Set tblsToDelete = New Collection

    For Each tbl In wsTarget.ListObjects
        If Left(tbl.Name, 7) = "_流出_加工_" Then
            tblsToDelete.Add tbl
        End If
    Next tbl

    Dim j As Long
    For j = 1 To tblsToDelete.Count
        tblsToDelete(j).Range.EntireRow.Delete
    Next j

    ' 既存のタイトル行（「流出_加工」）をV列以降で検索して削除
    Dim searchRange As Range
    Set searchRange = wsTarget.Range("V:XFD")
    If Not searchRange Is Nothing Then
        Dim foundCell As Range
        Set foundCell = searchRange.Find("流出_加工", LookIn:=xlValues, LookAt:=xlPart)
        Do While Not foundCell Is Nothing
            foundCell.EntireRow.Delete
            Set foundCell = searchRange.Find("流出_加工", LookIn:=xlValues, LookAt:=xlPart)
        Loop
    End If

    Err.Clear
    On Error GoTo ErrorHandler

    ' ============================================
    ' 期間テーブルから期間情報取得
    ' ============================================
    Dim tblPeriod As ListObject
    Set tblPeriod = wsTarget.ListObjects("_集計期間加工")

    Dim periodData As Range
    Set periodData = tblPeriod.DataBodyRange

    If periodData Is Nothing Then
        MsgBox "「_集計期間加工」テーブルにデータがありません。", vbExclamation
        GoTo Cleanup
    End If

    ' 期間情報の配列作成（動的に全行対応）
    Dim periodCount As Long
    periodCount = periodData.Rows.Count

    Dim periodInfo() As Variant
    ReDim periodInfo(1 To periodCount, 1 To 3) ' 期間, 開始日, 終了日

    Dim p As Long
    For p = 1 To periodCount
        periodInfo(p, 1) = CStr(periodData.Cells(p, 1).Value) ' 期間
        periodInfo(p, 2) = periodData.Cells(p, 2).Value       ' 開始日
        periodInfo(p, 3) = periodData.Cells(p, 3).Value       ' 終了日
    Next p

    ' ============================================
    ' 手直しテーブルの位置を動的検索して開始位置を決定
    ' ============================================
    Dim firstHandaosiTable As ListObject
    Dim startRow As Long

    ' 最初の手直しテーブルを動的に探す（"_手直し加工_"で始まる最初のテーブル）
    For Each tbl In wsTarget.ListObjects
        If Left(tbl.Name, 7) = "_手直し加工_" Then
            Set firstHandaosiTable = tbl
            Exit For
        End If
    Next tbl

    ' 手直しテーブルのタイトル行位置を取得（テーブルの1行上）
    If Not firstHandaosiTable Is Nothing Then
        startRow = firstHandaosiTable.Range.Row - 1
    Else
        ' 手直しテーブルがない場合はエラー
        MsgBox "手直しテーブルが見つかりません。先に手直しマクロを実行してください。", vbExclamation
        GoTo Cleanup
    End If

    Dim currentRow As Long
    currentRow = startRow

    ' ============================================
    ' 各期間の処理ループ
    ' ============================================
    For p = 1 To periodCount
        Application.StatusBar = "期間 " & p & "/" & periodCount & " を処理中..."

        Dim periodName As String, startDate As Date, endDate As Date
        periodName = periodInfo(p, 1)
        startDate = CDate(periodInfo(p, 2))
        endDate = CDate(periodInfo(p, 3))

        ' ============================================
        ' 手直しと廃棄のテーブルを動的に取得
        ' ============================================
        Dim handaosiTable As ListObject, haikTable As ListObject
        Set handaosiTable = Nothing
        Set haikTable = Nothing

        ' テーブル名を動的に構築して検索
        Dim handaosiTableName As String, haikTableName As String
        handaosiTableName = "_手直し加工_" & periodName
        haikTableName = "_廃棄_加工_" & periodName

        For Each tbl In wsTarget.ListObjects
            If tbl.Name = handaosiTableName Then
                Set handaosiTable = tbl
            ElseIf tbl.Name = haikTableName Then
                Set haikTable = tbl
            End If
        Next tbl

        ' テーブルの存在確認
        If handaosiTable Is Nothing Then
            Debug.Print "警告: " & handaosiTableName & " が見つかりません"
        End If
        If haikTable Is Nothing Then
            Debug.Print "警告: " & haikTableName & " が見つかりません"
        End If

        ' ============================================
        ' 空白期間スキップ処理
        ' ============================================
        ' 手直しと廃棄の両方のテーブルがない場合はスキップ
        ' (手直しと廃棄のマクロで空白期間スキップが働いているため、
        '  両方のテーブルがない=空白期間となる)
        If handaosiTable Is Nothing And haikTable Is Nothing Then
            Application.StatusBar = "期間 " & p & " はデータ無しのためスキップします..."
            GoTo NextPeriod
        End If

        ' ============================================
        ' データ配列の作成（9分類+1項目列=10列）+集計行追加
        ' ============================================
        Dim outputData() As Variant
        Dim baseRowCount As Long

        ' 行数を決定（廃棄テーブル優先、なければ手直し）
        ' DataBodyRangeから集計行を除外
        If Not haikTable Is Nothing Then
            baseRowCount = haikTable.DataBodyRange.Rows.Count
            ' 最後の行が「合計」かチェック
            If haikTable.DataBodyRange.Cells(baseRowCount, 1).Value = "合計" Then
                baseRowCount = baseRowCount - 1
            End If
        ElseIf Not handaosiTable Is Nothing Then
            baseRowCount = handaosiTable.DataBodyRange.Rows.Count
            ' 最後の行が「合計」かチェック
            If handaosiTable.DataBodyRange.Cells(baseRowCount, 1).Value = "合計" Then
                baseRowCount = baseRowCount - 1
            End If
        Else
            baseRowCount = 0
        End If

        ' 0:ヘッダー, 1～baseRowCount:データ, baseRowCount+1:集計行
        ReDim outputData(0 To baseRowCount + 1, 0 To 9)

        ' ヘッダー行設定
        outputData(0, 0) = "項目"
        outputData(0, 1) = "58050FrLH"
        outputData(0, 2) = "58050FrRH"
        outputData(0, 3) = "58050RrLH"
        outputData(0, 4) = "58050RrRH"
        outputData(0, 5) = "28050FrLH"
        outputData(0, 6) = "28050FrRH"
        outputData(0, 7) = "28050RrLH"
        outputData(0, 8) = "28050RrRH"
        outputData(0, 9) = "補給品"

        ' ============================================
        ' データ行の合計計算
        ' ============================================
        Dim i As Long, k As Long
        Dim colSums(1 To 9) As Double  ' 各列の合計用
        Dim csIdx As Long
        For csIdx = 1 To 9
            colSums(csIdx) = 0
        Next csIdx

        For i = 1 To baseRowCount
            ' 項目名の設定（手直しまたは廃棄から取得）
            If Not handaosiTable Is Nothing Then
                outputData(i, 0) = handaosiTable.DataBodyRange.Cells(i, 1).Value
            ElseIf Not haikTable Is Nothing Then
                outputData(i, 0) = haikTable.DataBodyRange.Cells(i, 1).Value
            End If

            ' 各品番列の合計計算（9列固定）
            For k = 1 To 9
                Dim handaosiValue As Double, haikValue As Double
                handaosiValue = 0
                haikValue = 0

                ' 手直しテーブルから値取得（k=9の補給品はスキップ）
                If Not handaosiTable Is Nothing And k <= 8 Then
                    If k <= handaosiTable.DataBodyRange.Columns.Count - 1 Then
                        If IsNumeric(handaosiTable.DataBodyRange.Cells(i, k + 1).Value) Then
                            handaosiValue = CDbl(handaosiTable.DataBodyRange.Cells(i, k + 1).Value)
                        End If
                    End If
                End If

                ' 廃棄テーブルから値取得
                If Not haikTable Is Nothing Then
                    If k <= haikTable.DataBodyRange.Columns.Count - 1 Then
                        If IsNumeric(haikTable.DataBodyRange.Cells(i, k + 1).Value) Then
                            haikValue = CDbl(haikTable.DataBodyRange.Cells(i, k + 1).Value)
                        End If
                    End If
                End If

                ' 合計値を設定
                outputData(i, k) = handaosiValue + haikValue
                ' 列合計に加算
                colSums(k) = colSums(k) + outputData(i, k)
            Next k
        Next i

        ' ============================================
        ' 集計行の追加
        ' ============================================
        outputData(baseRowCount + 1, 0) = "合計"
        For k = 1 To 9
            outputData(baseRowCount + 1, k) = colSums(k)
        Next k

        ' ============================================
        ' タイトル行の生成と出力（V列）
        ' ============================================
        Dim titleText As String
        titleText = "流出_加工_" & periodName & "_" & Format(startDate, "m/d") & "～" & Format(endDate, "m/d")

        ' タイトルセルの書式設定
        With wsTarget.Cells(currentRow, 22) ' V列 = 22
            .Value = titleText
            .ShrinkToFit = False  ' 縮小して全体を表示しない
            .WrapText = False     ' 折り返しなし
            .Font.Bold = True     ' 太字
            .Font.Size = 12       ' フォントサイズ12
        End With

        ' テーブル開始位置（V列）
        Dim startCell As Range
        Set startCell = wsTarget.Cells(currentRow + 1, 22) ' V列 = 22

        ' ============================================
        ' テーブル範囲への書き込み
        ' ============================================
        Dim tableRange As Range
        Set tableRange = startCell.Resize(UBound(outputData, 1) + 1, UBound(outputData, 2) + 1)
        tableRange.Value = outputData

        ' ============================================
        ' ListObjectとして設定
        ' ============================================
        Dim newTable As ListObject
        Set newTable = wsTarget.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        newTable.Name = "_流出_加工_" & periodName
        newTable.ShowAutoFilter = False ' フィルターボタンを非表示
        newTable.TableStyle = "TableStyleLight16" ' テーブルスタイル設定

        ' ============================================
        ' 書式設定：フォント、サイズ、列幅
        ' ============================================
        With newTable.Range
            .Font.Name = "游ゴシック"
            .Font.Size = 11
            .ShrinkToFit = True ' 縮小して全体を表示
        End With

        ' 列幅設定
        For i = 1 To newTable.Range.Columns.Count
            newTable.Range.Columns(i).ColumnWidth = 8
        Next i

        ' 次のテーブル位置を計算（現在のテーブル + 3行間隔）
        currentRow = startCell.Row + UBound(outputData, 1) + 3

NextPeriod:
    Next p

    ' 処理完了のステータスバー表示
    Application.StatusBar = "流出合計処理が完了しました（" & periodCount & "期間）"
    Application.Wait Now + TimeValue("00:00:01")

    GoTo Cleanup

ErrorHandler:
    ' エラー情報の詳細化
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear

    MsgBox "エラーが発生しました" & vbCrLf & _
           "エラー番号: " & errNum & vbCrLf & _
           "詳細: " & errDesc, vbCritical, "合計_流出_加工_期間対応 エラー"

Cleanup:
    ' 設定を確実に復元
    Application.StatusBar = False
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.DisplayAlerts = origDisplayAlerts
End Sub
