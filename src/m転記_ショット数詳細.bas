Attribute VB_Name = "m転記_ショット数詳細"
Option Explicit

' ========================================
' マクロ名: 転記_ショット数詳細
' 処理概要: アクティブシートの集計期間テーブルから期間を取得し、
'          その期間内のショット数詳細（ロット番号と数量）を表形式で出力
'
' 【処理の特徴】
' 1. 動的期間取得：アクティブシートの「_集計期間」で始まるテーブルの1行目から期間を取得
' 2. 動的出力位置：「_*_*_*」形式のテーブルの最終行+3行目から出力
' 3. 4品番対応：58050FrLH, 58050RrRH, 28050FrLH, 28050RrRH
' 4. 日付順ソート：ロット番号を日付昇順で出力
' 5. 0埋め形式：ロット番号を4桁0埋め（例：12 → 0012）
' 6. 表形式出力：縦に並べて視覚的に分かりやすく
'
' 【出力形式】
' 行N:   [B]58050FrLH  [C]       [D]58050RrRH  [E]       [F]28050FrLH  [G]       [H]28050RrRH  [I]
' 行N+1: [B]ロット    [C]数量   [D]ロット    [E]数量   [F]ロット    [G]数量   [H]ロット    [I]数量
' 行N+2: [B]0012      [C]200    [D]0108      [E]450    [F]0025      [G]100    [H]0099      [I]300
' 行N+3: [B]0034      [C]150    [D]0201      [E]500    [F]0067      [G]250    [H]0123      [I]400
' ...
'
' 【テーブル構成】
' 期間テーブル : アクティブシート、「_集計期間」で始まる名前のテーブル
' 出力位置基準 : アクティブシート、「_*_*_*」形式のテーブル（最も行番号が大きいもの）
' ソーステーブル : シート「ロット数量」、テーブル「_ロット数量」
'
' 【処理フロー】
' 1. アクティブシートから「_集計期間」で始まるテーブルを検索
' 2. テーブルの1行目データから開始日・終了日を取得
' 3. 「_*_*_*」形式のテーブルを検索し、最も行番号が大きいテーブルを特定
' 4. そのテーブルの最終行+3行目を出力開始行とする
' 5. ロット数量テーブルから期間内・工程=加工のデータを抽出
' 6. 品番ごとに日付順ソート
' 7. 出力範囲（B:I列）をクリア
' 8. ヘッダー行、列名行、データ行を縦方向に出力
'
' 【集計条件】
' - ロット数量：工程=加工、品番2が対象4品番、日付が期間内
' ========================================

Sub 転記_ショット数詳細()
    ' ============================================
    ' 最適化設定の保存と適用
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
    Application.StatusBar = "ショット数詳細転記処理を開始します..."

    ' ============================================
    ' アクティブシートから集計期間テーブルを検索
    ' ============================================
    Dim wsActive As Worksheet
    Set wsActive = ActiveSheet

    Dim tblPeriod As ListObject
    Set tblPeriod = Nothing

    Dim tbl As ListObject
    For Each tbl In wsActive.ListObjects
        If tbl.Name Like "_集計期間*" Then
            Set tblPeriod = tbl
            Exit For
        End If
    Next tbl

    If tblPeriod Is Nothing Then
        MsgBox "アクティブシートに「_集計期間」で始まるテーブルが見つかりません。", vbCritical
        GoTo Cleanup
    End If

    If tblPeriod.DataBodyRange Is Nothing Then
        MsgBox "「" & tblPeriod.Name & "」テーブルにデータがありません。", vbCritical
        GoTo Cleanup
    End If

    ' ============================================
    ' 期間テーブルの1行目から開始日・終了日を取得
    ' ============================================
    Dim startDate As Date, endDate As Date
    On Error Resume Next
    startDate = CDate(tblPeriod.DataBodyRange.Cells(1, 2).Value) ' 開始日は2列目
    endDate = CDate(tblPeriod.DataBodyRange.Cells(1, 3).Value)   ' 終了日は3列目
    On Error GoTo ErrorHandler

    If startDate = 0 Or endDate = 0 Then
        MsgBox "期間テーブルの1行目から開始日・終了日を取得できませんでした。", vbCritical
        GoTo Cleanup
    End If

    Application.StatusBar = "期間: " & Format(startDate, "yyyy/mm/dd") & " - " & Format(endDate, "yyyy/mm/dd")

    ' ============================================
    ' 「_*_*_*」形式のテーブルを検索（1個のみ許可）
    ' ============================================
    Dim tblBase As ListObject
    Set tblBase = Nothing
    Dim tblCount As Long
    tblCount = 0

    For Each tbl In wsActive.ListObjects
        ' テーブル名が「_*_*_*」形式かチェック（アンダースコア3個で区切られている）
        Dim tblNameParts() As String
        tblNameParts = Split(tbl.Name, "_")

        If UBound(tblNameParts) >= 3 Then
            tblCount = tblCount + 1
            Set tblBase = tbl
        End If
    Next tbl

    ' テーブルが1個だけかチェック
    If tblCount <> 1 Then
        MsgBox "このマクロは_*_*_*形式のテーブルが1個だけの時のみ実行できます", vbCritical
        GoTo Cleanup
    End If

    ' 出力開始行：基準テーブルの最終行 + 3行
    Dim outputStartRow As Long
    outputStartRow = tblBase.Range.Row + tblBase.Range.Rows.Count - 1 + 3

    Application.StatusBar = "出力開始行: " & outputStartRow

    ' ============================================
    ' ロット数量テーブルの取得
    ' ============================================
    Dim wsLot As Worksheet
    On Error Resume Next
    Set wsLot = ThisWorkbook.Worksheets("ロット数量")
    On Error GoTo ErrorHandler

    If wsLot Is Nothing Then
        MsgBox "シート「ロット数量」が見つかりません。", vbCritical
        GoTo Cleanup
    End If

    Dim tblLot As ListObject
    On Error Resume Next
    Set tblLot = wsLot.ListObjects("_ロット数量")
    On Error GoTo ErrorHandler

    If tblLot Is Nothing Then
        MsgBox "シート「ロット数量」にテーブル「_ロット数量」が見つかりません。", vbCritical
        GoTo Cleanup
    End If

    Dim lotData As Range
    Set lotData = tblLot.DataBodyRange
    If lotData Is Nothing Then
        MsgBox "ロット数量テーブルにデータがありません。", vbCritical
        GoTo Cleanup
    End If

    ' ============================================
    ' ロット数量テーブルの列インデックス取得
    ' ============================================
    Dim colLotHizuke As Long, colLotKoutei As Long, colLotHinban2 As Long
    Dim colLotNumber As Long, colLotSuuryou As Long

    colLotHizuke = tblLot.ListColumns("日付").Index
    colLotKoutei = tblLot.ListColumns("工程").Index
    colLotHinban2 = tblLot.ListColumns("品番2").Index
    colLotNumber = tblLot.ListColumns("ロット").Index
    colLotSuuryou = tblLot.ListColumns("ロット数量").Index

    ' ============================================
    ' 対象品番の定義（LH系4品番のみ）
    ' 品番 → {ロット列番号, 数量列番号}のマッピング
    ' ============================================
    Dim targetHinban As Object
    Set targetHinban = CreateObject("Scripting.Dictionary")
    ' B列=2, C列=3, D列=4, E列=5, F列=6, G列=7, H列=8, I列=9
    targetHinban("58050FrLH") = Array(2, 3)   ' B列:ロット, C列:数量
    targetHinban("58050RrRH") = Array(4, 5)   ' D列:ロット, E列:数量
    targetHinban("28050FrLH") = Array(6, 7)   ' F列:ロット, G列:数量
    targetHinban("28050RrRH") = Array(8, 9)   ' H列:ロット, I列:数量

    ' ============================================
    ' 品番ごとのデータ格納用Dictionary
    ' ============================================
    Dim hinbanData As Object
    Set hinbanData = CreateObject("Scripting.Dictionary")

    Dim hinbanKey As Variant
    For Each hinbanKey In targetHinban.Keys
        Set hinbanData(CStr(hinbanKey)) = CreateObject("Scripting.Dictionary")
    Next hinbanKey

    ' ============================================
    ' データ抽出ループ
    ' ============================================
    Dim lotArr As Variant
    lotArr = lotData.Value

    Dim r As Long
    Dim lotDate As Variant, dt As Date
    Dim koutei As String, hinban2 As String
    Dim lotNumber As Variant, lotQty As Variant
    Dim lotKey As String

    For r = 1 To UBound(lotArr, 1)
        lotDate = lotArr(r, colLotHizuke)

        If IsDate(lotDate) Then
            dt = CDate(lotDate)

            If dt >= startDate And dt <= endDate Then
                koutei = Trim(CStr(lotArr(r, colLotKoutei)))

                If koutei = "加工" Then
                    hinban2 = Trim(CStr(lotArr(r, colLotHinban2)))

                    If targetHinban.Exists(hinban2) Then
                        lotNumber = lotArr(r, colLotNumber)
                        lotQty = lotArr(r, colLotSuuryou)

                        If Not IsEmpty(lotNumber) And IsNumeric(lotQty) Then
                            ' データを格納（キー：日付_連番、値：日付,ロット番号,数量）
                            lotKey = Format(dt, "yyyymmdd") & "_" & r
                            hinbanData(hinban2)(lotKey) = Array(dt, lotNumber, CDbl(lotQty))
                        End If
                    End If
                End If
            End If
        End If

        ' 進捗表示（200行ごと）
        If (r Mod 200) = 0 Then
            Application.StatusBar = "データ抽出中: " & r & "/" & UBound(lotArr, 1) & " 行処理中..."
        End If
    Next r

    ' ============================================
    ' 出力範囲のクリア（B:I列、出力開始行から下）
    ' ============================================
    Application.StatusBar = "出力範囲をクリア中..."
    
    Dim clearStartRow As Long
    clearStartRow = outputStartRow
    Dim clearEndRow As Long
    clearEndRow = wsActive.Cells(wsActive.Rows.Count, 2).End(xlUp).Row

    ' 出力開始行より下で最大500行程度をクリア対象とする
    If clearEndRow < clearStartRow Then
        clearEndRow = clearStartRow + 500
    ElseIf clearEndRow - clearStartRow > 500 Then
        clearEndRow = clearStartRow + 500
    End If

    If clearEndRow >= clearStartRow Then
        wsActive.Range(wsActive.Cells(clearStartRow, 2), wsActive.Cells(clearEndRow, 9)).ClearContents
    End If

    ' ============================================
    ' ヘッダー行の出力
    ' ============================================
    Application.StatusBar = "ヘッダーを出力中..."
    
    Dim headerRow As Long
    headerRow = outputStartRow

    wsActive.Cells(headerRow, 2).Value = "58050FrLH"  ' B列
    wsActive.Cells(headerRow, 4).Value = "58050RrRH"  ' D列
    wsActive.Cells(headerRow, 6).Value = "28050FrLH"  ' F列
    wsActive.Cells(headerRow, 8).Value = "28050RrRH"  ' H列

    ' ============================================
    ' 列名行の出力
    ' ============================================
    Dim columnRow As Long
    columnRow = headerRow + 1

    wsActive.Cells(columnRow, 2).Value = "ロット"  ' B列
    wsActive.Cells(columnRow, 3).Value = "数量"    ' C列
    wsActive.Cells(columnRow, 4).Value = "ロット"  ' D列
    wsActive.Cells(columnRow, 5).Value = "数量"    ' E列
    wsActive.Cells(columnRow, 6).Value = "ロット"  ' F列
    wsActive.Cells(columnRow, 7).Value = "数量"    ' G列
    wsActive.Cells(columnRow, 8).Value = "ロット"  ' H列
    wsActive.Cells(columnRow, 9).Value = "数量"    ' I列

    ' ============================================
    ' データ行の出力（品番ごとに処理）
    ' ============================================
    Application.StatusBar = "データを出力中..."

    Dim maxOutputRow As Long
    maxOutputRow = columnRow

    Dim hinbanOrder As Variant
    hinbanOrder = Array("58050FrLH", "58050RrRH", "28050FrLH", "28050RrRH")

    Dim hinbanName As Variant
    For Each hinbanName In hinbanOrder
        Dim currentHinban As String
        currentHinban = CStr(hinbanName)

        Dim dataDict As Object
        Set dataDict = hinbanData(currentHinban)

        If dataDict.Count > 0 Then
            ' Dictionaryのデータを配列化
            Dim dataArr() As Variant
            ReDim dataArr(1 To dataDict.Count, 1 To 3)

            Dim idx As Long
            idx = 1
            Dim dataKey As Variant
            For Each dataKey In dataDict.Keys
                Dim dataItem As Variant
                dataItem = dataDict(dataKey)
                dataArr(idx, 1) = dataItem(0) ' 日付
                dataArr(idx, 2) = dataItem(1) ' ロット番号
                dataArr(idx, 3) = dataItem(2) ' 数量
                idx = idx + 1
            Next dataKey

            ' 日付順にソート
            Call SortByDate(dataArr)

            ' 出力列番号を取得
            Dim colInfo As Variant
            colInfo = targetHinban(currentHinban)
            Dim lotCol As Long, qtyCol As Long
            lotCol = colInfo(0)
            qtyCol = colInfo(1)

            ' データを縦方向に出力
            Dim dataRowStart As Long
            dataRowStart = columnRow + 1

            Dim i As Long
            For i = 1 To UBound(dataArr, 1)
                Dim currentRow As Long
                currentRow = dataRowStart + i - 1

                ' ロット番号を4桁0埋め
                Dim lotNumStr As String
                If IsNumeric(dataArr(i, 2)) Then
                    Dim lotNumValue As Long
                    lotNumValue = CLng(dataArr(i, 2))
                    lotNumStr = Format(lotNumValue, "0000")
                Else
                    lotNumStr = CStr(dataArr(i, 2))
                End If

                ' セルに書き込み
                wsActive.Cells(currentRow, lotCol).Value = lotNumStr
                wsActive.Cells(currentRow, qtyCol).Value = CLng(dataArr(i, 3))

                ' 最終行を更新
                If currentRow > maxOutputRow Then
                    maxOutputRow = currentRow
                End If
            Next i
        End If
    Next hinbanName

    Application.StatusBar = False

    ' ============================================
    ' 印刷範囲の自動設定
    ' ============================================
    Dim printStartRow As Long, printEndRow As Long
    Dim printStartCol As Long, printEndCol As Long

    printStartRow = tblBase.Range.Row
    printEndRow = maxOutputRow
    printStartCol = tblBase.Range.Column
    printEndCol = tblBase.Range.Column + tblBase.Range.Columns.Count - 1

    ' 列番号をアルファベットに変換
    Dim printStartColStr As String, printEndColStr As String
    printStartColStr = ColumnNumberToLetter(printStartCol)
    printEndColStr = ColumnNumberToLetter(printEndCol)

    ' 印刷範囲を設定
    wsActive.PageSetup.PrintArea = printStartColStr & printStartRow & ":" & printEndColStr & printEndRow

Cleanup:
    ' ============================================
    ' 最適化設定の復元
    ' ============================================
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.DisplayAlerts = origDisplayAlerts
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.DisplayAlerts = origDisplayAlerts
    Application.StatusBar = False
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
End Sub

' ========================================
' 日付順ソートサブルーチン（バブルソート）
' ========================================
Private Sub SortByDate(ByRef arr() As Variant)
    Dim i As Long, j As Long
    Dim temp As Variant
    
    For i = 1 To UBound(arr, 1) - 1
        For j = i + 1 To UBound(arr, 1)
            ' 日付で昇順ソート
            If arr(i, 1) > arr(j, 1) Then
                ' 行全体を入れ替え
                temp = arr(i, 1)
                arr(i, 1) = arr(j, 1)
                arr(j, 1) = temp
                
                temp = arr(i, 2)
                arr(i, 2) = arr(j, 2)
                arr(j, 2) = temp
                
                temp = arr(i, 3)
                arr(i, 3) = arr(j, 3)
                arr(j, 3) = temp
            End If
        Next j
    Next i
End Sub

' ========================================
' 列番号をアルファベットに変換（A, B, ..., Z, AA, AB, ...）
' ========================================
Private Function ColumnNumberToLetter(ByVal colNum As Long) As String
    Dim dividend As Long
    Dim columnName As String
    Dim modulo As Long

    dividend = colNum
    Do While dividend > 0
        modulo = (dividend - 1) Mod 26
        columnName = Chr(65 + modulo) & columnName
        dividend = (dividend - modulo) \ 26
    Loop

    ColumnNumberToLetter = columnName
End Function
