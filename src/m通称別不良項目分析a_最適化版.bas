Attribute VB_Name = "m通称別不良項目分析a_最適化版"
Sub 通称別不良項目分析a_最適化版()
    ' 「品番別aa」シートの「_品番別aa」テーブルから通称別に
    ' 不良項目ベスト3と残り項目合計のテーブルを作成するマクロ
    
    ' === 画面更新の抑制（最重要！）===
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' ステータスバーに処理状況を表示
    Application.StatusBar = "通称別不良項目分析a: 処理を開始します..."
    
    ' 変数宣言
    Dim srcSheet As Worksheet, srcTable As ListObject
    Dim srcData As Variant, headerRow As Range
    Dim dictTsusho As Object, dictData As Object
    Dim i As Long, j As Long, k As Long
    Dim tsusho As Variant, key As Variant
    Dim currentRow As Long, outputRow As Long
    Dim tableRange As Range, newTable As ListObject
    Dim tableName As String
    
    ' 不良項目列名配列（21項目）
    Dim furyoItems As Variant
    furyoItems = Array("黒出線", "ショート", "ウエルド", "シミ", "異物", "シルバー", _
                       "フローマーク", "ゴミ混入", "GCカス", "キズ", "ヒケ", "練込み", _
                       "充填不良", "マクレ", "突出不良", "離型剤跡", "コアカス", "その他", _
                       "チョコ状黒出線", "汚れ", "噴出不良")
    
    ' 列インデックス用変数
    Dim tsushoCol As Integer
    Dim furyoColIdx As Object  ' Dictionary for 不良項目のインデックス
    
    ' ソート用変数
    Dim sortData() As Variant
    Dim tempName As Variant, tempValue As Variant
    
    ' ベスト3と残り項目用変数
    Dim best3Names(0 To 2) As String
    Dim best3Values(0 To 2) As Double
    Dim remainingNames As String
    Dim remainingValue As Double
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' シートとテーブルの取得（オブジェクト参照で直接操作）
    On Error Resume Next
    Set srcSheet = ThisWorkbook.Worksheets("品番別aa")
    On Error GoTo 0
    
    If srcSheet Is Nothing Then
        GoTo CleanExit
    End If
    
    On Error Resume Next
    Set srcTable = srcSheet.ListObjects("_品番別aa")
    On Error GoTo 0
    
    If srcTable Is Nothing Then
        GoTo CleanExit
    End If
    
    ' ステータスバー更新
    Application.StatusBar = "通称別不良項目分析a: データ取得中..."
    
    ' データの取得（配列処理で高速化）
    srcData = srcTable.DataBodyRange.Value
    Set headerRow = srcTable.HeaderRowRange
    
    ' Dictionary オブジェクトの作成
    Set furyoColIdx = CreateObject("Scripting.Dictionary")
    
    ' 列インデックスの取得
    For i = 1 To headerRow.Cells.Count
        Dim colName As String
        colName = CStr(headerRow.Cells(1, i).Value)
        
        If colName = "通称" Then
            tsushoCol = i
        End If
        
        ' 不良項目の列インデックスを記録
        For j = 0 To UBound(furyoItems)
            If colName = furyoItems(j) Then
                furyoColIdx.Add furyoItems(j), i
                Exit For
            End If
        Next j
    Next i
    
    ' 必要な列が見つからない場合は処理中止
    If tsushoCol = 0 Then
        GoTo CleanExit
    End If
    
    ' ステータスバー更新
    Application.StatusBar = "通称別不良項目分析a: 通称別グループ化中..."
    
    ' 通称別にデータをグループ化
    Set dictTsusho = CreateObject("Scripting.Dictionary")
    
    For i = 1 To UBound(srcData, 1)
        tsusho = srcData(i, tsushoCol)
        
        If Not dictTsusho.Exists(tsusho) Then
            Set dictData = CreateObject("Scripting.Dictionary")
            
            ' 不良項目の初期化
            For j = 0 To UBound(furyoItems)
                dictData.Add furyoItems(j), 0
            Next j
            
            dictTsusho.Add tsusho, dictData
        End If
        
        ' データの集計
        Set dictData = dictTsusho(tsusho)
        
        ' 不良項目の集計
        For j = 0 To UBound(furyoItems)
            If furyoColIdx.Exists(furyoItems(j)) Then
                Dim colIdx As Integer
                colIdx = furyoColIdx(furyoItems(j))
                If IsNumeric(srcData(i, colIdx)) Then
                    dictData(furyoItems(j)) = dictData(furyoItems(j)) + CDbl(srcData(i, colIdx))
                End If
            End If
        Next j
    Next i
    
    ' ステータスバー更新
    Application.StatusBar = "通称別不良項目分析a: テーブル作成中..."
    
    ' 出力開始位置を取得（最終行から3行空ける）
    currentRow = srcTable.Range.Row + srcTable.Range.Rows.Count + 3
    
    ' === 一括データ作成のための配列準備 ===
    Dim outputData() As Variant
    Dim outputIndex As Long
    ReDim outputData(1 To dictTsusho.Count * 5, 1 To 5) ' 各通称につき最大5行
    outputIndex = 0
    
    ' 各通称に対してテーブルを作成
    For Each tsusho In dictTsusho.Keys
        Set dictData = dictTsusho(tsusho)
        
        ' 不良項目のソート用配列を作成
        ReDim sortData(0 To UBound(furyoItems), 0 To 1)
        
        For j = 0 To UBound(furyoItems)
            sortData(j, 0) = furyoItems(j)  ' 項目名
            sortData(j, 1) = dictData(furyoItems(j))  ' 値
        Next j
        
        ' 値の大きい順にソート（バブルソート）
        For j = 0 To UBound(furyoItems) - 1
            For k = j + 1 To UBound(furyoItems)
                If CDbl(sortData(j, 1)) < CDbl(sortData(k, 1)) Then
                    ' 項目名の交換
                    tempName = sortData(j, 0)
                    sortData(j, 0) = sortData(k, 0)
                    sortData(k, 0) = tempName
                    
                    ' 値の交換
                    tempValue = sortData(j, 1)
                    sortData(j, 1) = sortData(k, 1)
                    sortData(k, 1) = tempValue
                End If
            Next k
        Next j
        
        ' ベスト3を取得
        For j = 0 To 2
            best3Names(j) = CStr(sortData(j, 0))
            best3Values(j) = CDbl(sortData(j, 1))
        Next j
        
        ' 残り18項目の名前連結と値合計（値がゼロ以外）
        remainingNames = ""
        remainingValue = 0
        
        For j = 3 To UBound(furyoItems)
            If CDbl(sortData(j, 1)) <> 0 Then
                If remainingNames <> "" Then
                    remainingNames = remainingNames & "|"
                End If
                remainingNames = remainingNames & CStr(sortData(j, 0))
            End If
            remainingValue = remainingValue + CDbl(sortData(j, 1))
        Next j
        
        ' 残り項目がない場合のデフォルト値
        If remainingNames = "" Then
            remainingNames = "その他"
        End If
        
        ' === オブジェクト参照による直接書き込み（Activateを使わない）===
        outputRow = currentRow
        
        ' ヘッダー行の作成
        With srcSheet
            .Cells(outputRow, 1).Value = "通称"
            .Cells(outputRow, 2).Value = best3Names(0)
            .Cells(outputRow, 3).Value = best3Names(1)
            .Cells(outputRow, 4).Value = best3Names(2)
            .Cells(outputRow, 5).Value = remainingNames
            
            ' データ行の作成
            outputRow = outputRow + 1
            .Cells(outputRow, 1).Value = CStr(tsusho)
            .Cells(outputRow, 2).Value = best3Values(0)
            .Cells(outputRow, 3).Value = best3Values(1)
            .Cells(outputRow, 4).Value = best3Values(2)
            .Cells(outputRow, 5).Value = remainingValue
        End With
        
        ' テーブルの作成
        Set tableRange = srcSheet.Range(srcSheet.Cells(currentRow, 1), _
                                      srcSheet.Cells(outputRow, 5))
        
        tableName = "_" & CStr(tsusho) & "aa"
        
        ' 既存の同名テーブルがある場合は削除
        On Error Resume Next
        If Not srcSheet.ListObjects(tableName) Is Nothing Then
            srcSheet.ListObjects(tableName).Delete
        End If
        On Error GoTo 0
        
        ' 新しいテーブルを作成
        Set newTable = srcSheet.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        newTable.Name = tableName
        newTable.ShowAutoFilter = False
        
        ' テーブルの書式設定（範囲指定で一括設定）
        With tableRange
            .Font.Name = "Yu Gothic UI"
            .Font.Size = 11
            .ShrinkToFit = True
        End With
        
        ' ヘッダー行の書式設定
        With srcSheet.Range(srcSheet.Cells(currentRow, 1), _
                           srcSheet.Cells(currentRow, 5))
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .ShrinkToFit = True
        End With
        
        ' データ行の数値フォーマット設定（整数表示）
        With srcSheet.Range(srcSheet.Cells(outputRow, 2), _
                           srcSheet.Cells(outputRow, 5))
            .NumberFormat = "0"
            .ShrinkToFit = True
        End With
        
        ' 0の値を薄いグレーにする条件付き書式
        With srcSheet.Range(srcSheet.Cells(outputRow, 2), _
                           srcSheet.Cells(outputRow, 5))
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="0"
            .FormatConditions(1).Font.Color = RGB(192, 192, 192)
        End With
        
        ' 列幅設定（一括設定で高速化）
        srcSheet.Columns(1).ColumnWidth = 14  ' A列
        srcSheet.Columns(2).ColumnWidth = 7   ' B列
        srcSheet.Columns(3).ColumnWidth = 7   ' C列
        srcSheet.Columns(4).ColumnWidth = 7   ' D列
        srcSheet.Columns(5).ColumnWidth = 7   ' E列
        
        ' 次のテーブルの位置設定（2行空ける）
        currentRow = outputRow + 3
    Next tsusho
    
    ' 処理完了
    Application.StatusBar = "通称別不良項目分析a: 処理が完了しました。"
    
    ' 1秒待機してステータスバークリア
    Application.Wait Now + TimeValue("00:00:01")
    
CleanExit:
    ' === 画面更新設定を必ず元に戻す ===
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
    Resume CleanExit
End Sub