Attribute VB_Name = "m日別集計_TG作業者別"
Option Explicit

' TG作業者別集計マクロ（画面ちらつき防止版）
' 「加工3」のデータを日付・作業者でグループ化して集計
Sub 日別集計_TG作業者別()
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim tblSource As ListObject
    Dim tblOutput As ListObject
    Dim dict As Object 'Scripting.Dictionary
    Dim outputArray() As Variant
    Dim dataArray As Variant
    Dim sortKeys() As String ' ソート用のキー配列
    
    Dim sourceSheetName As String
    Dim sourceTableName As String
    Dim outputSheetName As String
    Dim outputTableName As String
    Dim outputStartCellAddress As String
    Dim outputHeader As Range
    
    Dim i As Long, r As Long, j As Long, k As Long
    Dim colDate As Long, colProcess As Long, colWorker As Long
    Dim colJisseki As Long, colDandori As Long, colKadou As Long, colFuryo As Long
    
    Dim currentDate As Date
    Dim currentWorker As String
    Dim dictKey As String
    Dim jissekiVal As Double, dandoriVal As Double, kadouVal As Double, furyoVal As Double
    Dim item As Variant
    Dim key As Variant
    Dim tempKey As String
    
    ' 基本設定
    Set wb = ThisWorkbook
    sourceSheetName = "全工程"
    sourceTableName = "_全工程"
    outputSheetName = "TG作業者別"
    outputTableName = "_TG作業者別a"
    outputStartCellAddress = "A3"
    
    ' ステータスバー表示
    Application.StatusBar = "TG作業者別集計を開始します..."
    
    ' ★★★ 画面更新設定は削除（CommandButtonで一括管理） ★★★
    ' Application.ScreenUpdating = False ← 削除
    ' Application.Calculation = xlCalculationManual ← 削除
    ' Application.DisplayAlerts = False ← 削除
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' 1. 入力元シート・テーブルの存在確認と取得
    On Error Resume Next
    Set wsSource = wb.Sheets(sourceSheetName)
    If wsSource Is Nothing Then
        MsgBox "シート「" & sourceSheetName & "」が見つかりません。", vbCritical
        GoTo Cleanup
    End If
    
    Set tblSource = wsSource.ListObjects(sourceTableName)
    If tblSource Is Nothing Then
        MsgBox "テーブル「" & sourceTableName & "」がシート「" & sourceSheetName & "」に見つかりません。", vbCritical
        GoTo Cleanup
    End If
    On Error GoTo ErrorHandler
    
    ' データがない場合は終了
    If tblSource.DataBodyRange Is Nothing Then
        MsgBox "テーブル「" & sourceTableName & "」にデータがありません。", vbInformation
        GoTo Cleanup
    End If
    
    ' 2. 「全工程」テーブルの列インデックス取得
    colDate = GetColumnIndex(tblSource, "日付")
    colProcess = GetColumnIndex(tblSource, "工程")
    colWorker = GetColumnIndex(tblSource, "作業者")
    colJisseki = GetColumnIndex(tblSource, "実績")
    colDandori = GetColumnIndex(tblSource, "段取時間")
    colKadou = GetColumnIndex(tblSource, "稼働時間")
    colFuryo = GetColumnIndex(tblSource, "不良")
    
    If colDate = 0 Or colProcess = 0 Or colWorker = 0 Or colJisseki = 0 Or colDandori = 0 Or colKadou = 0 Or colFuryo = 0 Then
        MsgBox "「全工程」テーブルに必要な列（日付, 工程, 作業者, 実績, 段取時間, 稼働時間, 不良）が見つかりません。列名を確認してください。", vbCritical
        GoTo Cleanup
    End If
    
    ' 3. データ集計 (Dictionaryを使用)
    Set dict = CreateObject("Scripting.Dictionary")
    dataArray = tblSource.DataBodyRange.Value2 ' 高速化のため配列で処理
    
    Application.StatusBar = "データを集計中..."
    
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        ' 「工程」列の値が「加工3」と完全一致する行を抽出
        If CStr(dataArray(i, colProcess)) = "加工3" Then
            ' 日付の妥当性チェックと変換
            If IsDate(dataArray(i, colDate)) Then
                currentDate = CDate(dataArray(i, colDate))
            ElseIf IsNumeric(dataArray(i, colDate)) Then
                ' 数値の場合は日付シリアル値として扱う
                currentDate = CDate(CLng(dataArray(i, colDate)))
            Else
                ' 日付として認識できないデータはスキップ
                Debug.Print "警告: 日付として認識できないデータが見つかりました。行 " & i + tblSource.HeaderRowRange.row & ", 値: " & dataArray(i, colDate)
                GoTo NextIteration
            End If
            
            ' 作業者の取得
            currentWorker = CStr(dataArray(i, colWorker))
            
            ' 複合キーの作成（日付|作業者）
            dictKey = Format(currentDate, "yyyy/mm/dd") & "|" & currentWorker
            
            jissekiVal = val(dataArray(i, colJisseki))
            dandoriVal = val(dataArray(i, colDandori))
            kadouVal = val(dataArray(i, colKadou))
            furyoVal = val(dataArray(i, colFuryo))
            
            If dict.Exists(dictKey) Then
                item = dict(dictKey)
                item(0) = item(0) + jissekiVal '実績
                item(1) = item(1) + furyoVal  '不良
                item(2) = item(2) + kadouVal  '稼働時間
                item(3) = item(3) + dandoriVal '段取時間
                dict(dictKey) = item
            Else
                ReDim newItem(0 To 3) As Double
                newItem(0) = jissekiVal
                newItem(1) = furyoVal
                newItem(2) = kadouVal
                newItem(3) = dandoriVal
                dict.Add dictKey, newItem
            End If
        End If
NextIteration:
    Next i
    
    If dict.Count = 0 Then
        MsgBox "工程「加工3」に該当するデータが集計されませんでした。", vbInformation
        ' この場合でもシートと空のテーブルは作成されるようにする
    End If
    
    ' 4. 出力先シートの準備
    Application.StatusBar = "出力先シートを準備中..."
    
    On Error Resume Next
    Set wsOutput = wb.Sheets(outputSheetName)
    If wsOutput Is Nothing Then
        Set wsOutput = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsOutput.Name = outputSheetName
    End If
    ' ★★★ wsOutput.Activate を削除（画面切り替えの原因） ★★★
    On Error GoTo ErrorHandler
    
    ' 5. 出力先テーブルの準備
    Set outputHeader = wsOutput.Range(outputStartCellAddress)
    
    ' 既存テーブルがあるかチェック
    On Error Resume Next
    Set tblOutput = wsOutput.ListObjects(outputTableName)
    On Error GoTo ErrorHandler
    
    Dim isNewTable As Boolean
    isNewTable = (tblOutput Is Nothing)
    
    If Not isNewTable Then
        ' 既存テーブルの場合：データ部分のみクリア
        On Error Resume Next
        If Not tblOutput.DataBodyRange Is Nothing Then
            tblOutput.DataBodyRange.ClearContents
        End If
        On Error GoTo ErrorHandler
    Else
        ' ヘッダー書き込み（新規テーブルの場合のみ）
        outputHeader.Resize(1, 6).Value = Array("日付", "作業者", "実績", "不良", "稼働時間", "段取時間")
    End If
    
    ' 6. 集計結果を配列に変換し、日付・作業者でソート
    If dict.Count > 0 Then
        Application.StatusBar = "データをソート中..."
        
        ' ソート用のキー配列を作成
        ReDim sortKeys(1 To dict.Count)
        i = 0
        For Each key In dict.Keys
            i = i + 1
            sortKeys(i) = CStr(key)
        Next key
        
        ' バブルソートで日付・作業者順に並べ替え
        ' （実務ではQuickSortなどを使うべきだが、ここでは簡潔にバブルソート）
        For i = 1 To dict.Count - 1
            For j = i + 1 To dict.Count
                If sortKeys(i) > sortKeys(j) Then
                    tempKey = sortKeys(i)
                    sortKeys(i) = sortKeys(j)
                    sortKeys(j) = tempKey
                End If
            Next j
        Next i
        
        ' 出力配列の作成
        ReDim outputArray(1 To dict.Count, 1 To 6)
        For r = 1 To dict.Count
            key = sortKeys(r)
            item = dict(key)
            
            ' キーを分解して日付と作業者を取得
            Dim keyParts() As String
            keyParts = Split(key, "|")
            
            outputArray(r, 1) = CDate(keyParts(0)) '日付
            outputArray(r, 2) = keyParts(1)        '作業者
            outputArray(r, 3) = item(0)            '実績
            outputArray(r, 4) = item(1)            '不良
            outputArray(r, 5) = item(2)            '稼働時間
            outputArray(r, 6) = item(3)            '段取時間
        Next r
        
        ' 7. データ出力
        Application.StatusBar = "データを出力中..."
        
        If Not isNewTable Then
            ' 既存テーブルの場合：テーブルサイズ調整後にデータ書き込み
            tblOutput.Resize outputHeader.Resize(UBound(outputArray, 1) + 1, 6)
        End If
        outputHeader.Offset(1, 0).Resize(UBound(outputArray, 1), UBound(outputArray, 2)).Value = outputArray
    End If
    
    ' 8. テーブル作成（新規の場合のみ）または更新
    If isNewTable Then
        ' データがない場合でもヘッダーのみのテーブルを作成
        Dim dataRangeForTable As Range
        If dict.Count > 0 Then
            Set dataRangeForTable = outputHeader.Resize(dict.Count + 1, 6)
        Else
            Set dataRangeForTable = outputHeader.Resize(1, 6) ' ヘッダーのみ
        End If
        
        Set tblOutput = wsOutput.ListObjects.Add(xlSrcRange, dataRangeForTable, , xlYes)
        tblOutput.Name = outputTableName
        tblOutput.TableStyle = "TableStyleMedium9"
    ElseIf dict.Count = 0 Then
        ' 既存テーブルでデータがない場合：ヘッダーのみにリサイズ
        tblOutput.Resize outputHeader.Resize(1, 6)
    End If
    
    ' テーブルのフィルターボタンを非表示に設定
    tblOutput.ShowAutoFilter = False
    
    ' 日付列の書式設定
    If dict.Count > 0 Then
        tblOutput.ListColumns("日付").DataBodyRange.NumberFormatLocal = "yyyy/mm/dd"
    End If
    
    ' ========== 追加書式設定 ==========
    Application.StatusBar = "書式を設定中..."
    
    ' 1. データ範囲の「縮小して全体を表示する」設定
    If dict.Count > 0 Then
        tblOutput.DataBodyRange.ShrinkToFit = True
    End If
    
    ' 2. 全列の列幅を6.4に設定
    Dim col As ListColumn
    For Each col In tblOutput.ListColumns
        col.Range.ColumnWidth = 6.4
    Next col
    
    ' 3. 「稼働時間」「段取時間」列の小数点以下2桁設定
    If dict.Count > 0 Then
        On Error Resume Next
        tblOutput.ListColumns("稼働時間").DataBodyRange.NumberFormatLocal = "0.00"
        tblOutput.ListColumns("段取時間").DataBodyRange.NumberFormatLocal = "0.00"
        On Error GoTo ErrorHandler
    End If
    
    ' 4. A1セルにタイトルを設定
    With wsOutput.Range("A1")
        .Value = "TG作業者別データ抽出"
        .Font.Name = "Yu Gothic UI"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' 処理完了
    Application.StatusBar = False
    
    Exit Sub
    
ErrorHandler:
    ' エラー時の処理
    Application.StatusBar = False
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
    
Cleanup:
    ' 後処理
    Set dict = Nothing
    Set wsSource = Nothing
    Set wsOutput = Nothing
    Set tblSource = Nothing
    Set tblOutput = Nothing
    Set wb = Nothing
    
    ' ★★★ 画面更新設定を戻す処理も削除（CommandButtonで一括管理） ★★★
    ' Application.ScreenUpdating = True ← 削除
    ' Application.Calculation = xlCalculationAutomatic ← 削除
    ' Application.DisplayAlerts = True ← 削除
    Application.StatusBar = False
End Sub

' テーブルの列名から列インデックスを取得するヘルパー関数
Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    Dim col As ListColumn
    Dim i As Long
    i = 0
    On Error Resume Next
    i = tbl.ListColumns(columnName).Index
    On Error GoTo 0
    GetColumnIndex = i
End Function

