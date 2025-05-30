Attribute VB_Name = "m日別集計_モールFR別"
Option Explicit

' モールF/R統合集計マクロ（画面更新抑制版）
' 「全工程」テーブルから「モール」工程のデータを日付・F/Rでグループ化して集計
Sub 日別集計_モールFR別()
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim tblSource As ListObject
    Dim tblOutput As ListObject
    Dim dict As Object 'Scripting.Dictionary
    Dim outputArray() As Variant
    Dim dataArray As Variant
    Dim sortKeys() As String ' ソート用のキー配列
    
    ' 対象品番リスト（モールF用）
    Dim targetHinbanListF As Variant
    targetHinbanListF = Array("58020F", "58030F", "58040F", "58050F", "58060F", "58830F", _
                            "58021", "58022", "58031", "58032", "58041", "58042", _
                            "58051", "58052", "58061", "58062", "47030F", "47030R", _
                            "47031", "47032", "47035", "47036", "58221F", "58223", "58224")
    
    ' 対象品番リスト（モールR用）
    Dim targetHinbanListR As Variant
    targetHinbanListR = Array("58020R", "58030R", "58040R", "58050R", "58060R", "58830R", _
                            "58025", "58026", "58035", "58036", "58045", "58046", _
                            "58055", "58056", "58065", "58066", "58221R", "58015", "58016")
    
    Dim sourceSheetName As String
    Dim sourceTableName As String
    Dim outputSheetName As String
    Dim outputTableName As String
    Dim outputStartCellAddress As String
    Dim outputHeader As Range
    
    Dim i As Long, r As Long, j As Long, k As Long
    Dim colDate As Long, colProcess As Long, colHinban As Long
    Dim colJisseki As Long, colDandori As Long, colKadou As Long, colFuryo As Long
    
    Dim currentDate As Date
    Dim currentFR As String
    Dim dictKey As String
    Dim jissekiVal As Double, dandoriVal As Double, kadouVal As Double, furyoVal As Double
    Dim item As Variant
    Dim key As Variant
    Dim tempKey As String
    Dim hinbanValue As String
    Dim isTargetHinban As Boolean
    
    ' 基本設定
    Set wb = ThisWorkbook
    sourceSheetName = "全工程"
    sourceTableName = "_全工程"
    outputSheetName = "モールFR別"
    outputTableName = "_モールFR別a"
    outputStartCellAddress = "A3"
    
    ' ステータスバー表示
    Application.StatusBar = "モールFR別集計を開始します..."
    
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
    colHinban = GetColumnIndex(tblSource, "品番")
    colJisseki = GetColumnIndex(tblSource, "実績")
    colDandori = GetColumnIndex(tblSource, "段取時間")
    colKadou = GetColumnIndex(tblSource, "稼働時間")
    colFuryo = GetColumnIndex(tblSource, "不良")
    
    If colDate = 0 Or colProcess = 0 Or colHinban = 0 Or colJisseki = 0 Or colDandori = 0 Or colKadou = 0 Or colFuryo = 0 Then
        MsgBox "「全工程」テーブルに必要な列（日付, 工程, 品番, 実績, 段取時間, 稼働時間, 不良）が見つかりません。列名を確認してください。", vbCritical
        GoTo Cleanup
    End If
    
    ' 3. データ集計 (Dictionary使用)
    Set dict = CreateObject("Scripting.Dictionary")
    dataArray = tblSource.DataBodyRange.Value2 ' 高速化のため配列で処理
    
    Application.StatusBar = "データを集計中..."
    
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        ' 「工程」列の値が「モール」を含むかチェック
        If InStr(1, CStr(dataArray(i, colProcess)), "モール", vbTextCompare) > 0 Then
            ' 「品番」列が対象リストに含まれるかチェック
            hinbanValue = CStr(dataArray(i, colHinban))
            isTargetHinban = False
            currentFR = ""
            
            ' モールF品番チェック
            For k = LBound(targetHinbanListF) To UBound(targetHinbanListF)
                If InStr(1, hinbanValue, targetHinbanListF(k), vbTextCompare) > 0 Then
                    isTargetHinban = True
                    currentFR = "F"
                    Exit For
                End If
            Next k
            
            ' モールR品番チェック（Fで見つからなかった場合）
            If Not isTargetHinban Then
                For k = LBound(targetHinbanListR) To UBound(targetHinbanListR)
                    If InStr(1, hinbanValue, targetHinbanListR(k), vbTextCompare) > 0 Then
                        isTargetHinban = True
                        currentFR = "R"
                        Exit For
                    End If
                Next k
            End If
            
            If isTargetHinban Then
                ' 日付の妥当性チェックと変換
                If IsDate(dataArray(i, colDate)) Then
                    currentDate = CDate(dataArray(i, colDate))
                ElseIf IsNumeric(dataArray(i, colDate)) Then
                    ' 数値の場合は日付シリアル値として扱う
                    currentDate = CDate(CLng(dataArray(i, colDate)))
                Else
                    ' 日付として認識できないデータはスキップ
                    Debug.Print "警告: 日付として認識できないデータがありました。行 " & i + tblSource.HeaderRowRange.row & ", 値: " & dataArray(i, colDate)
                    GoTo NextIteration
                End If
                
                ' 辞書キーの作成（日付|F/R）
                dictKey = Format(currentDate, "yyyy/mm/dd") & "|" & currentFR
                
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
        End If
NextIteration:
    Next i
    
    If dict.Count = 0 Then
        MsgBox "モール工程（指定品番）に該当するデータが集計されませんでした。", vbInformation
        ' その場合でもシートと空のテーブルは作成するように続行
    End If
    
    ' 4. 出力先シートの準備
    Application.StatusBar = "出力先シートを準備中..."
    
    On Error Resume Next
    Set wsOutput = wb.Sheets(outputSheetName)
    If wsOutput Is Nothing Then
        Set wsOutput = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        wsOutput.Name = outputSheetName
    End If
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
        ' ヘッダー行を設定（新規テーブルの場合のみ）：「F/R」列を追加
        outputHeader.Resize(1, 6).Value = Array("日付", "F/R", "実績", "不良", "稼働時間", "段取時間")
    End If
    
    ' 6. 集計結果を配列に変換し、日付・F/Rでソート
    If dict.Count > 0 Then
        Application.StatusBar = "データをソート中..."
        
        ' ソート用のキー配列を作成
        ReDim sortKeys(1 To dict.Count)
        i = 0
        For Each key In dict.Keys
            i = i + 1
            sortKeys(i) = CStr(key)
        Next key
        
        ' バブルソートで日付・F/R順に並び替え
        For i = 1 To dict.Count - 1
            For j = i + 1 To dict.Count
                If sortKeys(i) > sortKeys(j) Then
                    tempKey = sortKeys(i)
                    sortKeys(i) = sortKeys(j)
                    sortKeys(j) = tempKey
                End If
            Next j
        Next i
        
        ' 出力配列の作成：「F/R」列を追加して列数を6列
        ReDim outputArray(1 To dict.Count, 1 To 6)
        For r = 1 To dict.Count
            key = sortKeys(r)
            item = dict(key)
            
            ' キーを分解して日付とF/Rを取得
            Dim keyParts() As String
            keyParts = Split(key, "|")
            
            outputArray(r, 1) = CDate(keyParts(0))          '日付
            outputArray(r, 2) = keyParts(1)                 'F/R
            outputArray(r, 3) = item(0)                     '実績
            outputArray(r, 4) = item(1)                     '不良
            outputArray(r, 5) = item(2)                     '稼働時間
            outputArray(r, 6) = item(3)                     '段取時間
        Next r
        
        ' 7. データ出力
        Application.StatusBar = "データを出力中..."
        
        If Not isNewTable Then
            ' 既存テーブルの場合：テーブルサイズを調整してデータを挿入
            tblOutput.Resize outputHeader.Resize(UBound(outputArray, 1) + 1, 6) ' 列数を6列
        End If
        outputHeader.Offset(1, 0).Resize(UBound(outputArray, 1), UBound(outputArray, 2)).Value = outputArray
    End If
    
    ' 8. テーブル作成（新規の場合のみ）または更新
    If isNewTable Then
        ' データがない場合でもヘッダーのみのテーブルを作成
        Dim dataRangeForTable As Range
        If dict.Count > 0 Then
            Set dataRangeForTable = outputHeader.Resize(dict.Count + 1, 6) ' 列数を6列
        Else
            Set dataRangeForTable = outputHeader.Resize(1, 6) ' ヘッダーのみ、列数を6列
        End If
        
        Set tblOutput = wsOutput.ListObjects.Add(xlSrcRange, dataRangeForTable, , xlYes)
        tblOutput.Name = outputTableName
        tblOutput.TableStyle = "TableStyleMedium9"
    ElseIf dict.Count = 0 Then
        ' 既存テーブルでデータがない場合：ヘッダーのみにリサイズ
        tblOutput.Resize outputHeader.Resize(1, 6) ' 列数を6列
    End If
    
    ' テーブルのフィルターボタンを非表示に設定
    If Not tblOutput Is Nothing Then
      tblOutput.ShowAutoFilter = False
    End If
    
    ' 日付列の書式設定
    If dict.Count > 0 And Not tblOutput Is Nothing Then
        tblOutput.ListColumns("日付").DataBodyRange.NumberFormatLocal = "yyyy/mm/dd"
    End If
    
    ' ========== 追加書式設定 ==========
    Application.StatusBar = "書式設定中..."
    
    If Not tblOutput Is Nothing Then
        ' 1. データ範囲の「縮小して全体を表示する」設定
        If dict.Count > 0 Then
            tblOutput.DataBodyRange.ShrinkToFit = True
        End If
        
        ' 2. 全列の列幅を6.4に設定
        Dim col As ListColumn
        For Each col In tblOutput.ListColumns
            col.Range.ColumnWidth = 6.4
        Next col
        
        ' 3. 「稼働時間」「段取時間」列の書式：小数点以下2桁設定
        If dict.Count > 0 Then
            On Error Resume Next ' 列が存在しない場合のエラー無視
            tblOutput.ListColumns("稼働時間").DataBodyRange.NumberFormatLocal = "0.00"
            tblOutput.ListColumns("段取時間").DataBodyRange.NumberFormatLocal = "0.00"
            On Error GoTo ErrorHandler
        End If
    End If
    
    ' 4. A1セルにタイトルを設定
    With wsOutput.Range("A1")
        .Value = "モールFR別データ出力"
        .Font.Name = "Yu Gothic UI"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' 完了処理
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
    
    Application.StatusBar = False
End Sub

' テーブルの列名からインデックスを取得するヘルパー関数
Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    Dim col As ListColumn
    Dim i As Long
    i = 0
    On Error Resume Next
    i = tbl.ListColumns(columnName).Index
    On Error GoTo 0
    GetColumnIndex = i
End Function