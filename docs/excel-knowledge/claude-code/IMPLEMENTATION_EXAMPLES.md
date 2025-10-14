# VBAマクロ実装例集

実際にコピペして使えるコード集。基本的な使い方は[QUICK_REFERENCE.md](QUICK_REFERENCE.md)、
詳細な背景説明は[メインナレッジベース](EXCEL_MACRO_KNOWLEDGE_BASE.md)を参照。

## 📋 目次

1. [テンプレート](#templates)
2. [辞書化パターン](#dictionary-patterns)
3. [配列処理](#array-processing)
4. [エラーハンドリング](#error-handling)
5. [テーブル操作](#table-operations)
6. [ユーザーフォーム](#userform-examples)

---

## <a name="templates"></a>1. テンプレート

### 標準版テンプレート（複雑な処理用）

```vba
Option Explicit

' ========================================
' マクロ名: m処理名_詳細名
' 処理概要: [1行で説明]
' ソース: シート「○○」テーブル「××」
' ========================================

Sub OptimizedMacroTemplate()
    ' 最適化設定の保存
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    Dim origEnableEvents As Boolean
    Dim origDisplayAlerts As Boolean
    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation
    origEnableEvents = Application.EnableEvents
    origDisplayAlerts = Application.DisplayAlerts
    
    ' 最適化設定（これが最重要）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' ステータスバー初期化
    Application.StatusBar = "処理を開始します..."
    
    ' =================================
    ' メイン処理をここに記述
    ' 注意：Activateは絶対に使わない！
    ' =================================
    
    ' 処理完了のステータスバー表示
    Application.StatusBar = "処理が完了しました"
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
           "詳細: " & errDesc, vbCritical, "エラー"
    
Cleanup:
    ' 設定を確実に復元
    Application.StatusBar = False
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.DisplayAlerts = origDisplayAlerts
End Sub
```

### CommandButton用テンプレート（一括実行）

```vba
Private Sub CommandButton1_Click()
    ' CommandButtonレベルで設定管理
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    ' 複数マクロの順次実行
    Call マクロ1_データ準備
    Call マクロ2_メイン処理
    Call マクロ3_後処理
    
    ' 設定復元
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "処理中にエラーが発生しました: " & Err.Description, vbCritical
End Sub
```

---

## <a name="dictionary-patterns"></a>2. 辞書化パターン

### 基本的な辞書処理

```vba
' グループ別集計の例
Sub DictionaryGroupingExample()
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    ' 辞書オブジェクトの作成
    Dim groupDict As Object
    Set groupDict = CreateObject("Scripting.Dictionary")
    
    ' データ範囲の取得
    Dim dataRange As Range
    Set dataRange = ActiveSheet.Range("A2:C100")
    
    ' グループ化処理
    Dim i As Long
    Dim groupKey As String
    Dim value As Double
    
    For i = 1 To dataRange.Rows.Count
        groupKey = dataRange.Cells(i, 1).Value  ' A列をキー
        value = dataRange.Cells(i, 3).Value     ' C列を値
        
        If groupDict.Exists(groupKey) Then
            groupDict(groupKey) = groupDict(groupKey) + value
        Else
            groupDict(groupKey) = value
        End If
    Next i
    
    ' 結果出力
    Dim outputRow As Long
    outputRow = 2
    
    Dim key As Variant
    For Each key In groupDict.Keys
        Cells(outputRow, 5).Value = key
        Cells(outputRow, 6).Value = groupDict(key)
        outputRow = outputRow + 1
    Next key
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub
```

### 複数キーでの辞書管理

```vba
' 製番_品番_工程でグループ化する例
Function CreateGroupKey(seibanNo As String, hinbanNo As String, koutei As String) As String
    CreateGroupKey = seibanNo & "_" & hinbanNo & "_" & koutei
End Function

' 使用例
Dim complexKey As String
complexKey = CreateGroupKey(ws.Cells(i, 2).Value, ws.Cells(i, 3).Value, ws.Cells(i, 5).Value)

If Not groupDict.Exists(complexKey) Then
    Set groupDict(complexKey) = CreateObject("Scripting.Dictionary")
    groupDict(complexKey)("Count") = 0
    groupDict(complexKey)("Sum") = 0
End If

groupDict(complexKey)("Count") = groupDict(complexKey)("Count") + 1
groupDict(complexKey)("Sum") = groupDict(complexKey)("Sum") + cellValue
```

---

## <a name="array-processing"></a>3. 配列処理

### 範囲を配列に読み込んで高速処理

```vba
Sub ArrayProcessingExample()
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    ' データ範囲を配列に読み込み
    Dim dataArray As Variant
    dataArray = ActiveSheet.Range("A1:E1000").Value
    
    ' 配列内で処理（高速）
    Dim i As Long, j As Long
    For i = 1 To UBound(dataArray, 1)
        For j = 1 To UBound(dataArray, 2)
            ' 例：空白を0に変換
            If IsEmpty(dataArray(i, j)) Then
                dataArray(i, j) = 0
            End If
        Next j
    Next i
    
    ' 結果を一括書き戻し
    ActiveSheet.Range("A1:E1000").Value = dataArray
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub
```

---

## <a name="error-handling"></a>4. エラーハンドリング

### 詳細なエラー情報取得

```vba
Sub DetailedErrorHandling()
    On Error GoTo ErrorHandler
    
    ' メイン処理
    
    Exit Sub
    
ErrorHandler:
    Dim errNum As Long, errDesc As String, errSource As String
    errNum = Err.Number
    errDesc = Err.Description
    errSource = Err.Source
    
    ' エラーログ出力（イミディエイトウィンドウ）
    Debug.Print "=== エラー発生 ==="
    Debug.Print "発生時刻: " & Now
    Debug.Print "エラー番号: " & errNum
    Debug.Print "エラー内容: " & errDesc
    Debug.Print "エラー元: " & errSource
    Debug.Print "=================="
    
    ' ユーザーへの通知
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & vbCrLf & _
           "エラー番号: " & errNum & vbCrLf & _
           "詳細: " & errDesc & vbCrLf & vbCrLf & _
           "詳細はイミディエイトウィンドウを確認してください。", _
           vbCritical, "エラー"
    
    ' 設定の復元を忘れずに
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
```

---

## <a name="table-operations"></a>5. テーブル操作

### 安全なテーブル削除と再作成

```vba
Sub SafeTableRecreation()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False  ' 削除確認ダイアログ抑制
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("データ")
    
    ' 既存テーブルの完全削除
    On Error Resume Next
    Dim existingTable As ListObject
    Set existingTable = Nothing
    Set existingTable = ws.ListObjects("テーブル名")
    
    If Not existingTable Is Nothing Then
        existingTable.Unlist              ' テーブル形式解除
        existingTable.Range.Clear         ' 範囲の完全クリア
    End If
    Err.Clear
    On Error GoTo ErrorHandler
    
    ' 新規テーブル作成
    Dim newRange As Range
    Set newRange = ws.Range("A1:E100")
    
    Dim newTable As ListObject
    Set newTable = ws.ListObjects.Add(xlSrcRange, newRange, , xlYes)
    newTable.Name = "テーブル名"
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "テーブル操作エラー: " & Err.Description, vbCritical
End Sub
```

---

## <a name="userform-examples"></a>6. ユーザーフォーム

### エラー表示フォームの基本構造

```vba
' === ユーザーフォームモジュール (frmErrorDisplay) ===

Private Sub UserForm_Initialize()
    ' ListBox設定
    With lstErrors
        .ColumnCount = 3
        .ColumnWidths = "80;400;100"  ' 行番号｜エラー内容｜種別
        .ColumnHeads = True
    End With
    
    ' エラーデータの読み込み
    LoadErrorData
End Sub

Private Sub btnGoTo_Click()
    Dim selectedIndex As Long
    selectedIndex = lstErrors.ListIndex
    
    If selectedIndex >= 0 Then
        ' 行番号抽出
        Dim rowNum As Long
        rowNum = ExtractRowNumber(lstErrors.List(selectedIndex, 1))
        
        ' ジャンプ（Activate使わない）
        Application.Goto Worksheets("sysdata").Cells(rowNum, 1), True
        
        Me.Hide
    End If
End Sub

Private Sub lstErrors_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call btnGoTo_Click
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

' === 呼び出し側のマクロ ===
Sub ShowErrorDialog()
    ' フォームをモーダル表示
    frmErrorDisplay.Show vbModal
End Sub
```

### 進捗表示付き処理

```vba
Sub LongProcessWithProgress()
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    Dim totalRows As Long
    totalRows = 10000
    
    Dim i As Long
    For i = 1 To totalRows
        ' 100行ごとに進捗更新
        If i Mod 100 = 0 Then
            Application.StatusBar = "処理中... " & Format(i / totalRows, "0%") & _
                                   " (" & i & "/" & totalRows & ")"
            DoEvents  ' 画面更新を許可
        End If
        
        ' メイン処理
        ' ...
    Next i
    
    ' 完了表示
    Application.StatusBar = "処理完了 - " & totalRows & "行を処理しました"
    Application.Wait Now + TimeValue("00:00:02")
    Application.StatusBar = False
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub
```

---

## 💡 使用上の注意

1. **コピペ前に必ず確認**
   - シート名、テーブル名を実際のものに変更
   - 範囲指定を適切に調整
   - 不要な設定は削除

2. **パフォーマンスの考慮**
   - 小規模データ（1000行以下）→ シンプルな処理でOK
   - 大規模データ（1万行以上）→ 配列処理・辞書化を検討

3. **エラーハンドリング**
   - 基本的なエラー処理は必須
   - 詳細ログは開発時のみ使用

詳細な解説は[メインナレッジベース](EXCEL_MACRO_KNOWLEDGE_BASE.md)を参照。