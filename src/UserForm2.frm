VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2
   Caption         =   "入力支援"
   ClientHeight    =   3912
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8484.001
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' クラスハンドラ配列
Private zoneHandlers()     As CTextBoxEvent
Private numberHandlers()   As CTextBoxEvent
Private quantityHandlers() As CTextBoxEvent
Private returnHandlers()   As CTextBoxEvent  ' 戻めし用追加

' フォームレベル変数
Private currentRowCount As Integer
Private Const ROW_HEIGHT As Integer = 24
Private Const INITIAL_FORM_HEIGHT As Integer = 224
Private Const INITIAL_TEXTBOX4_HEIGHT As Integer = 33
Private Const MAX_ROWS As Integer = 10

Private Sub Label11_Click()

End Sub

Private Sub TextBox4_Change()

End Sub

'==============================================
' 初期化
Private Sub UserForm_Initialize()
    ' フォームサイズ・位置
    Me.Height = INITIAL_FORM_HEIGHT
    Me.Width = 435
    Me.Left = 0
    Me.Top = 0

    ' 行数リセット
    currentRowCount = 1

    ' クラスハンドラ配列初期化
    ReDim zoneHandlers(1 To MAX_ROWS)
    ReDim numberHandlers(1 To MAX_ROWS)
    ReDim quantityHandlers(1 To MAX_ROWS)
    ReDim returnHandlers(1 To MAX_ROWS)  ' 差戻し用追加

    ' ----- 静的コントロール初期設定 -----
    With ComboBox1
        .AddItem "成形": .AddItem "塗装": .AddItem "モール": .AddItem "加工"
        .IMEMode = fmIMEModeHiragana
        .TabIndex = 0
    End With

    With TextBox1
        .IMEMode = fmIMEModeDisable
        .TabIndex = 1
    End With

    With ComboBox2
        .IMEMode = fmIMEModeHiragana
        .TabIndex = 2
    End With

    ' ComboBox3（品番末尾）削除
    ' TextBox2（注番月）削除

    With TextBox3
        .IMEMode = fmIMEModeDisable
        .TabIndex = 3
    End With

    With TextBox4
        .Font.Name = "Yu Gothic UI": .Font.Size = 12: .Font.Bold = True
        .MultiLine = True: .IMEMode = fmIMEModeDisable
        .Height = INITIAL_TEXTBOX4_HEIGHT: .Width = 256
        .Left = 156: .Top = 114: .TabIndex = 12
    End With
    ' --------------------------------------

    ' 動的1行目を作成
    CreateRowControls 1

    ' ボタン・フォーム調整
    PositionButtons
    ResizeForm
    ResizeTextBox4
    UpdateButtonTabIndexes

    ' IMEモード初期設定
    SetJapaneseIME ComboBox1
    SetJapaneseIME ComboBox2

    ' 初期フォーカス
    ComboBox1.SetFocus
End Sub

'==============================================
' ユーザーフォームのキーダウンイベント
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' IME制御
    HandleFormKeyDown Me

    ' 移動チェック処理
    Dim ctrl As MSForms.Control
    Set ctrl = Me.ActiveControl

    If KeyCode = vbKeyTab Then
        If TypeOf ctrl Is MSForms.TextBox Then
            If ctrl.Tag = "Zone" Then
                ' ゾーン入力チェック
                Dim zoneCtrl As MSForms.TextBox
                Set zoneCtrl = ctrl
                If zoneCtrl.Value <> "" Then
                    If Not (zoneCtrl.Value Like "[A-E]") Then
                        MsgBox "ゾーンはA,B,C,D,Eのいずれかを入力してください。", vbExclamation
                        KeyCode = 0
                        zoneCtrl.Value = ""
                        zoneCtrl.SetFocus
                    End If
                End If
            ElseIf ctrl.Tag = "Number" Then
                ' 番号入力チェック
                Dim numCtrl As MSForms.TextBox
                Set numCtrl = ctrl
                If numCtrl.Value <> "" Then
                    If Not (IsNumeric(numCtrl.Value) Or numCtrl.Value Like "[A-Za-z]*") Then
                        MsgBox "番号は数字またはアルファベットのみ入力してください。", vbExclamation
                        KeyCode = 0
                        numCtrl.Value = ""
                        numCtrl.SetFocus
                    End If
                End If
            ElseIf ctrl.Tag = "Quantity" Then
                ' 数量入力チェック
                Dim qtyCtrl As MSForms.TextBox
                Set qtyCtrl = ctrl
                If qtyCtrl.Value <> "" Then
                    If Not (IsNumeric(qtyCtrl.Value) And Val(qtyCtrl.Value) = Int(Val(qtyCtrl.Value)) And Val(qtyCtrl.Value) > 0) Then
                        MsgBox "数量は正の整数を入力してください。", vbExclamation
                        KeyCode = 0
                        qtyCtrl.Value = ""
                        qtyCtrl.SetFocus
                    End If
                End If
            ElseIf ctrl.Tag = "Return" Then
                ' 差戻し入力チェック（0、1のみ）
                Dim retCtrl As MSForms.TextBox
                Set retCtrl = ctrl
                If retCtrl.Value <> "" Then
                    If retCtrl.Value <> "0" And retCtrl.Value <> "1" Then
                        MsgBox "差戻しは「0」または「1」を入力してください。", vbExclamation
                        KeyCode = 0
                        retCtrl.Value = "0"
                        retCtrl.SetFocus
                    End If
                End If

                ' 最終行の差戻し入力欄から入力→追加ボタンへフォーカス移動
                Dim rowNum As String
                rowNum = Replace(ctrl.Name, "TextBoxReturn_", "")
                If CInt(rowNum) = currentRowCount Then
                    KeyCode = 0
                    CommandButton4.SetFocus
                End If
            End If
        End If
    End If
End Sub

'==============================================
' 動的行コントロール作成
Private Sub CreateRowControls(ByVal rowNum As Integer)
    Dim topPos As Integer: topPos = 114 + (rowNum - 1) * ROW_HEIGHT
    Dim baseTab As Integer: baseTab = 4 + (rowNum - 1) * 4  ' タブ順調整

    Dim tbZone As MSForms.TextBox, tbNum As MSForms.TextBox, tbQty As MSForms.TextBox, tbRet As MSForms.TextBox
    Dim hZone As CTextBoxEvent, hNum As CTextBoxEvent, hQty As CTextBoxEvent, hRet As CTextBoxEvent

    ' --- ゾーン TextBox ---
    Set tbZone = Me.Controls.Add("Forms.TextBox.1", "TextBoxZone_" & rowNum, True)
    With tbZone
        .Left = 12: .Top = topPos: .Width = 30: .Height = 23
        .Font.Name = "Yu Gothic UI": .Font.Size = 12: .Font.Bold = True
        .IMEMode = fmIMEModeDisable: .Tag = "Zone": .TabIndex = baseTab
        .MaxLength = 1
    End With
    Set hZone = New CTextBoxEvent: Set hZone.TB = tbZone
    Set zoneHandlers(rowNum) = hZone

    ' IMEモードを半角英数に強制設定
    SetAlphaIME tbZone

    ' --- 番号 TextBox ---
    Set tbNum = Me.Controls.Add("Forms.TextBox.1", "TextBoxNum_" & rowNum, True)
    With tbNum
        .Left = 48: .Top = topPos: .Width = 30: .Height = 23
        .Font.Name = "Yu Gothic UI": .Font.Size = 12: .Font.Bold = True
        .IMEMode = fmIMEModeDisable: .Tag = "Number": .TabIndex = baseTab + 1
    End With
    Set hNum = New CTextBoxEvent: Set hNum.TB = tbNum
    Set numberHandlers(rowNum) = hNum

    ' IMEモードを半角英数に強制設定
    SetAlphaIME tbNum

    ' --- 数量 TextBox ---
    Set tbQty = Me.Controls.Add("Forms.TextBox.1", "TextBoxQty_" & rowNum, True)
    With tbQty
        .Left = 84: .Top = topPos: .Width = 30: .Height = 23
        .Font.Name = "Yu Gothic UI": .Font.Size = 12: .Font.Bold = True
        .IMEMode = fmIMEModeDisable: .Tag = "Quantity": .TabIndex = baseTab + 2
    End With
    Set hQty = New CTextBoxEvent: Set hQty.TB = tbQty
    Set quantityHandlers(rowNum) = hQty

    ' IMEモードを半角英数に強制設定
    SetAlphaIME tbQty

    ' --- 差戻し TextBox --- 新規追加
    Set tbRet = Me.Controls.Add("Forms.TextBox.1", "TextBoxReturn_" & rowNum, True)
    With tbRet
        .Left = 120: .Top = topPos: .Width = 30: .Height = 23
        .Font.Name = "Yu Gothic UI": .Font.Size = 12: .Font.Bold = True
        .IMEMode = fmIMEModeDisable: .Tag = "Return": .TabIndex = baseTab + 3
        .MaxLength = 1
    End With
    Set hRet = New CTextBoxEvent: Set hRet.TB = tbRet
    Set returnHandlers(rowNum) = hRet

    ' IMEモードを半角英数に強制設定
    SetAlphaIME tbRet
End Sub

'==============================================
' 追加ボタン (CommandButton4)
Private Sub CommandButton4_Click()
    If currentRowCount >= MAX_ROWS Then
        MsgBox "最大行数（" & MAX_ROWS & "行）に達しました。", vbExclamation
        Exit Sub
    End If
    currentRowCount = currentRowCount + 1
    CreateRowControls currentRowCount

    ' 修正: 順序的な処理で呼び出す
    PositionButtons  ' まずボタン位置を更新
    ResizeForm       ' 次にフォームサイズを更新
    ResizeTextBox4   ' 最後にTextBox4サイズを更新
    UpdateButtonTabIndexes

    ' フォームを強制的に再描画
    Me.Repaint

    Me.Controls("TextBoxZone_" & currentRowCount).SetFocus

    ' 新しい行のIMEモードを設定
    SetAlphaIME Me.Controls("TextBoxZone_" & currentRowCount)
End Sub

'==============================================
' 各種ボタンのタブ順更新
Private Sub UpdateButtonTabIndexes()
    Dim lastTab As Integer
    lastTab = 4 + currentRowCount * 4  ' タブ順調整
    CommandButton4.TabIndex = lastTab
    CommandButton1.TabIndex = lastTab + 1
    CommandButton2.TabIndex = lastTab + 2
    CommandButton3.TabIndex = lastTab + 3
End Sub

'==============================================
' ボタン位置調整
Private Sub PositionButtons()
    ' ボタン位置計算 - 重要なので詳細にコメント
    Dim topPos As Integer
    topPos = 162 + (currentRowCount - 1) * ROW_HEIGHT  ' 最初のボタン行の基準位置は162

    ' 各ボタンの位置を設定
    With CommandButton4  ' 追加ボタン
        .Left = 12
        .Top = topPos
        .Width = 54
        .Height = 24
    End With

    With CommandButton1  ' 展開ボタン
        .Left = 144
        .Top = topPos
        .Width = 54
        .Height = 24
    End With

    With CommandButton2  ' 転記ボタン
        .Left = 210
        .Top = topPos
        .Width = 54
        .Height = 24
    End With

    With CommandButton3  ' 終了ボタン
        .Left = 360
        .Top = topPos
        .Width = 54
        .Height = 24
    End With
End Sub

'==============================================
' フォーム高さ調整
Private Sub ResizeForm()
    ' ボタンの下端を基準にフォーム高さを計算（余白40px）
    Dim newHeight As Integer
    newHeight = CommandButton1.Top + CommandButton1.Height + 40

    ' 計算した高さをフォームに設定
    Me.Height = newHeight
End Sub

'==============================================
' TextBox4 高さ調整
Private Sub ResizeTextBox4()
    TextBox4.Height = INITIAL_TEXTBOX4_HEIGHT + (currentRowCount - 1) * ROW_HEIGHT
End Sub

'==============================================
' ComboBox1 GotFocus - IME制御追加
Private Sub ComboBox1_GotFocus()
    ComboBox1.IMEMode = fmIMEModeHiragana
    ' IMEモードを強制的に日本語入力ONへ
    SetJapaneseIME ComboBox1
End Sub

' ComboBox1のキーダウン時のIME維持
Private Sub ComboBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' CapsLockなどのキー入力時に日本語入力モードを維持
    SetJapaneseIME ComboBox1
End Sub

' ComboBox1のキー入力中のIME維持
Private Sub ComboBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' キー入力中も日本語入力モードを維持
    SetJapaneseIME ComboBox1
End Sub

'==============================================
' ComboBox2 GotFocus - IME制御追加
Private Sub ComboBox2_GotFocus()
    ComboBox2.IMEMode = fmIMEModeHiragana
    ' IMEモードを強制的に日本語入力ONへ
    SetJapaneseIME ComboBox2
End Sub

' ComboBox2のキーダウン時のIME維持
Private Sub ComboBox2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' CapsLockなどのキー入力時に日本語入力モードを維持
    SetJapaneseIME ComboBox2
End Sub

' ComboBox2のキー入力中のIME維持
Private Sub ComboBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' キー入力中も日本語入力モードを維持
    SetJapaneseIME ComboBox2
End Sub

'==============================================
' TextBox1 GotFocus - IME制御追加
Private Sub TextBox1_GotFocus()
    TextBox1.IMEMode = fmIMEModeDisable
    ' IMEモードを強制的に半角英数へ
    SetAlphaIME TextBox1
End Sub

'==============================================
' TextBox3 GotFocus - IME制御追加
Private Sub TextBox3_GotFocus()
    TextBox3.IMEMode = fmIMEModeDisable
    ' IMEモードを強制的に半角英数へ
    SetAlphaIME TextBox3
End Sub

'==============================================
' 展開ボタン (CommandButton1)
Private Sub CommandButton1_Click()
    If TextBox1.Value = "" Then
        MsgBox "日付を入力してください。", vbExclamation: TextBox1.SetFocus: Exit Sub
    End If
    Dim dateValue As String: dateValue = StrConv(TextBox1.Value, vbNarrow)
    If IsDate(dateValue) Then
        TextBox1.Value = Format(CDate(dateValue), "yyyy/m/d")
    Else
        MsgBox "日付は「yyyy/m/d」または「m/d」形式で入力してください。", vbExclamation
        TextBox1.SetFocus: Exit Sub
    End If
    If ComboBox1.Value = "" Then MsgBox "工程を選択してください。", vbExclamation: ComboBox1.SetFocus: Exit Sub
    If ComboBox2.Value = "" Then MsgBox "品番を選択してください。", vbExclamation: ComboBox2.SetFocus: Exit Sub
    ' ComboBox3（品番末尾）のバリデーション削除
    ' TextBox2（注番月）のバリデーション削除
    If TextBox3.Value = "" Then MsgBox "ロットを入力してください。", vbExclamation: TextBox3.SetFocus: Exit Sub

    Dim i As Integer, allRowsValid As Boolean: allRowsValid = True
    For i = 1 To currentRowCount
        With Me
            If .Controls("TextBoxZone_" & i).Value <> "" _
            Or .Controls("TextBoxNum_" & i).Value <> "" _
            Or .Controls("TextBoxQty_" & i).Value <> "" _
            Or .Controls("TextBoxReturn_" & i).Value <> "" Then

                If .Controls("TextBoxZone_" & i).Value = "" Then
                    MsgBox i & "行目のゾーンを入力してください。", vbExclamation
                    .Controls("TextBoxZone_" & i).SetFocus: allRowsValid = False: Exit For
                End If
                If .Controls("TextBoxNum_" & i).Value = "" Then
                    MsgBox i & "行目の番号を入力してください。", vbExclamation
                    .Controls("TextBoxNum_" & i).SetFocus: allRowsValid = False: Exit For
                End If
                If .Controls("TextBoxQty_" & i).Value = "" Then
                    MsgBox i & "行目の数量を入力してください。", vbExclamation
                    .Controls("TextBoxQty_" & i).SetFocus: allRowsValid = False: Exit For
                End If
                ' 差戻しは任意項目なのでチェックしない
            End If
        End With
    Next i
    If Not allRowsValid Then Exit Sub

    ' 進捗表示
    Application.StatusBar = "品番を展開中..."

    GeneratePartNumberExpansion

    ' ステータスバーをクリア
    Application.StatusBar = False
End Sub

'==============================================
' 品番展開ロジック
Private Sub GeneratePartNumberExpansion()
    Dim dateStr As String: dateStr = TextBox1.Value
    Dim itemCode As String: itemCode = ComboBox2.Value
    ' itemSuffix（品番末尾）削除
    ' monthStr（注番月）削除
    Dim lotStr As String: lotStr = TextBox3.Value
    Dim processSymbol As String
    Select Case ComboBox1.Value
    Case "成形": processSymbol = "S"
    Case "塗装": processSymbol = "T"
    Case "モール": processSymbol = "M"
    Case "加工": processSymbol = "K"
    End Select

    Dim expandedText As String: expandedText = ""
    Dim i As Integer
    For i = 1 To currentRowCount
        With Me
            If .Controls("TextBoxZone_" & i).Value <> "" _
            And .Controls("TextBoxNum_" & i).Value <> "" _
            And .Controls("TextBoxQty_" & i).Value <> "" Then
                If expandedText <> "" Then expandedText = expandedText & vbNewLine
                ' 品番末尾、注番月を削除した展開形式
                expandedText = expandedText _
                    & dateStr & "-" & itemCode & "-" _
                    & lotStr & "-" & processSymbol & "-" _
                    & .Controls("TextBoxZone_" & i).Value & "-" _
                    & .Controls("TextBoxNum_" & i).Value & "-" _
                    & .Controls("TextBoxQty_" & i).Value & "-" _
                    & .Controls("TextBoxReturn_" & i).Value  ' 差戻し追加
            End If
        End With
    Next i

    TextBox4.Value = expandedText
    CommandButton2.SetFocus
End Sub

'==============================================
' 転記ボタン (CommandButton2)
Private Sub CommandButton2_Click()
    If TextBox4.Value = "" Then
        MsgBox "先に展開ボタンを押して品番を展開してください。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "「_不良集計ゾーン別S」テーブルを検索中..."

    Dim tbl As ListObject, ws As Worksheet
    Dim tableFound As Boolean: tableFound = False

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            If tbl.Name = "_不良集計ゾーン別S" Then
                tableFound = True: Exit For
            End If
        Next tbl
        If tableFound Then Exit For
    Next ws
    On Error GoTo 0

    If Not tableFound Then
        MsgBox "「_不良集計ゾーン別S」テーブルが見つかりません。", vbCritical
        GoTo CleanExit
    End If

    ' 最初の空白行を取り出す
    Dim dateCol As Range: Set dateCol = tbl.ListColumns("日付").DataBodyRange
    Dim firstEmptyRow As Long: firstEmptyRow = 0

    If Not dateCol Is Nothing Then
        Dim arr As Variant: arr = dateCol.Value
        Dim i As Long

        If Not IsEmpty(arr) Then
            For i = 1 To UBound(arr, 1)
                If IsEmpty(arr(i, 1)) Or arr(i, 1) = "" Then
                    firstEmptyRow = i: Exit For
                End If
            Next i
        End If
    End If

    If firstEmptyRow = 0 Then firstEmptyRow = tbl.ListRows.Count + 1

    ' 必要行数を追加
    Dim dataRowCount As Long: dataRowCount = 0
    For i = 1 To currentRowCount
        With Me
            If .Controls("TextBoxZone_" & i).Value <> "" _
            And .Controls("TextBoxNum_" & i).Value <> "" _
            And .Controls("TextBoxQty_" & i).Value <> "" Then
                dataRowCount = dataRowCount + 1
            End If
        End With
    Next i
    If dataRowCount = 0 Then
        MsgBox "転記するデータがありません。", vbExclamation
        GoTo CleanExit
    End If

    Dim currentRows As Long: currentRows = tbl.ListRows.Count
    If firstEmptyRow + dataRowCount - 1 > currentRows Then
        Dim rowsToAdd As Long: rowsToAdd = (firstEmptyRow + dataRowCount - 1) - currentRows
        For i = 1 To rowsToAdd: tbl.ListRows.Add: Next i
    End If

    ' データ転記
    Application.StatusBar = "データを転記中..."
    Dim rowIndex As Long: rowIndex = firstEmptyRow
    For i = 1 To currentRowCount
        Dim zVal As String: zVal = Me.Controls("TextBoxZone_" & i).Value
        Dim nVal As String: nVal = Me.Controls("TextBoxNum_" & i).Value
        Dim qVal As String: qVal = Me.Controls("TextBoxQty_" & i).Value
        ' 差戻し値をInteger型で取得（0または1）
        Dim rVal As Integer
        rVal = IIf(Me.Controls("TextBoxReturn_" & i).Value = "1", 1, 0)

        If zVal <> "" And nVal <> "" And qVal <> "" Then
            With tbl.ListRows(rowIndex)
                .Range.Cells(1, tbl.ListColumns("日付").Index).Value = TextBox1.Value
                .Range.Cells(1, tbl.ListColumns("品番").Index).Value = ComboBox2.Value
                ' 品番末尾の転記削除
                ' 注番月の転記削除
                .Range.Cells(1, tbl.ListColumns("ロット").Index).Value = TextBox3.Value
                .Range.Cells(1, tbl.ListColumns("発見").Index).Value = _
                    IIf(ComboBox1.Value = "成形", "S", IIf(ComboBox1.Value = "塗装", "T", _
                    IIf(ComboBox1.Value = "モール", "M", "K")))
                .Range.Cells(1, tbl.ListColumns("ゾーン").Index).Value = zVal
                .Range.Cells(1, tbl.ListColumns("番号").Index).Value = nVal
                .Range.Cells(1, tbl.ListColumns("数量").Index).Value = qVal
                .Range.Cells(1, tbl.ListColumns("差戻し").Index).Value = rVal  ' 差戻しを数値(0/1)で転記
            End With
            rowIndex = rowIndex + 1
        End If
    Next i

    ' 入力欄をリセット（工程と日付を保持）
    ClearInputFields

    ws.Activate
    tbl.ListColumns("日付").DataBodyRange.Cells(rowIndex - 1, 1).Select
    Application.StatusBar = "転記完了"
    Application.Run "ステータスバークリア"

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Application.StatusBar = "エラー: " & Err.Description
        Application.Run "ステータスバークリア"
    End If
End Sub

'==============================================
' 終了ボタン (CommandButton3)
Private Sub CommandButton3_Click()
    Unload Me
End Sub

'==============================================
' ClearInputFields（入力欄をリセット - 工程と日付を保持）
Private Sub ClearInputFields()
    ' 静的コントロールクリア（ComboBox1=工程とTextBox1=日付は保持）
    ' TextBox2（注番月）削除済み
    TextBox3.Value = ""    ' ロット
    TextBox4.Value = ""    ' 品番展開
    ComboBox2.Clear        ' 品番
    ' ComboBox3（品番末尾）削除済み

    ' 動的行（2行目以降）削除
    Dim i As Integer
    For i = currentRowCount To 2 Step -1
        Me.Controls.Remove "TextBoxZone_" & i
        Me.Controls.Remove "TextBoxNum_" & i
        Me.Controls.Remove "TextBoxQty_" & i
        Me.Controls.Remove "TextBoxReturn_" & i  ' 差戻し削除
        Set zoneHandlers(i) = Nothing
        Set numberHandlers(i) = Nothing
        Set quantityHandlers(i) = Nothing
        Set returnHandlers(i) = Nothing  ' 差戻しハンドラ削除
    Next i

    ' 1行目クリア
    Me.Controls("TextBoxZone_1").Value = ""
    Me.Controls("TextBoxNum_1").Value = ""
    Me.Controls("TextBoxQty_1").Value = ""
    Me.Controls("TextBoxReturn_1").Value = ""  ' 差戻しクリア

    ' 行数リセット
    currentRowCount = 1

    ' 再調整
    PositionButtons
    ResizeForm
    ResizeTextBox4
    UpdateButtonTabIndexes

    ' フォームを再描画して確実に更新
    Me.Repaint

    ' 工程選択に応じて品番コンボボックスを初期化
    If ComboBox1.Value <> "" Then
        ComboBox1_Change
    End If

    ' TextBox1（日付）にフォーカスを設定
    TextBox1.SetFocus

    ' IMEモードを設定（日付入力モード）
    SetAlphaIME TextBox1
End Sub

'==============================================
' ComboBox1 Change & Exit
Private Sub ComboBox1_Change()
    ComboBox2.Clear
    ' ComboBox3（品番末尾）削除済み
    Select Case ComboBox1.Value
    Case "成形", "塗装", "加工"
        ComboBox2.AddItem "ノアFrLH": ComboBox2.AddItem "ノアFrRH"
        ComboBox2.AddItem "ノアRrLH": ComboBox2.AddItem "ノアRrRH"
        ComboBox2.AddItem "アルFrLH": ComboBox2.AddItem "アルFrRH"
        ComboBox2.AddItem "アルRrLH": ComboBox2.AddItem "アルRrRH"
    Case "モール"
        ComboBox2.AddItem "アルFrLH": ComboBox2.AddItem "アルFrRH"
        ComboBox2.AddItem "アルRrLH": ComboBox2.AddItem "アルRrRH"
    End Select
End Sub

Private Sub ComboBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If ComboBox1.Value = "" Then Exit Sub
    If Not (ComboBox1.Value = "成形" Or ComboBox1.Value = "塗装" _
         Or ComboBox1.Value = "モール" Or ComboBox1.Value = "加工") Then
        MsgBox "工程は「成形」「塗装」「モール」「加工」から選択してください。", vbExclamation
        Cancel = True: ComboBox1.SetFocus
    End If
End Sub

'==============================================
' ComboBox2 Exit（品番末尾関連の処理削除）
Private Sub ComboBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If ComboBox2.Value = "" Or ComboBox1.Value = "" Then Exit Sub
    Dim valid As Boolean: valid = False
    Select Case ComboBox1.Value
    Case "成形", "塗装", "加工"
        Select Case ComboBox2.Value
        Case "ノアFrLH", "ノアFrRH", "ノアRrLH", "ノアRrRH", _
             "アルFrLH", "アルFrRH", "アルRrLH", "アルRrRH"
            valid = True
        End Select
    Case "モール"
        Select Case ComboBox2.Value
        Case "アルFrLH", "アルFrRH", "アルRrLH", "アルRrRH"
            valid = True
        End Select
    End Select
    If Not valid Then
        MsgBox "無効な品番です。リストから選択してください。", vbExclamation
        Cancel = True: ComboBox2.SetFocus
    End If
End Sub

' ComboBox2_Change削除（品番末尾の連動処理不要）
' ComboBox3_Exit削除（品番末尾削除済み）

'==============================================
' TextBox1 Exit (日付)
Private Sub TextBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBox1.Value = "" Then Exit Sub
    Dim d As String: d = StrConv(TextBox1.Value, vbNarrow)
    If IsDate(d) Then
        TextBox1.Value = Format(CDate(d), "yyyy/m/d")
    Else
        MsgBox "日付は「yyyy/m/d」または「m/d」形式で入力してください。", vbExclamation
        TextBox1.Value = "": Cancel = True: TextBox1.SetFocus
    End If
End Sub

' TextBox2_Exit削除（注番月削除済み）

'==============================================
' TextBox3 Exit (ロット)
Private Sub TextBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBox3.Value = "" Then Exit Sub
    Dim v As String: v = StrConv(TextBox3.Value, vbNarrow)
    If Not IsNumeric(v) Or Val(v) <> Int(Val(v)) Or Val(v) <= 0 Then
        MsgBox "ロットは正の整数を入力してください。", vbExclamation
        TextBox3.Value = "": Cancel = True: TextBox3.SetFocus
    Else
        TextBox3.Value = v
    End If
End Sub
