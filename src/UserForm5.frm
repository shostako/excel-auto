VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5
   Caption         =   "データ修正・削除"
   ClientHeight    =   4000
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9000
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ========================================
' フォーム名: UserForm5
' 処理概要: Accessデータベースのレコードを修正・削除するためのユーザーフォーム
' ターゲットテーブル: Accessデータベース「不良調査表DB-{年}.accdb」テーブル「_不良集計ゾーン別」
' 作成日: 2026-01-20
' 更新日: 2026-01-20（入力制限・フォント設定・ID窪み対応）
' ========================================

' 定数定義
Private Const MAX_RECORDS As Integer = 10
Private Const ROW_HEIGHT As Integer = 24
Private Const HEADER_TOP As Integer = 100
Private Const DATA_START_TOP As Integer = 114
Private Const BUTTON_MARGIN As Integer = 15

' 列位置定数（ヘッダーとデータで共通使用）
Private Const COL_ID As Integer = 12
Private Const COL_DATE As Integer = 60
Private Const COL_ITEM As Integer = 134
Private Const COL_LOT As Integer = 213
Private Const COL_FIND As Integer = 253
Private Const COL_ZONE As Integer = 293
Private Const COL_NUM As Integer = 333
Private Const COL_QTY As Integer = 373
Private Const COL_RET As Integer = 413

' DBパス設定（mゾーン別データ転送ADOと同じ）
Private Const DB_BASE_PATH As String = "Z:\全社共有\オート事業部\日報\不良集計\不良集計表\"
Private Const DB_FILE_PREFIX As String = "不良調査表DB-"

' フォームレベル変数
Private currentRecordCount As Integer
Private recordIDs() As Long  ' 検索結果のID配列
Private headerCreated As Boolean  ' ヘッダー生成フラグ

' 外部参照用フラグ（修正または削除が実行されたか）
Public DataModified As Boolean

' CTextBoxEventクラスハンドラ配列（入力制御用）
Private dateHandlers() As CTextBoxEvent
Private lotHandlers() As CTextBoxEvent
Private zoneHandlers() As CTextBoxEvent
Private numberHandlers() As CTextBoxEvent
Private quantityHandlers() As CTextBoxEvent
Private returnHandlers() As CTextBoxEvent

' ============================================
' 初期化
' ============================================
Private Sub UserForm_Initialize()
    ' フォームサイズ初期設定
    Me.Width = 500
    Me.Height = 180

    ' 年コンボボックス初期化（過去5年〜現在）
    Dim i As Integer
    Dim currentYear As Integer
    currentYear = Year(Date)

    ComboBoxYear.Clear
    For i = currentYear - 5 To currentYear
        ComboBoxYear.AddItem i
    Next i
    ComboBoxYear.Value = currentYear

    ' トグルボタン初期化（修正モード）
    ToggleButtonMode.Value = False
    ToggleButtonMode.Caption = "修正"
    With ToggleButtonMode.Font
        .Name = "Yu Gothic UI": .Size = 12: .Bold = True
    End With

    ' ID入力欄の設定
    TextBoxIDs.IMEMode = fmIMEModeDisable  ' 半角英数モード

    ' タブ順設定（静的コントロール）
    ComboBoxYear.TabIndex = 0
    ToggleButtonMode.TabIndex = 1
    TextBoxIDs.TabIndex = 2
    CommandButtonSearch.TabIndex = 3
    ' 動的コントロールは4から始まる（CreateRecordRowで設定）
    ' 実行・閉じるボタンは動的コントロールの後（ResizeFormで設定）

    ' コマンドボタンのフォント設定
    With CommandButtonSearch.Font
        .Name = "Yu Gothic UI": .Size = 12: .Bold = True
    End With
    With CommandButtonExecute.Font
        .Name = "Yu Gothic UI": .Size = 12: .Bold = True
    End With
    With CommandButtonClose.Font
        .Name = "Yu Gothic UI": .Size = 12: .Bold = True
    End With

    ' レコード数初期化
    currentRecordCount = 0
    headerCreated = False
    DataModified = False  ' 修正・削除フラグ初期化

    ' クラスハンドラ配列初期化
    ReDim dateHandlers(1 To MAX_RECORDS)
    ReDim lotHandlers(1 To MAX_RECORDS)
    ReDim zoneHandlers(1 To MAX_RECORDS)
    ReDim numberHandlers(1 To MAX_RECORDS)
    ReDim quantityHandlers(1 To MAX_RECORDS)
    ReDim returnHandlers(1 To MAX_RECORDS)

    ' ボタン状態初期化
    CommandButtonExecute.Enabled = False
End Sub

' ============================================
' トグルボタン：モード切替
' ============================================
Private Sub ToggleButtonMode_Click()
    If ToggleButtonMode.Value Then
        ToggleButtonMode.Caption = "削除"
    Else
        ToggleButtonMode.Caption = "修正"
    End If
End Sub

' ============================================
' 検索ボタン：IDでAccessからデータ取得
' ============================================
Private Sub CommandButtonSearch_Click()
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim parsedIDs As Variant
    Dim errMsg As String
    Dim i As Integer
    Dim yearValue As Integer
    Dim dbPath As String
    Dim validIDList As String

    ' 入力チェック
    If ComboBoxYear.Value = "" Then
        MsgBox "年を選択してください。", vbExclamation
        ComboBoxYear.SetFocus
        Exit Sub
    End If

    If Trim(TextBoxIDs.Value) = "" Then
        MsgBox "IDを入力してください。", vbExclamation
        TextBoxIDs.SetFocus
        Exit Sub
    End If

    ' 既存の動的コントロールをクリア
    ClearDynamicControls

    ' ID解析（カンマ区切り＋範囲指定対応）
    parsedIDs = ParseIDInput(TextBoxIDs.Value, errMsg)

    ' エラーチェック
    If errMsg <> "" Then
        MsgBox errMsg, vbExclamation
        TextBoxIDs.SetFocus
        Exit Sub
    End If

    ' 結果がない場合
    If Not IsArray(parsedIDs) Then
        MsgBox "有効なIDがありません。", vbExclamation
        TextBoxIDs.SetFocus
        Exit Sub
    End If

    ' 空配列チェック（UBoundが-1の場合）
    On Error Resume Next
    Dim idCount As Long
    idCount = UBound(parsedIDs) - LBound(parsedIDs) + 1
    If Err.Number <> 0 Or idCount = 0 Then
        On Error GoTo 0
        MsgBox "有効なIDがありません。", vbExclamation
        TextBoxIDs.SetFocus
        Exit Sub
    End If
    On Error GoTo 0

    ' recordIDs配列にコピーし、SQL用のリスト作成
    ReDim recordIDs(LBound(parsedIDs) To UBound(parsedIDs))
    validIDList = ""
    For i = LBound(parsedIDs) To UBound(parsedIDs)
        recordIDs(i) = parsedIDs(i)
        If validIDList <> "" Then validIDList = validIDList & ","
        validIDList = validIDList & recordIDs(i)
    Next i

    ' DBパス構築
    yearValue = CInt(ComboBoxYear.Value)
    dbPath = DB_BASE_PATH & yearValue & "年\" & DB_FILE_PREFIX & yearValue & ".accdb"

    ' DBファイル存在チェック
    If Dir(dbPath) = "" Then
        MsgBox yearValue & "年のデータベースが見つかりません。" & vbCrLf & dbPath, vbExclamation
        Exit Sub
    End If

    ' ステータス表示
    Application.StatusBar = "データを検索中..."

    On Error GoTo ErrorHandler

    ' ADO接続
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"

    ' SQL実行
    sql = "SELECT ID, 日付, 品番, ロット, 発見, ゾーン, 番号, 数量, 差戻し " & _
          "FROM [_不良集計ゾーン別] " & _
          "WHERE ID IN (" & validIDList & ") " & _
          "ORDER BY ID"

    Set rs = conn.Execute(sql)

    ' 結果をフォームに表示
    currentRecordCount = 0

    Do While Not rs.EOF
        ' 最初の1件目でヘッダー生成
        If currentRecordCount = 0 Then
            CreateHeaderLabels
        End If
        currentRecordCount = currentRecordCount + 1
        CreateRecordRow currentRecordCount, rs
        rs.MoveNext
    Loop

    rs.Close
    conn.Close

    ' 結果チェック
    If currentRecordCount = 0 Then
        MsgBox "指定されたIDのデータが見つかりませんでした。", vbInformation
        CommandButtonExecute.Enabled = False
    Else
        ' フォームサイズ調整
        ResizeForm
        CommandButtonExecute.Enabled = True

        ' 見つからなかったIDを報告
        If currentRecordCount < idCount Then
            MsgBox currentRecordCount & "件のデータが見つかりました。" & vbCrLf & _
                   "（入力: " & idCount & "件）", vbInformation
        End If

        ' 検索後は上部コントロールを無効化（誤操作防止）
        ToggleButtonMode.Enabled = False
        TextBoxIDs.Enabled = False
        CommandButtonSearch.Enabled = False

        ' モードによってフォーカス先と編集可否を切り替え
        If ToggleButtonMode.Value Then
            ' 削除モード：データ編集不可、実行ボタンにフォーカス
            LockDataControls True
            CommandButtonExecute.SetFocus
        Else
            ' 修正モード：1行目の日付にフォーカス
            Me.Controls("TextBoxDate_1").SetFocus
        End If
    End If

    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
End Sub

' ============================================
' ヘッダーラベル動的生成
' ============================================
Private Sub CreateHeaderLabels()
    Dim lbl As MSForms.Label
    Const LABEL_INDENT As Integer = 4  ' ラベル用インデント（半角分）

    ' ID
    Set lbl = Me.Controls.Add("Forms.Label.1", "LabelHdr_ID", True)
    With lbl
        .Left = COL_ID + LABEL_INDENT: .Top = HEADER_TOP: .Width = 45: .Height = 16
        .Caption = "ID"
        .Font.Name = "Meiryo UI": .Font.Size = 9: .Font.Bold = True
    End With

    ' 日付
    Set lbl = Me.Controls.Add("Forms.Label.1", "LabelHdr_Date", True)
    With lbl
        .Left = COL_DATE + LABEL_INDENT: .Top = HEADER_TOP: .Width = 70: .Height = 16
        .Caption = "日付"
        .Font.Name = "Meiryo UI": .Font.Size = 9: .Font.Bold = True
    End With

    ' 品番
    Set lbl = Me.Controls.Add("Forms.Label.1", "LabelHdr_Item", True)
    With lbl
        .Left = COL_ITEM + LABEL_INDENT: .Top = HEADER_TOP: .Width = 75: .Height = 16
        .Caption = "品番"
        .Font.Name = "Meiryo UI": .Font.Size = 9: .Font.Bold = True
    End With

    ' ロット
    Set lbl = Me.Controls.Add("Forms.Label.1", "LabelHdr_Lot", True)
    With lbl
        .Left = COL_LOT + LABEL_INDENT: .Top = HEADER_TOP: .Width = 36: .Height = 16
        .Caption = "ロット"
        .Font.Name = "Meiryo UI": .Font.Size = 9: .Font.Bold = True
    End With

    ' 発見
    Set lbl = Me.Controls.Add("Forms.Label.1", "LabelHdr_Find", True)
    With lbl
        .Left = COL_FIND + LABEL_INDENT: .Top = HEADER_TOP: .Width = 36: .Height = 16
        .Caption = "発見"
        .Font.Name = "Meiryo UI": .Font.Size = 9: .Font.Bold = True
    End With

    ' ゾーン
    Set lbl = Me.Controls.Add("Forms.Label.1", "LabelHdr_Zone", True)
    With lbl
        .Left = COL_ZONE + LABEL_INDENT: .Top = HEADER_TOP: .Width = 36: .Height = 16
        .Caption = "ゾーン"
        .Font.Name = "Meiryo UI": .Font.Size = 9: .Font.Bold = True
    End With

    ' 番号
    Set lbl = Me.Controls.Add("Forms.Label.1", "LabelHdr_Num", True)
    With lbl
        .Left = COL_NUM + LABEL_INDENT: .Top = HEADER_TOP: .Width = 36: .Height = 16
        .Caption = "番号"
        .Font.Name = "Meiryo UI": .Font.Size = 9: .Font.Bold = True
    End With

    ' 数量
    Set lbl = Me.Controls.Add("Forms.Label.1", "LabelHdr_Qty", True)
    With lbl
        .Left = COL_QTY + LABEL_INDENT: .Top = HEADER_TOP: .Width = 36: .Height = 16
        .Caption = "数量"
        .Font.Name = "Meiryo UI": .Font.Size = 9: .Font.Bold = True
    End With

    ' 差戻
    Set lbl = Me.Controls.Add("Forms.Label.1", "LabelHdr_Ret", True)
    With lbl
        .Left = COL_RET + LABEL_INDENT: .Top = HEADER_TOP: .Width = 30: .Height = 16
        .Caption = "差戻"
        .Font.Name = "Meiryo UI": .Font.Size = 9: .Font.Bold = True
    End With

    headerCreated = True
End Sub

' ============================================
' 動的コントロール生成：1レコード分
' ============================================
Private Sub CreateRecordRow(rowNum As Integer, rs As Object)
    Dim topPos As Integer
    Dim baseTab As Integer
    topPos = DATA_START_TOP + (rowNum - 1) * ROW_HEIGHT
    baseTab = 4 + (rowNum - 1) * 8  ' 各行8コントロール（日付〜差戻）

    ' ID（読取専用TextBox：窪み付き）
    Dim tbID As MSForms.TextBox
    Set tbID = Me.Controls.Add("Forms.TextBox.1", "TextBoxID_" & rowNum, True)
    With tbID
        .Left = COL_ID: .Top = topPos: .Width = 45: .Height = 20
        .Value = rs("ID").Value
        .Font.Name = "Meiryo UI"
        .Font.Size = 9
        .Locked = True                         ' 編集不可
        .TabStop = False                       ' タブ移動対象外
        .SpecialEffect = fmSpecialEffectSunken ' 窪み
        .BackColor = &H8000000F                ' システムボタンフェイス色
    End With

    ' 日付
    Dim tbDate As MSForms.TextBox
    Dim hDate As CTextBoxEvent
    Set tbDate = Me.Controls.Add("Forms.TextBox.1", "TextBoxDate_" & rowNum, True)
    With tbDate
        .Left = COL_DATE: .Top = topPos: .Width = 70: .Height = 20
        .Font.Name = "Meiryo UI"
        .Font.Size = 9
        .Tag = "Date"
        .IMEMode = fmIMEModeDisable            ' 半角英数
        .MaxLength = 10                        ' yyyy/m/d
        .TabIndex = baseTab + 0
        If IsDate(rs("日付").Value) Then
            .Value = Format(rs("日付").Value, "yyyy/m/d")
        End If
    End With
    Set hDate = New CTextBoxEvent: Set hDate.TB = tbDate
    Set dateHandlers(rowNum) = hDate
    SetAlphaIME tbDate

    ' 品番
    Dim cbItem As MSForms.ComboBox
    Set cbItem = Me.Controls.Add("Forms.ComboBox.1", "ComboBoxItem_" & rowNum, True)
    With cbItem
        .Left = COL_ITEM: .Top = topPos: .Width = 75: .Height = 20
        .Font.Name = "Meiryo UI"
        .Font.Size = 9
        .IMEMode = fmIMEModeHiragana           ' 日本語入力
        .Style = fmStyleDropDownList           ' リストから選択のみ
        .TabIndex = baseTab + 1
        .AddItem "ノアFrLH": .AddItem "ノアFrRH"
        .AddItem "ノアRrLH": .AddItem "ノアRrRH"
        .AddItem "アルFrLH": .AddItem "アルFrRH"
        .AddItem "アルRrLH": .AddItem "アルRrRH"
        .Value = rs("品番").Value & ""
    End With

    ' ロット
    Dim tbLot As MSForms.TextBox
    Dim hLot As CTextBoxEvent
    Set tbLot = Me.Controls.Add("Forms.TextBox.1", "TextBoxLot_" & rowNum, True)
    With tbLot
        .Left = COL_LOT: .Top = topPos: .Width = 36: .Height = 20
        .Font.Name = "Meiryo UI"
        .Font.Size = 9
        .Tag = "Lot"
        .IMEMode = fmIMEModeDisable            ' 半角英数
        .MaxLength = 4                         ' ロット番号は4桁
        .TabIndex = baseTab + 2
        .Value = rs("ロット").Value & ""
    End With
    Set hLot = New CTextBoxEvent: Set hLot.TB = tbLot
    Set lotHandlers(rowNum) = hLot
    SetAlphaIME tbLot

    ' 発見
    Dim cbFind As MSForms.ComboBox
    Set cbFind = Me.Controls.Add("Forms.ComboBox.1", "ComboBoxFind_" & rowNum, True)
    With cbFind
        .Left = COL_FIND: .Top = topPos: .Width = 36: .Height = 20
        .Font.Name = "Meiryo UI"
        .Font.Size = 9
        .IMEMode = fmIMEModeDisable            ' 半角英数
        .Style = fmStyleDropDownList           ' リストから選択のみ
        .TabIndex = baseTab + 3
        .AddItem "S": .AddItem "T": .AddItem "M": .AddItem "K"
        .Value = rs("発見").Value & ""
    End With

    ' ゾーン
    Dim tbZone As MSForms.TextBox
    Dim hZone As CTextBoxEvent
    Set tbZone = Me.Controls.Add("Forms.TextBox.1", "TextBoxZone_" & rowNum, True)
    With tbZone
        .Left = COL_ZONE: .Top = topPos: .Width = 36: .Height = 20
        .Font.Name = "Meiryo UI"
        .Font.Size = 9
        .Tag = "Zone"
        .IMEMode = fmIMEModeDisable            ' 半角英数
        .MaxLength = 1                         ' A-E の1文字
        .TabIndex = baseTab + 4
        .Value = rs("ゾーン").Value & ""
    End With
    Set hZone = New CTextBoxEvent: Set hZone.TB = tbZone
    Set zoneHandlers(rowNum) = hZone
    SetAlphaIME tbZone

    ' 番号
    Dim tbNum As MSForms.TextBox
    Dim hNum As CTextBoxEvent
    Set tbNum = Me.Controls.Add("Forms.TextBox.1", "TextBoxNum_" & rowNum, True)
    With tbNum
        .Left = COL_NUM: .Top = topPos: .Width = 36: .Height = 20
        .Font.Name = "Meiryo UI"
        .Font.Size = 9
        .Tag = "Number"
        .IMEMode = fmIMEModeDisable            ' 半角英数
        .MaxLength = 5                         ' 番号は5桁まで
        .TabIndex = baseTab + 5
        .Value = rs("番号").Value & ""
    End With
    Set hNum = New CTextBoxEvent: Set hNum.TB = tbNum
    Set numberHandlers(rowNum) = hNum
    SetAlphaIME tbNum

    ' 数量
    Dim tbQty As MSForms.TextBox
    Dim hQty As CTextBoxEvent
    Set tbQty = Me.Controls.Add("Forms.TextBox.1", "TextBoxQty_" & rowNum, True)
    With tbQty
        .Left = COL_QTY: .Top = topPos: .Width = 36: .Height = 20
        .Font.Name = "Meiryo UI"
        .Font.Size = 9
        .Tag = "Quantity"
        .IMEMode = fmIMEModeDisable            ' 半角英数
        .MaxLength = 4                         ' 数量は4桁まで
        .TabIndex = baseTab + 6
        .Value = rs("数量").Value & ""
    End With
    Set hQty = New CTextBoxEvent: Set hQty.TB = tbQty
    Set quantityHandlers(rowNum) = hQty
    SetAlphaIME tbQty

    ' 差戻し
    Dim tbRet As MSForms.TextBox
    Dim hRet As CTextBoxEvent
    Set tbRet = Me.Controls.Add("Forms.TextBox.1", "TextBoxRet_" & rowNum, True)
    With tbRet
        .Left = COL_RET: .Top = topPos: .Width = 30: .Height = 20
        .Font.Name = "Meiryo UI"
        .Font.Size = 9
        .Tag = "Return"
        .IMEMode = fmIMEModeDisable            ' 半角英数
        .MaxLength = 1                         ' 0 or 1
        .TabIndex = baseTab + 7
        .Value = rs("差戻し").Value & ""
    End With
    Set hRet = New CTextBoxEvent: Set hRet.TB = tbRet
    Set returnHandlers(rowNum) = hRet
    SetAlphaIME tbRet
End Sub

' ============================================
' 動的コントロールの無効化切替（削除モード用）
' ============================================
Private Sub LockDataControls(lockState As Boolean)
    Dim i As Integer

    On Error Resume Next
    For i = 1 To currentRecordCount
        ' 編集可能なコントロールを無効化/有効化（Enabledで見た目も変わる）
        Me.Controls("TextBoxDate_" & i).Enabled = Not lockState
        Me.Controls("ComboBoxItem_" & i).Enabled = Not lockState
        Me.Controls("TextBoxLot_" & i).Enabled = Not lockState
        Me.Controls("ComboBoxFind_" & i).Enabled = Not lockState
        Me.Controls("TextBoxZone_" & i).Enabled = Not lockState
        Me.Controls("TextBoxNum_" & i).Enabled = Not lockState
        Me.Controls("TextBoxQty_" & i).Enabled = Not lockState
        Me.Controls("TextBoxRet_" & i).Enabled = Not lockState
    Next i
    On Error GoTo 0
End Sub

' ============================================
' 動的コントロールクリア
' ============================================
Private Sub ClearDynamicControls()
    Dim i As Integer

    ' ヘッダーラベル削除
    If headerCreated Then
        On Error Resume Next
        Me.Controls.Remove "LabelHdr_ID"
        Me.Controls.Remove "LabelHdr_Date"
        Me.Controls.Remove "LabelHdr_Item"
        Me.Controls.Remove "LabelHdr_Lot"
        Me.Controls.Remove "LabelHdr_Find"
        Me.Controls.Remove "LabelHdr_Zone"
        Me.Controls.Remove "LabelHdr_Num"
        Me.Controls.Remove "LabelHdr_Qty"
        Me.Controls.Remove "LabelHdr_Ret"
        On Error GoTo 0
        headerCreated = False
    End If

    ' データ行削除（逆順）
    For i = currentRecordCount To 1 Step -1
        On Error Resume Next
        Me.Controls.Remove "TextBoxID_" & i
        Me.Controls.Remove "TextBoxDate_" & i
        Me.Controls.Remove "ComboBoxItem_" & i
        Me.Controls.Remove "TextBoxLot_" & i
        Me.Controls.Remove "ComboBoxFind_" & i
        Me.Controls.Remove "TextBoxZone_" & i
        Me.Controls.Remove "TextBoxNum_" & i
        Me.Controls.Remove "TextBoxQty_" & i
        Me.Controls.Remove "TextBoxRet_" & i

        ' CTextBoxEventハンドラ解放
        Set dateHandlers(i) = Nothing
        Set lotHandlers(i) = Nothing
        Set zoneHandlers(i) = Nothing
        Set numberHandlers(i) = Nothing
        Set quantityHandlers(i) = Nothing
        Set returnHandlers(i) = Nothing
        On Error GoTo 0
    Next i

    currentRecordCount = 0
End Sub

' ============================================
' フォームサイズ調整
' ============================================
Private Sub ResizeForm()
    Dim newHeight As Integer
    newHeight = DATA_START_TOP + currentRecordCount * ROW_HEIGHT + BUTTON_MARGIN + 60

    Me.Height = newHeight

    ' ボタン位置調整
    Dim buttonTop As Integer
    buttonTop = DATA_START_TOP + currentRecordCount * ROW_HEIGHT + BUTTON_MARGIN

    CommandButtonSearch.Top = buttonTop
    CommandButtonExecute.Top = buttonTop
    CommandButtonClose.Top = buttonTop

    ' ボタンのタブ順更新（動的コントロールの後）
    Dim lastTab As Integer
    lastTab = 4 + currentRecordCount * 8  ' 動的コントロールの次
    CommandButtonSearch.TabIndex = lastTab
    CommandButtonExecute.TabIndex = lastTab + 1
    CommandButtonClose.TabIndex = lastTab + 2
End Sub

' ============================================
' 実行ボタン：修正または削除
' ============================================
Private Sub CommandButtonExecute_Click()
    If currentRecordCount = 0 Then
        MsgBox "先に検索を実行してください。", vbExclamation
        Exit Sub
    End If

    ' モード判定（全件処理）
    If ToggleButtonMode.Value Then
        ' 削除モード
        ExecuteDelete currentRecordCount
    Else
        ' 修正モード
        ExecuteUpdate currentRecordCount
    End If
End Sub

' ============================================
' 修正実行
' ============================================
Private Sub ExecuteUpdate(selectedCount As Integer)
    Dim confirm As VbMsgBoxResult
    confirm = MsgBox(selectedCount & "件のデータを修正します。よろしいですか？", _
                     vbQuestion + vbYesNo, "確認")

    If confirm <> vbYes Then Exit Sub

    Dim conn As Object
    Dim sql As String
    Dim yearValue As Integer
    Dim dbPath As String
    Dim i As Integer
    Dim successCount As Integer
    Dim recordID As Long

    yearValue = CInt(ComboBoxYear.Value)
    dbPath = DB_BASE_PATH & yearValue & "年\" & DB_FILE_PREFIX & yearValue & ".accdb"

    On Error GoTo ErrorHandler

    Application.StatusBar = "データを修正中..."

    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"

    conn.BeginTrans

    successCount = 0
    For i = 1 To currentRecordCount
        recordID = CLng(Me.Controls("TextBoxID_" & i).Value)

        sql = "UPDATE [_不良集計ゾーン別] SET " & _
              "[日付] = " & FormatDateForSQL(Me.Controls("TextBoxDate_" & i).Value) & ", " & _
              "[品番] = '" & EscapeSQL(Me.Controls("ComboBoxItem_" & i).Value) & "', " & _
              "[ロット] = " & FormatNumForSQL(Me.Controls("TextBoxLot_" & i).Value) & ", " & _
              "[発見] = '" & EscapeSQL(Me.Controls("ComboBoxFind_" & i).Value) & "', " & _
              "[ゾーン] = '" & EscapeSQL(Me.Controls("TextBoxZone_" & i).Value) & "', " & _
              "[番号] = '" & EscapeSQL(Me.Controls("TextBoxNum_" & i).Value) & "', " & _
              "[数量] = " & FormatNumForSQL(Me.Controls("TextBoxQty_" & i).Value) & ", " & _
              "[差戻し] = " & FormatNumForSQL(Me.Controls("TextBoxRet_" & i).Value) & " " & _
              "WHERE [ID] = " & recordID

        conn.Execute sql
        successCount = successCount + 1
    Next i

    conn.CommitTrans
    conn.Close

    Application.StatusBar = False
    MsgBox successCount & "件のデータを修正しました。", vbInformation

    ' データ変更フラグをセット
    DataModified = True

    ' フォームをリセット
    ClearDynamicControls
    TextBoxIDs.Value = ""
    CommandButtonExecute.Enabled = False
    Me.Height = 180

    ' 上部コントロールを再有効化
    ToggleButtonMode.Enabled = True
    TextBoxIDs.Enabled = True
    CommandButtonSearch.Enabled = True

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    If Not conn Is Nothing Then
        If conn.State = 1 Then
            conn.RollbackTrans
            conn.Close
        End If
    End If
    MsgBox "修正中にエラーが発生しました: " & Err.Description, vbCritical

    ' エラー時も再有効化
    ToggleButtonMode.Enabled = True
    TextBoxIDs.Enabled = True
    CommandButtonSearch.Enabled = True
End Sub

' ============================================
' 削除実行
' ============================================
Private Sub ExecuteDelete(selectedCount As Integer)
    Dim confirm As VbMsgBoxResult
    confirm = MsgBox(selectedCount & "件のデータを削除します。" & vbCrLf & _
                     "この操作は取り消せません。よろしいですか？", _
                     vbExclamation + vbYesNo, "警告")

    If confirm <> vbYes Then Exit Sub

    Dim conn As Object
    Dim sql As String
    Dim yearValue As Integer
    Dim dbPath As String
    Dim i As Integer
    Dim successCount As Integer
    Dim recordID As Long

    yearValue = CInt(ComboBoxYear.Value)
    dbPath = DB_BASE_PATH & yearValue & "年\" & DB_FILE_PREFIX & yearValue & ".accdb"

    On Error GoTo ErrorHandler

    Application.StatusBar = "データを削除中..."

    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"

    conn.BeginTrans

    successCount = 0
    For i = 1 To currentRecordCount
        recordID = CLng(Me.Controls("TextBoxID_" & i).Value)

        sql = "DELETE FROM [_不良集計ゾーン別] WHERE [ID] = " & recordID

        conn.Execute sql
        successCount = successCount + 1
    Next i

    conn.CommitTrans
    conn.Close

    Application.StatusBar = False
    MsgBox successCount & "件のデータを削除しました。", vbInformation

    ' データ変更フラグをセット
    DataModified = True

    ' 削除後はフォームを閉じる
    Unload Me

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    If Not conn Is Nothing Then
        If conn.State = 1 Then
            conn.RollbackTrans
            conn.Close
        End If
    End If
    MsgBox "削除中にエラーが発生しました: " & Err.Description, vbCritical

    ' エラー時も再有効化
    ToggleButtonMode.Enabled = True
    TextBoxIDs.Enabled = True
    CommandButtonSearch.Enabled = True
End Sub

' ============================================
' 閉じるボタン
' ============================================
Private Sub CommandButtonClose_Click()
    Unload Me
End Sub

' ============================================
' 補助関数：ID入力を解析（カンマ区切り＋範囲指定対応）
' 入力例: "22,23,24" または "22-25" または "22-25,30,32-34"
' 戻り値: ID配列（Long型）、エラー時は空配列
' ============================================
Private Function ParseIDInput(inputText As String, ByRef errMsg As String) As Variant
    Dim parts() As String
    Dim result() As Long
    Dim resultCount As Long
    Dim i As Long, j As Long
    Dim part As String
    Dim rangeStart As Long, rangeEnd As Long
    Dim hyphenPos As Long

    errMsg = ""
    resultCount = 0
    ReDim result(0 To MAX_RECORDS - 1)  ' 最大件数分を確保

    ' 前処理：スペース除去
    inputText = Replace(inputText, " ", "")
    inputText = Replace(inputText, "　", "")

    If Len(inputText) = 0 Then
        errMsg = "IDを入力してください。"
        ParseIDInput = Array()
        Exit Function
    End If

    ' カンマで分割
    parts = Split(inputText, ",")

    For i = 0 To UBound(parts)
        part = Trim(parts(i))
        If Len(part) = 0 Then GoTo NextPart

        hyphenPos = InStr(part, "-")

        If hyphenPos > 0 Then
            ' 範囲指定（例: 22-25）
            Dim leftPart As String, rightPart As String
            leftPart = Left(part, hyphenPos - 1)
            rightPart = Mid(part, hyphenPos + 1)

            ' 数値チェック
            If Not IsNumeric(leftPart) Or Not IsNumeric(rightPart) Then
                errMsg = "範囲指定が不正です: " & part
                ParseIDInput = Array()
                Exit Function
            End If

            rangeStart = CLng(leftPart)
            rangeEnd = CLng(rightPart)

            ' 範囲の妥当性チェック
            If rangeStart > rangeEnd Then
                errMsg = "範囲指定の開始値が終了値より大きいです: " & part
                ParseIDInput = Array()
                Exit Function
            End If

            ' 範囲を展開
            For j = rangeStart To rangeEnd
                If resultCount >= MAX_RECORDS Then
                    errMsg = "展開後のID数が最大" & MAX_RECORDS & "件を超えます。"
                    ParseIDInput = Array()
                    Exit Function
                End If
                result(resultCount) = j
                resultCount = resultCount + 1
            Next j
        Else
            ' 単一ID
            If Not IsNumeric(part) Then
                errMsg = "IDは数値で入力してください: " & part
                ParseIDInput = Array()
                Exit Function
            End If

            If resultCount >= MAX_RECORDS Then
                errMsg = "ID数が最大" & MAX_RECORDS & "件を超えます。"
                ParseIDInput = Array()
                Exit Function
            End If
            result(resultCount) = CLng(part)
            resultCount = resultCount + 1
        End If
NextPart:
    Next i

    ' 結果がない場合
    If resultCount = 0 Then
        errMsg = "有効なIDがありません。"
        ParseIDInput = Array()
        Exit Function
    End If

    ' 実際のサイズにリサイズ
    ReDim Preserve result(0 To resultCount - 1)
    ParseIDInput = result
End Function

' ============================================
' 補助関数：日付をSQL用にフォーマット
' ============================================
Private Function FormatDateForSQL(dateValue As Variant) As String
    If IsDate(dateValue) Then
        FormatDateForSQL = "#" & Format(CDate(dateValue), "yyyy/mm/dd") & "#"
    Else
        FormatDateForSQL = "NULL"
    End If
End Function

' ============================================
' 補助関数：数値をSQL用にフォーマット
' ============================================
Private Function FormatNumForSQL(numValue As Variant) As String
    If IsNumeric(numValue) Then
        FormatNumForSQL = CStr(numValue)
    Else
        FormatNumForSQL = "NULL"
    End If
End Function

' ============================================
' 補助関数：SQLエスケープ
' ============================================
Private Function EscapeSQL(text As Variant) As String
    If IsNull(text) Or text = "" Then
        EscapeSQL = ""
    Else
        EscapeSQL = Replace(CStr(text), "'", "''")
    End If
End Function
