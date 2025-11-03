VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTextBoxEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' クラスモジュール名: CTextBoxEvent
Option Explicit

' 操作対象の TextBox を WithEvents で保持
Public WithEvents TB As MSForms.TextBox

'==============================================
' Change イベント：タグに応じた変換処理
Private Sub TB_Change()
    Dim lastChar As String
    Dim cursorPos As Long

    If Len(TB.Value) = 0 Then Exit Sub

    lastChar = Right(TB.Value, 1)
    cursorPos = TB.SelStart

    Select Case TB.Tag
    Case "Zone"
        ' ゾーンは a-e を A-E に
        If lastChar >= "a" And lastChar <= "e" Then
            TB.Value = Left(TB.Value, Len(TB.Value) - 1) & UCase(lastChar)
            TB.SelStart = cursorPos
        End If

    Case "Number"
        ' 番号は a-z を A-Z に
        If lastChar >= "a" And lastChar <= "z" Then
            TB.Value = Left(TB.Value, Len(TB.Value) - 1) & UCase(lastChar)
            TB.SelStart = cursorPos
        End If

    Case "Date"
        ' 日付は特に何もしない (KeyPress/Exit で制御)

    Case "Month"
        ' 月は特に何もしない (KeyPress/Exit で制御)

    Case "Lot"
        ' ロットは特に何もしない (KeyPress で制御)

    Case "Quantity"
        ' 数量は Change 時には特に何もしない (KeyPress で制御)

    Case "Return"
        ' 差戻しは特に変換処理なし（数値0/1のみ許可）
    End Select
End Sub

'==============================================
' KeyPress イベント：入力時の制限
Private Sub TB_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case TB.Tag
        Case "Zone"
            ' A-E（大文字小文字）のみ許可
            Select Case KeyAscii
                Case 65 To 69   ' A-E を許可
                Case 97 To 101  ' a-e を許可
                Case 8          ' Backspace を許可
                Case Else
                    KeyAscii = 0  ' それ以外は無効化
            End Select

        Case "Quantity", "Lot", "Month"
            ' 数値のみ入力許可
            Select Case KeyAscii
                Case 48 To 57   ' 0-9 を許可
                Case 8          ' Backspace を許可
                Case Else
                    KeyAscii = 0  ' それ以外は無効化
            End Select

        Case "Date"
            ' 日付は数字とスラッシュのみ許可
            Select Case KeyAscii
                Case 48 To 57   ' 0-9 を許可
                Case 47         ' / を許可
                Case 8          ' Backspace を許可
                Case Else
                    KeyAscii = 0  ' それ以外は無効化
            End Select

        Case "Return"
            ' スペースキーで0→1→0循環切り替え
            If KeyAscii = 32 Then  ' Space
                Select Case TB.Value
                    Case "", "0"
                        TB.Value = "1"  ' 0→1 または 空→1
                    Case "1"
                        TB.Value = "0"  ' 1→0
                    Case Else
                        TB.Value = "0"  ' その他→0（安全のため）
                End Select
                KeyAscii = 0  ' スペース入力をキャンセル
            ' 数値の0または1のみ許可
            ElseIf KeyAscii = 48 Or KeyAscii = 49 Then  ' 0 or 1
                ' 0または1を許可（何もしない）
            ElseIf KeyAscii = 8 Then  ' Backspace
                ' Backspaceを許可（何もしない）
            Else
                KeyAscii = 0  ' それ以外は無効化
            End If
    End Select
End Sub

'==============================================
' Exit イベント：入力チェックと正規化
Private Sub TB_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim v As String
    v = StrConv(TB.Value, vbNarrow)  ' 全角→半角変換

    Select Case TB.Tag
    Case "Zone"
        ' 大文字化＆ A-E のみ許可
        v = UCase(v)
        If v = "" Then
            TB.Value = ""
        ElseIf v Like "[A-E]" Then
            TB.Value = v
        Else
            MsgBox "ゾーンは A〜E のいずれかを入力してください。", vbExclamation
            TB.Value = ""
            TB.SetFocus
            Cancel = True
        End If

    Case "Number"
        ' 数字またはアルファベットのみ
        If v = "" Then
            TB.Value = ""
        ElseIf IsNumeric(v) Or v Like "[A-Za-z]*" Then
            TB.Value = v
        Else
            MsgBox "番号は数字またはアルファベットのみ入力してください。", vbExclamation
            TB.Value = ""
            TB.SetFocus
            Cancel = True
        End If

    Case "Date"
        ' 日付形式チェック
        If v = "" Then
            TB.Value = ""
        ElseIf IsDate(v) Then
            TB.Value = Format(CDate(v), "m/d")
        Else
            MsgBox "日付は「yyyy/m/d」または「m/d」形式で入力してください。", vbExclamation
            TB.Value = ""
            TB.SetFocus
            Cancel = True
        End If

    Case "Month"
        ' 0-12の整数のみ
        If v = "" Then
            TB.Value = ""
        ElseIf IsNumeric(v) And Val(v) = Int(Val(v)) And Val(v) >= 0 And Val(v) <= 12 Then
            TB.Value = v
        Else
            MsgBox "注番月は0〜12の整数を入力してください。", vbExclamation
            TB.Value = ""
            TB.SetFocus
            Cancel = True
        End If

    Case "Lot"
        ' 正の整数のみ
        If v = "" Then
            TB.Value = ""
        ElseIf IsNumeric(v) And Val(v) = Int(Val(v)) And Val(v) > 0 Then
            TB.Value = v
        Else
            MsgBox "ロットは正の整数を入力してください。", vbExclamation
            TB.Value = ""
            TB.SetFocus
            Cancel = True
        End If

    Case "Quantity"
        ' 正の整数のみ
        If v = "" Then
            TB.Value = ""
        ElseIf IsNumeric(v) And Val(v) = Int(Val(v)) And Val(v) > 0 Then
            TB.Value = v
        Else
            MsgBox "数量は正の整数を入力してください。", vbExclamation
            TB.Value = ""
            TB.SetFocus
            Cancel = True
        End If

    Case "Return"
        ' 差戻しは「0」「1」または空白のみ許可（空白は0として扱う）
        If v = "" Then
            TB.Value = "0"  ' 空白を0に置き換え
        ElseIf v = "0" Or v = "1" Then
            TB.Value = v
        Else
            MsgBox "差戻しは「0」または「1」を入力してください。", vbExclamation
            TB.Value = "0"
            TB.SetFocus
            Cancel = True
        End If
    End Select
End Sub
