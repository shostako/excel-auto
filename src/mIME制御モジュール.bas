Attribute VB_Name = "mIME制御モジュール"
' ========================================
' モジュール名: IME制御モジュール
' 処理概要: UserForm上のコントロールのIME入力モードを制御
' 対象コントロール: ComboBox（日本語入力）、TextBox（半角英数入力）
' 依存API: Windows SendMessage API（64bit/32bit対応）
' 使用場面: UserForm_KeyDownイベントから呼び出し
' ========================================

Option Explicit

' ============================================
' Windows API宣言：SendMessage（64bit/32bit対応）
' ============================================
#If VBA7 Then
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, _
        lParam As Any) As LongPtr
#Else
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
        lParam As Any) As Long
#End If

' ============================================
' IME制御用定数定義
' ============================================
Private Const WM_IME_CONTROL As Long = &H283
Private Const IMC_GETOPENSTATUS As Long = &H5
Private Const IMC_SETOPENSTATUS As Long = &H6

' ============================================
' IME制御プロシージャ群
' ============================================

Public Sub SetJapaneseIME(ctrl As MSForms.Control)
    On Error Resume Next
    ctrl.IMEMode = fmIMEModeHiragana
    On Error GoTo 0
End Sub

Public Sub SetAlphaIME(ctrl As MSForms.Control)
    On Error Resume Next
    ctrl.IMEMode = fmIMEModeDisable
    On Error GoTo 0
End Sub

Public Sub HandleFormKeyDown(frm As Object)
    On Error Resume Next

    If Not frm.ActiveControl Is Nothing Then
        If TypeName(frm.ActiveControl) = "ComboBox" Then
            If frm.ActiveControl.IMEMode = fmIMEModeHiragana Then
                SetJapaneseIME frm.ActiveControl
            End If
        ElseIf TypeName(frm.ActiveControl) = "TextBox" Then
            SetAlphaIME frm.ActiveControl
        End If
    End If

    On Error GoTo 0
End Sub
