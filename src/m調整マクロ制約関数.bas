Attribute VB_Name = "m調整マクロ制約関数"
Option Explicit

' ==========================================
' 調整マクロ用制約関数モジュール
' ==========================================
' m調整_自動均し、m調整_完全自動均し、m調整_グループ日程移動 で使用
'
' 【含まれる関数】
' 1. 移動先制約チェック: 品番移動時の制約チェック
' 2. グループ移動制約チェック: グループ移動時の制約チェック
' 3. 指定日に号口単品あり: 指定日の号口単品存在チェック
' 4. 指定日に補給品あり: 指定日の補給品存在チェック
' ==========================================

' ==========================================
' 1. 移動先制約チェック（単品番用）
' ==========================================
Public Function 移動先制約チェック( _
    ByVal 品番 As String, _
    ByVal 移動先日 As Long, _
    ByRef arr均し As Variant, _
    ByRef 品番マスタ As Object, _
    ByVal 成形品番列 As Long, _
    ByVal 開始列 As Long _
) As Boolean

    ' 号口単品チェック
    Dim 号補 As String, セット As String
    号補 = CStr(品番マスタ(品番)("号/補"))
    セット = CStr(品番マスタ(品番)("セット"))

    If 号補 = "号口" And セット <> "SET" Then
        ' 号口単品の場合：移動先日に他の号口単品または補給品があればNG
        If 指定日に号口単品あり(移動先日, arr均し, 品番マスタ, 成形品番列, 開始列, 品番) Then
            Debug.Print "制約違反: 号口単品[" & 品番 & "]は" & 移動先日 & "日に他の号口単品あり"
            移動先制約チェック = False
            Exit Function
        End If

        If 指定日に補給品あり(移動先日, arr均し, 品番マスタ, 成形品番列, 開始列) Then
            Debug.Print "制約違反: 号口単品[" & 品番 & "]は" & 移動先日 & "日に補給品あり"
            移動先制約チェック = False
            Exit Function
        End If
    End If

    ' 補給品チェック
    If 号補 = "補給品" Then
        ' 補給品の場合：移動先日に号口単品があればNG
        If 指定日に号口単品あり(移動先日, arr均し, 品番マスタ, 成形品番列, 開始列, "") Then
            Debug.Print "制約違反: 補給品[" & 品番 & "]は" & 移動先日 & "日に号口単品あり"
            移動先制約チェック = False
            Exit Function
        End If
    End If

    ' 制約OK
    移動先制約チェック = True
End Function

' ==========================================
' 2. グループ移動制約チェック
' ==========================================
Public Function グループ移動制約チェック( _
    ByRef グループ品番リスト As Object, _
    ByVal 移動先日 As Long, _
    ByRef arr均し As Variant, _
    ByRef 品番マスタ As Object, _
    ByVal 成形品番列 As Long, _
    ByVal 開始列 As Long _
) As String

    Dim 制約違反リスト As String
    制約違反リスト = ""

    Dim key As Variant, 品番 As String
    Dim 号補 As String, セット As String

    ' グループ内の各品番について制約チェック
    For Each key In グループ品番リスト.Keys
        品番 = CStr(key)

        If Not 品番マスタ.Exists(品番) Then GoTo NextItem

        号補 = CStr(品番マスタ(品番)("号/補"))
        セット = CStr(品番マスタ(品番)("セット"))

        ' 号口単品の場合
        If 号補 = "号口" And セット <> "SET" Then
            ' 移動先日に他の号口単品があるかチェック
            If 指定日に号口単品あり(移動先日, arr均し, 品番マスタ, 成形品番列, 開始列, 品番) Then
                制約違反リスト = 制約違反リスト & "・号口単品[" & 品番 & "]: 移動先日に他の号口単品が存在" & vbCrLf
            End If

            ' 移動先日に補給品があるかチェック
            If 指定日に補給品あり(移動先日, arr均し, 品番マスタ, 成形品番列, 開始列) Then
                制約違反リスト = 制約違反リスト & "・号口単品[" & 品番 & "]: 移動先日に補給品が存在" & vbCrLf
            End If
        End If

        ' 補給品の場合
        If 号補 = "補給品" Then
            ' 移動先日に号口単品があるかチェック
            If 指定日に号口単品あり(移動先日, arr均し, 品番マスタ, 成形品番列, 開始列, "") Then
                制約違反リスト = 制約違反リスト & "・補給品[" & 品番 & "]: 移動先日に号口単品が存在" & vbCrLf
            End If
        End If
NextItem:
    Next key

    グループ移動制約チェック = 制約違反リスト
End Function

' ==========================================
' 3. 指定日に号口単品あり（arr均しベース）
' ==========================================
Private Function 指定日に号口単品あり( _
    ByVal 指定日 As Long, _
    ByRef arr均し As Variant, _
    ByRef 品番マスタ As Object, _
    ByVal 成形品番列 As Long, _
    ByVal 開始列 As Long, _
    Optional ByVal 除外品番 As String = "" _
) As Boolean

    Dim r As Long, 品番 As String
    Dim 号補 As String, セット As String
    Dim 数量 As Long

    For r = 1 To UBound(arr均し, 1)
        品番 = CStr(arr均し(r, 成形品番列))

        ' 除外品番はスキップ
        If 品番 = 除外品番 Then GoTo NextRow

        ' 品番マスタ存在チェック
        If Not 品番マスタ.Exists(品番) Then GoTo NextRow

        ' 号口単品かチェック
        号補 = CStr(品番マスタ(品番)("号/補"))
        セット = CStr(品番マスタ(品番)("セット"))

        If 号補 = "号口" And セット <> "SET" Then
            ' 指定日に数量があるかチェック
            数量 = 0
            On Error Resume Next
            数量 = CLng(arr均し(r, 開始列 + 指定日 - 1))
            On Error GoTo 0

            If 数量 > 0 Then
                指定日に号口単品あり = True
                Exit Function
            End If
        End If
NextRow:
    Next r

    指定日に号口単品あり = False
End Function

' ==========================================
' 4. 指定日に補給品あり（arr均しベース）
' ==========================================
Private Function 指定日に補給品あり( _
    ByVal 指定日 As Long, _
    ByRef arr均し As Variant, _
    ByRef 品番マスタ As Object, _
    ByVal 成形品番列 As Long, _
    ByVal 開始列 As Long _
) As Boolean

    Dim r As Long, 品番 As String
    Dim 号補 As String
    Dim 数量 As Long

    For r = 1 To UBound(arr均し, 1)
        品番 = CStr(arr均し(r, 成形品番列))

        ' 品番マスタ存在チェック
        If Not 品番マスタ.Exists(品番) Then GoTo NextRow

        ' 補給品かチェック
        号補 = CStr(品番マスタ(品番)("号/補"))

        If 号補 = "補給品" Then
            ' 指定日に数量があるかチェック
            数量 = 0
            On Error Resume Next
            数量 = CLng(arr均し(r, 開始列 + 指定日 - 1))
            On Error GoTo 0

            If 数量 > 0 Then
                指定日に補給品あり = True
                Exit Function
            End If
        End If
NextRow:
    Next r

    指定日に補給品あり = False
End Function

