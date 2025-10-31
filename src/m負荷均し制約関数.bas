Attribute VB_Name = "m負荷均し制約関数"
Option Explicit

' ==========================================
' 負荷均しマクロ用制約関数モジュール
' ==========================================
' m転記_負荷均し に追加する制約チェック関数群
'
' 【含まれる関数】
' 1. 補給品優先配置_モール: モール品補給品を最初の納期に配置
' 2. 補給品優先配置_通常: 非モール補給品をグループごとに配置
' 3. 系列別均し処理: 系列ごとに均し処理を実行
' 4. 当日号口単品配置済み: 号口単品の同日配置チェック
' 5. 当日補給品あり: 当日の補給品存在チェック
' 6. 当日号口単品あり: 当日の号口単品存在チェック
' 7. 品番の最初の納期を取得: 展開シートから最初の納期を取得
' ==========================================

' ==========================================
' 1. モール品補給品の優先配置
' ==========================================
Public Sub 補給品優先配置_モール( _
    ByRef 品番マスタ As Object, _
    ByRef 月間残数 As Object, _
    ByRef 転記データ As Object, _
    ByRef 当日割り当て As Object, _
    ByRef 当日型番割り当て As Object, _
    ByRef arr展開 As Variant, _
    ByVal 成形品番列 As Long, _
    ByVal 開始列 As Long, _
    ByVal maxDay As Long, _
    ByVal 日次目標 As Double, _
    ByVal 誤差許容率 As Double, _
    ByRef workDays As Object _
)
    Debug.Print "--- モール品補給品配置開始 ---"

    Dim key As Variant, 品番 As String
    Dim 仕様 As String, 号補 As String
    Dim 最初の納期 As Long, 残数 As Long
    Dim 配置数量 As Long, 単位 As Long, 上限 As Long
    Dim 転記キー As String

    For Each key In 品番マスタ.Keys
        品番 = CStr(key)
        仕様 = CStr(品番マスタ(品番)("仕様"))
        号補 = CStr(品番マスタ(品番)("号/補"))

        ' モール品かつ補給品のみ処理
        If 号補 = "補給品" And InStr(仕様, "モール") > 0 Then
            残数 = CLng(月間残数(品番))
            If 残数 = 0 Then GoTo NextItem

            ' 最初の納期を取得
            最初の納期 = 品番の最初の納期を取得(品番, arr展開, 成形品番列, 開始列, maxDay)

            If 最初の納期 = 0 Then
                Debug.Print "警告: 品番[" & 品番 & "]の納期が見つかりません"
                GoTo NextItem
            End If

            ' 配置数量算出（残数全て、ただし単位・上限制約あり）
            単位 = CLng(品番マスタ(品番)("単位"))
            上限 = CLng(品番マスタ(品番)("上限"))

            配置数量 = 残数
            If 配置数量 > 上限 Then 配置数量 = 上限
            配置数量 = Int(配置数量 / 単位) * 単位

            If 配置数量 > 0 Then
                ' 転記データに記録
                転記キー = 品番 & "_" & 最初の納期
                If 転記データ.Exists(転記キー) Then
                    転記データ(転記キー) = CLng(転記データ(転記キー)) + 配置数量
                Else
                    転記データ(転記キー) = 配置数量
                End If

                ' 残数更新
                月間残数(品番) = 残数 - 配置数量

                ' 当日割り当て累積
                If 当日割り当て.Exists(最初の納期) Then
                    当日割り当て(最初の納期) = CLng(当日割り当て(最初の納期)) + 配置数量
                Else
                    当日割り当て(最初の納期) = 配置数量
                End If

                Debug.Print "モール補給品[" & 品番 & "]: " & 最初の納期 & "日に" & 配置数量 & "個配置 (残数: " & 月間残数(品番) & ")"
            End If
        End If
NextItem:
    Next key

    Debug.Print "--- モール品補給品配置完了 ---"
End Sub

' ==========================================
' 2. 非モール補給品の優先配置（グループごと）
' ==========================================
Public Sub 補給品優先配置_通常( _
    ByRef 品番マスタ As Object, _
    ByRef 月間残数 As Object, _
    ByRef 転記データ As Object, _
    ByRef 当日割り当て As Object, _
    ByRef 当日型番割り当て As Object, _
    ByRef グループ初回日 As Object, _
    ByRef workDays As Object, _
    ByVal 日次目標 As Double, _
    ByVal 誤差許容率 As Double _
)
    Debug.Print "--- 非モール補給品配置開始 ---"

    ' グループ別に補給品をまとめる
    Dim グループ別補給品 As Object
    Set グループ別補給品 = CreateObject("Scripting.Dictionary")

    Dim key As Variant, 品番 As String
    Dim 仕様 As String, 号補 As String, グループID As String

    For Each key In 品番マスタ.Keys
        品番 = CStr(key)
        仕様 = CStr(品番マスタ(品番)("仕様"))
        号補 = CStr(品番マスタ(品番)("号/補"))

        ' 非モール補給品のみ処理
        If 号補 = "補給品" And InStr(仕様, "モール") = 0 Then
            グループID = CStr(品番マスタ(品番)("グループ"))

            If Not グループ別補給品.Exists(グループID) Then
                Set グループ別補給品(グループID) = CreateObject("Scripting.Dictionary")
            End If
            グループ別補給品(グループID)(品番) = True
        End If
    Next key

    Debug.Print "補給品グループ数: " & グループ別補給品.Count

    ' グループごとに稼働日の早い順に配置
    Dim グループKey As Variant
    Dim wdIdx As Long, 稼働日 As Long
    Dim 配置成功 As Boolean

    For Each グループKey In グループ別補給品.Keys
        グループID = CStr(グループKey)
        配置成功 = False

        Debug.Print "グループ[" & グループID & "]の補給品配置開始"

        For wdIdx = 1 To workDays.Count
            稼働日 = CLng(workDays(wdIdx))

            ' このグループの全品番を配置可能かチェック
            Dim グループ品番リスト As Object
            Set グループ品番リスト = グループ別補給品(グループID)

            Dim 配置可能 As Boolean
            配置可能 = True

            ' 簡易チェック：日次目標の150%まで許容
            Dim 当日既割り当て As Long
            当日既割り当て = 0
            If 当日割り当て.Exists(稼働日) Then
                当日既割り当て = CLng(当日割り当て(稼働日))
            End If

            Dim グループ合計数量 As Long
            グループ合計数量 = 0

            Dim 品番Key As Variant
            For Each 品番Key In グループ品番リスト.Keys
                品番 = CStr(品番Key)
                グループ合計数量 = グループ合計数量 + CLng(月間残数(品番))
            Next 品番Key

            ' 150%制限チェック
            If 当日既割り当て + グループ合計数量 > 日次目標 * 1.5 Then
                配置可能 = False
            End If

            If 配置可能 Then
                ' 配置実行
                For Each 品番Key In グループ品番リスト.Keys
                    品番 = CStr(品番Key)
                    Dim 残数 As Long, 配置数量 As Long
                    Dim 単位 As Long, 上限 As Long
                    Dim 転記キー As String

                    残数 = CLng(月間残数(品番))
                    If 残数 = 0 Then GoTo NextGroupItem

                    単位 = CLng(品番マスタ(品番)("単位"))
                    上限 = CLng(品番マスタ(品番)("上限"))

                    配置数量 = 残数
                    If 配置数量 > 上限 Then 配置数量 = 上限
                    配置数量 = Int(配置数量 / 単位) * 単位

                    If 配置数量 > 0 Then
                        転記キー = 品番 & "_" & 稼働日
                        If 転記データ.Exists(転記キー) Then
                            転記データ(転記キー) = CLng(転記データ(転記キー)) + 配置数量
                        Else
                            転記データ(転記キー) = 配置数量
                        End If

                        月間残数(品番) = 残数 - 配置数量

                        If 当日割り当て.Exists(稼働日) Then
                            当日割り当て(稼働日) = CLng(当日割り当て(稼働日)) + 配置数量
                        Else
                            当日割り当て(稼働日) = 配置数量
                        End If

                        Debug.Print "  品番[" & 品番 & "]: " & 稼働日 & "日に" & 配置数量 & "個配置"
                    End If
NextGroupItem:
                Next 品番Key

                ' グループ初回日を記録
                If グループID <> "" And Not グループ初回日.Exists(グループID) Then
                    グループ初回日(グループID) = 稼働日
                End If

                配置成功 = True
                Exit For
            End If
        Next wdIdx

        If Not 配置成功 Then
            Debug.Print "警告: グループ[" & グループID & "]の補給品を配置できませんでした"
        End If
    Next グループKey

    Debug.Print "--- 非モール補給品配置完了 ---"
End Sub

' ==========================================
' 3. 系列別均し処理
' ==========================================
Public Sub 系列別均し処理( _
    ByVal モール条件 As String, _
    ByVal 系列 As String, _
    ByVal 優先度 As Long, _
    ByRef 品番マスタ As Object, _
    ByRef セットペアマスタ As Object, _
    ByRef 月間残数 As Object, _
    ByRef 転記データ As Object, _
    ByRef 当日割り当て As Object, _
    ByRef 当日型番割り当て As Object, _
    ByRef グループ初回日 As Object, _
    ByRef workDays As Object, _
    ByVal 日次目標 As Double, _
    ByVal 誤差許容率 As Double, _
    ByVal maxDay As Long, _
    ByVal 開始列 As Long, _
    ByRef arr均し As Variant _
)
    Debug.Print "系列別均し: モール=" & モール条件 & ", 系列=" & 系列 & ", 優先度=" & 優先度

    Dim key As Variant, 品番 As String
    Dim 仕様 As String, 系列値 As String
    Dim wdIdx As Long, 稼働日 As Long
    Dim グループID As String, 転記キー As String
    Dim 割り当て As Long

    ' ==========================================
    ' 非セット品番の処理
    ' ==========================================
    For Each key In 品番マスタ.Keys
        品番 = CStr(key)

        ' 優先度フィルタ
        If 品番マスタ(品番)("優先度") <> 優先度 Then GoTo NextItem1

        ' セット品はスキップ
        If 品番マスタ(品番)("セット") = "SET" Then GoTo NextItem1

        ' 系列フィルタ
        系列値 = CStr(品番マスタ(品番)("系列"))
        If 系列値 <> 系列 Then GoTo NextItem1

        ' モール条件フィルタ
        仕様 = CStr(品番マスタ(品番)("仕様"))
        If モール条件 = "モール" Then
            If InStr(仕様, "モール") = 0 Then GoTo NextItem1
        Else
            If InStr(仕様, "モール") > 0 Then GoTo NextItem1
        End If

        ' 残数チェック
        If 月間残数(品番) = 0 Then GoTo NextItem1

        グループID = 品番マスタ(品番)("グループ")

        ' 稼働日ループ
        For wdIdx = 1 To workDays.Count
            稼働日 = CLng(workDays(wdIdx))

            Dim 対象稼働日 As Long
            対象稼働日 = 稼働日

            ' グループ制約チェック（初回割り当て日に追従）
            If グループID <> "" And グループ初回日.Exists(グループID) Then
                Dim 初回日 As Long
                初回日 = CLng(グループ初回日(グループID))

                ' 初回日を優先的に試す
                Dim 初回日割り当て As Long
                初回日割り当て = 割り当て可能数を算出_制約付き(品番, 初回日, 品番マスタ, 月間残数, 当日割り当て, 当日型番割り当て, 転記データ, 日次目標, 誤差許容率)

                If 初回日割り当て > 0 Then
                    対象稼働日 = 初回日
                End If
            End If

            ' 割り当て可能数算出
            割り当て = 割り当て可能数を算出_制約付き(品番, 対象稼働日, 品番マスタ, 月間残数, 当日割り当て, 当日型番割り当て, 転記データ, 日次目標, 誤差許容率)

            If 割り当て > 0 Then
                ' グループ初回日記録
                If グループID <> "" And Not グループ初回日.Exists(グループID) Then
                    グループ初回日(グループID) = 対象稼働日
                End If

                ' 転記データ記録
                転記キー = 品番 & "_" & 対象稼働日
                If 転記データ.Exists(転記キー) Then
                    転記データ(転記キー) = CLng(転記データ(転記キー)) + 割り当て
                Else
                    転記データ(転記キー) = 割り当て
                End If

                ' 残数更新
                月間残数(品番) = CLng(月間残数(品番)) - 割り当て

                ' 当日割り当て累積
                If 当日割り当て.Exists(対象稼働日) Then
                    当日割り当て(対象稼働日) = CLng(当日割り当て(対象稼働日)) + 割り当て
                Else
                    当日割り当て(対象稼働日) = 割り当て
                End If

                ' 残数ゼロなら次の品番へ
                If 月間残数(品番) = 0 Then Exit For
            End If
        Next wdIdx
NextItem1:
    Next key

    ' ==========================================
    ' セット品番の処理（ペア単位）
    ' ==========================================
    Dim セットベース As String
    Dim ペア品番群 As Object
    Dim ペア品番 As Variant
    Dim 代表品番 As String

    For Each key In セットペアマスタ.Keys
        セットベース = CStr(key)
        Set ペア品番群 = セットペアマスタ(セットベース)

        ' ペアの優先度チェック（代表品番で確認）
        For Each ペア品番 In ペア品番群.Keys
            代表品番 = CStr(ペア品番)
            Exit For
        Next ペア品番

        If 品番マスタ(代表品番)("優先度") <> 優先度 Then GoTo NextPair

        ' 系列フィルタ
        系列値 = CStr(品番マスタ(代表品番)("系列"))
        If 系列値 <> 系列 Then GoTo NextPair

        ' モール条件フィルタ
        仕様 = CStr(品番マスタ(代表品番)("仕様"))
        If モール条件 = "モール" Then
            If InStr(仕様, "モール") = 0 Then GoTo NextPair
        Else
            If InStr(仕様, "モール") > 0 Then GoTo NextPair
        End If

        ' ペア全体の残数チェック
        Dim ペア残数あり As Boolean
        ペア残数あり = False
        For Each ペア品番 In ペア品番群.Keys
            If 月間残数(CStr(ペア品番)) > 0 Then
                ペア残数あり = True
                Exit For
            End If
        Next ペア品番
        If Not ペア残数あり Then GoTo NextPair

        ' グループ制約チェック（ペア全体で共通）
        グループID = 品番マスタ(代表品番)("グループ")

        ' 稼働日ループ
        For wdIdx = 1 To workDays.Count
            稼働日 = CLng(workDays(wdIdx))
            対象稼働日 = 稼働日

            If グループID <> "" And グループ初回日.Exists(グループID) Then
                初回日 = CLng(グループ初回日(グループID))

                ' 初回日を優先的に試す（ペア全体で）
                Dim ペア初回日割り当て As Long
                ペア初回日割り当て = セットペア割り当て可能数を算出_制約付き(ペア品番群, 初回日, 品番マスタ, 月間残数, 当日割り当て, 当日型番割り当て, 転記データ, 日次目標, 誤差許容率)

                If ペア初回日割り当て > 0 Then
                    対象稼働日 = 初回日
                End If
            End If

            ' セットペア割り当て可能数算出
            割り当て = セットペア割り当て可能数を算出_制約付き(ペア品番群, 対象稼働日, 品番マスタ, 月間残数, 当日割り当て, 当日型番割り当て, 転記データ, 日次目標, 誤差許容率)

            If 割り当て > 0 Then
                ' グループ初回日記録
                If グループID <> "" And Not グループ初回日.Exists(グループID) Then
                    グループ初回日(グループID) = 対象稼働日
                End If

                ' 全ペア品番に同数割り当て
                For Each ペア品番 In ペア品番群.Keys
                    品番 = CStr(ペア品番)

                    転記キー = 品番 & "_" & 対象稼働日
                    If 転記データ.Exists(転記キー) Then
                        転記データ(転記キー) = CLng(転記データ(転記キー)) + 割り当て
                    Else
                        転記データ(転記キー) = 割り当て
                    End If

                    月間残数(品番) = CLng(月間残数(品番)) - 割り当て

                    If 当日割り当て.Exists(対象稼働日) Then
                        当日割り当て(対象稼働日) = CLng(当日割り当て(対象稼働日)) + 割り当て
                    Else
                        当日割り当て(対象稼働日) = 割り当て
                    End If
                Next ペア品番

                ' 型番累積更新（ペア全体で）
                Dim 型番 As String
                型番 = CStr(品番マスタ(代表品番)("型番"))
                If 型番 <> "" Then
                    Dim 型番日キー As String
                    型番日キー = 型番 & "_" & 対象稼働日
                    Dim ペア品番数 As Long
                    ペア品番数 = ペア品番群.Count

                    If 当日型番割り当て.Exists(型番日キー) Then
                        当日型番割り当て(型番日キー) = CLng(当日型番割り当て(型番日キー)) + 割り当て * ペア品番数
                    Else
                        当日型番割り当て(型番日キー) = 割り当て * ペア品番数
                    End If
                End If

                ' ペア全体の残数チェック
                Dim 全ペア完了 As Boolean
                全ペア完了 = True
                For Each ペア品番 In ペア品番群.Keys
                    If 月間残数(CStr(ペア品番)) > 0 Then
                        全ペア完了 = False
                        Exit For
                    End If
                Next ペア品番

                If 全ペア完了 Then Exit For
            End If
        Next wdIdx
NextPair:
    Next key
End Sub

' ==========================================
' 4. 割り当て可能数を算出（制約付き）
' ==========================================
Private Function 割り当て可能数を算出_制約付き( _
    ByVal 品番 As String, _
    ByVal 稼働日 As Long, _
    ByRef 品番マスタ As Object, _
    ByRef 月間残数 As Object, _
    ByRef 当日割り当て As Object, _
    ByRef 当日型番割り当て As Object, _
    ByRef 転記データ As Object, _
    ByVal 日次目標 As Double, _
    ByVal 誤差許容率 As Double _
) As Long

    ' 基本制約
    Dim 残数 As Long, 上限 As Long, 単位 As Long
    残数 = CLng(月間残数(品番))
    上限 = CLng(品番マスタ(品番)("上限"))
    単位 = CLng(品番マスタ(品番)("単位"))

    ' 当日既割り当て数
    Dim 当日既割り当て As Long
    当日既割り当て = 0
    If 当日割り当て.Exists(稼働日) Then
        当日既割り当て = CLng(当日割り当て(稼働日))
    End If

    ' 日次目標制約（許容範囲考慮）
    Dim 許容下限 As Long, 許容上限 As Long
    許容下限 = CLng(日次目標 * (1 - 誤差許容率 / 100))
    許容上限 = CLng(日次目標 * (1 + 誤差許容率 / 100))

    If 許容下限 < 0 Then 許容下限 = 0

    Dim 当日最大 As Long
    当日最大 = 許容上限 - 当日既割り当て
    If 当日最大 > 上限 Then 当日最大 = 上限
    If 当日最大 < 0 Then 当日最大 = 0

    ' 下限未達の場合、優先的に割り当て
    If 当日既割り当て < 許容下限 Then
        Dim 下限不足 As Long
        下限不足 = 許容下限 - 当日既割り当て
        当日最大 = 当日最大 + 下限不足
        If 当日最大 > 上限 Then 当日最大 = 上限
    End If

    ' --- 型番制約チェック（セット品のみ） ---
    Dim セット As String
    セット = CStr(品番マスタ(品番)("セット"))

    If セット = "SET" Then
        Dim 型番 As String
        型番 = CStr(品番マスタ(品番)("型番"))

        If 型番 <> "" Then
            Dim 型番上限 As Long
            型番上限 = CLng(品番マスタ(品番)("上限"))

            Dim 型番当日既割り当て As Long
            Dim 型番日キー As String
            型番日キー = 型番 & "_" & 稼働日

            If 当日型番割り当て.Exists(型番日キー) Then
                型番当日既割り当て = CLng(当日型番割り当て(型番日キー))
            Else
                型番当日既割り当て = 0
            End If

            Dim 型番残余 As Long
            型番残余 = 型番上限 - 型番当日既割り当て

            If 型番残余 <= 0 Then
                割り当て可能数を算出_制約付き = 0
                Exit Function
            End If

            If 当日最大 > 型番残余 Then
                当日最大 = 型番残余
            End If
        End If
    End If

    ' --- 新規制約1: 号口単品の分散配置 ---
    Dim 号補 As String
    号補 = CStr(品番マスタ(品番)("号/補"))

    If 号補 = "号口" And セット <> "SET" Then
        ' 当日に既に他の号口単品が配置されているかチェック
        If 当日号口単品配置済み(稼働日, 品番, 転記データ, 品番マスタ) Then
            割り当て可能数を算出_制約付き = 0
            Exit Function
        End If
    End If

    ' --- 新規制約2: 補給品と号口単品の同日配置禁止 ---
    If 号補 = "補給品" Then
        ' 補給品の場合：当日に号口単品があればNG
        If 当日号口単品あり(稼働日, 転記データ, 品番マスタ) Then
            割り当て可能数を算出_制約付き = 0
            Exit Function
        End If
    ElseIf 号補 = "号口" And セット <> "SET" Then
        ' 号口単品の場合：当日に補給品があればNG
        If 当日補給品あり(稼働日, 転記データ, 品番マスタ) Then
            割り当て可能数を算出_制約付き = 0
            Exit Function
        End If
    End If

    ' 単位制約（倍数に丸める）
    Dim 割り当て候補 As Long
    割り当て候補 = 残数
    If 割り当て候補 > 当日最大 Then 割り当て候補 = 当日最大

    割り当て候補 = Int(割り当て候補 / 単位) * 単位

    割り当て可能数を算出_制約付き = 割り当て候補
End Function

' ==========================================
' 5. セットペア割り当て可能数を算出（制約付き）
' ==========================================
Private Function セットペア割り当て可能数を算出_制約付き( _
    ByRef ペア品番群 As Object, _
    ByVal 稼働日 As Long, _
    ByRef 品番マスタ As Object, _
    ByRef 月間残数 As Object, _
    ByRef 当日割り当て As Object, _
    ByRef 当日型番割り当て As Object, _
    ByRef 転記データ As Object, _
    ByVal 日次目標 As Double, _
    ByVal 誤差許容率 As Double _
) As Long

    Dim 最小割り当て As Long
    最小割り当て = 999999

    ' 各ペア品番の割り当て可能数を個別算出
    Dim ペア品番 As Variant
    For Each ペア品番 In ペア品番群.Keys
        Dim 個別割り当て As Long
        個別割り当て = 割り当て可能数を算出_制約付き( _
            CStr(ペア品番), 稼働日, 品番マスタ, 月間残数, _
            当日割り当て, 当日型番割り当て, 転記データ, _
            日次目標, 誤差許容率)

        If 個別割り当て < 最小割り当て Then
            最小割り当て = 個別割り当て
        End If
    Next ペア品番

    セットペア割り当て可能数を算出_制約付き = 最小割り当て
End Function

' ==========================================
' 6. 当日号口単品配置済みチェック
' ==========================================
Private Function 当日号口単品配置済み( _
    ByVal 稼働日 As Long, _
    ByVal 対象品番 As String, _
    ByRef 転記データ As Object, _
    ByRef 品番マスタ As Object _
) As Boolean

    ' 転記データから当日の全品番をチェック
    Dim key As Variant, parts() As String
    Dim 品番 As String, 日 As Long
    Dim 号補 As String, セット As String

    For Each key In 転記データ.Keys
        parts = Split(CStr(key), "_")
        品番 = parts(0)
        日 = CLng(parts(1))

        ' 同じ日かつ異なる品番
        If 日 = 稼働日 And 品番 <> 対象品番 Then
            ' 品番マスタで号口単品かチェック
            If 品番マスタ.Exists(品番) Then
                号補 = CStr(品番マスタ(品番)("号/補"))
                セット = CStr(品番マスタ(品番)("セット"))

                If 号補 = "号口" And セット <> "SET" Then
                    当日号口単品配置済み = True
                    Exit Function
                End If
            End If
        End If
    Next key

    当日号口単品配置済み = False
End Function

' ==========================================
' 7. 当日補給品ありチェック
' ==========================================
Private Function 当日補給品あり( _
    ByVal 稼働日 As Long, _
    ByRef 転記データ As Object, _
    ByRef 品番マスタ As Object _
) As Boolean

    Dim key As Variant, parts() As String
    Dim 品番 As String, 日 As Long
    Dim 号補 As String

    For Each key In 転記データ.Keys
        parts = Split(CStr(key), "_")
        品番 = parts(0)
        日 = CLng(parts(1))

        If 日 = 稼働日 Then
            If 品番マスタ.Exists(品番) Then
                号補 = CStr(品番マスタ(品番)("号/補"))
                If 号補 = "補給品" Then
                    当日補給品あり = True
                    Exit Function
                End If
            End If
        End If
    Next key

    当日補給品あり = False
End Function

' ==========================================
' 8. 当日号口単品ありチェック
' ==========================================
Private Function 当日号口単品あり( _
    ByVal 稼働日 As Long, _
    ByRef 転記データ As Object, _
    ByRef 品番マスタ As Object _
) As Boolean

    Dim key As Variant, parts() As String
    Dim 品番 As String, 日 As Long
    Dim 号補 As String, セット As String

    For Each key In 転記データ.Keys
        parts = Split(CStr(key), "_")
        品番 = parts(0)
        日 = CLng(parts(1))

        If 日 = 稼働日 Then
            If 品番マスタ.Exists(品番) Then
                号補 = CStr(品番マスタ(品番)("号/補"))
                セット = CStr(品番マスタ(品番)("セット"))

                If 号補 = "号口" And セット <> "SET" Then
                    当日号口単品あり = True
                    Exit Function
                End If
            End If
        End If
    Next key

    当日号口単品あり = False
End Function

' ==========================================
' 9. 品番の最初の納期を取得
' ==========================================
Private Function 品番の最初の納期を取得( _
    ByVal 品番 As String, _
    ByRef arr展開 As Variant, _
    ByVal 成形品番列 As Long, _
    ByVal 開始列 As Long, _
    ByVal maxDay As Long _
) As Long

    Dim r As Long, d As Long, 数量 As Long

    ' 該当品番の行を探す
    For r = 1 To UBound(arr展開, 1)
        If CStr(arr展開(r, 成形品番列)) = 品番 Then
            ' 日付順に数量をチェック
            For d = 1 To maxDay
                If 開始列 + d - 1 <= UBound(arr展開, 2) Then
                    数量 = 0
                    On Error Resume Next
                    数量 = CLng(arr展開(r, 開始列 + d - 1))
                    On Error GoTo 0

                    If 数量 > 0 Then
                        品番の最初の納期を取得 = d
                        Exit Function
                    End If
                End If
            Next d
            Exit For
        End If
    Next r

    ' 見つからなかった場合
    品番の最初の納期を取得 = 0
End Function

