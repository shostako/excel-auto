Attribute VB_Name = "m転記_グラフ_流出詳細"
Option Explicit

' グラフ表示用の流出詳細データを集計・転記するマクロ
' 期間指定で_手直しと_廃棄テーブルから条件別に集計し、_入力テーブルに転記
Sub 転記_グラフ_流出詳細()
    Application.ScreenUpdating = False
    Application.StatusBar = "流出詳細データを集計中..."

    On Error GoTo ErrorHandler

    ' 期間テーブルから開始日・終了日を取得
    Dim ws期間 As Worksheet
    Set ws期間 = ThisWorkbook.Worksheets("グラフ")

    Dim lo期間 As ListObject
    Set lo期間 = ws期間.ListObjects("_期間_滝")

    Dim 開始日 As Date, 終了日 As Date
    開始日 = lo期間.ListColumns("開始日").DataBodyRange(1, 1).Value
    終了日 = lo期間.ListColumns("終了日").DataBodyRange(1, 1).Value

    ' 集計用変数初期化
    Dim 成形_流出 As Long, 塗装発見_流出 As Long, 塗装_流出 As Long, 加工流出総数_流出 As Long
    Dim 成形_廃棄 As Long, 塗装_廃棄 As Long, 加工流出総数_廃棄 As Long

    ' === _手直しテーブルから流出数を集計 ===
    Application.StatusBar = "手直しデータを集計中..."

    Dim ws手直し As Worksheet
    Set ws手直し = ThisWorkbook.Worksheets("手直し")

    Dim lo手直し As ListObject
    Set lo手直し = ws手直し.ListObjects("_手直し")

    If Not lo手直し.DataBodyRange Is Nothing Then
        Dim arr手直し As Variant
        arr手直し = lo手直し.DataBodyRange.Value

        ' 列インデックス取得
        Dim col手_日付 As Long, col手_発見 As Long, col手_発生 As Long, col手_数量 As Long
        col手_日付 = lo手直し.ListColumns("日付").Index
        col手_発見 = lo手直し.ListColumns("発見2").Index
        col手_発生 = lo手直し.ListColumns("発生").Index
        col手_数量 = lo手直し.ListColumns("数量").Index

        Dim i As Long
        For i = 1 To UBound(arr手直し, 1)
            ' 日付チェック
            If IsDate(arr手直し(i, col手_日付)) Then
                Dim 日付 As Date
                日付 = arr手直し(i, col手_日付)

                If 日付 >= 開始日 And 日付 <= 終了日 Then
                    Dim 発見 As String, 発生 As String, 数量 As Long
                    発見 = Trim(arr手直し(i, col手_発見) & "")
                    発生 = Trim(arr手直し(i, col手_発生) & "")
                    数量 = 0
                    If IsNumeric(arr手直し(i, col手_数量)) Then
                        数量 = CLng(arr手直し(i, col手_数量))
                    End If

                    ' 条件2: 発見={塗装,モール,加工} かつ 発生={成形} → 成形の流出
                    If (発見 = "塗装" Or 発見 = "モール" Or 発見 = "加工") And 発生 = "成形" Then
                        成形_流出 = 成形_流出 + 数量
                    End If

                    ' 条件3: 発見={塗装} かつ 発生={成形} → 塗装発見の流出
                    If 発見 = "塗装" And 発生 = "成形" Then
                        塗装発見_流出 = 塗装発見_流出 + 数量
                    End If

                    ' 条件4: 発見={モール,加工} かつ 発生={塗装} → 塗装の流出
                    If (発見 = "モール" Or 発見 = "加工") And 発生 = "塗装" Then
                        塗装_流出 = 塗装_流出 + 数量
                    End If

                    ' 条件5: 発見={モール,加工} かつ 発生={成形,塗装} → 加工流出総数の流出
                    If (発見 = "モール" Or 発見 = "加工") And (発生 = "成形" Or 発生 = "塗装") Then
                        加工流出総数_流出 = 加工流出総数_流出 + 数量
                    End If
                End If
            End If
        Next i
    End If

    ' === _廃棄テーブルから廃棄数を集計 ===
    Application.StatusBar = "廃棄データを集計中..."

    Dim ws廃棄 As Worksheet
    Set ws廃棄 = ThisWorkbook.Worksheets("廃棄")

    Dim lo廃棄 As ListObject
    Set lo廃棄 = ws廃棄.ListObjects("_廃棄")

    If Not lo廃棄.DataBodyRange Is Nothing Then
        Dim arr廃棄 As Variant
        arr廃棄 = lo廃棄.DataBodyRange.Value

        ' 列インデックス取得
        Dim col廃_日付 As Long, col廃_工程 As Long, col廃_件数 As Long
        col廃_日付 = lo廃棄.ListColumns("日付").Index
        col廃_工程 = lo廃棄.ListColumns("工程").Index
        col廃_件数 = lo廃棄.ListColumns("件数").Index

        For i = 1 To UBound(arr廃棄, 1)
            ' 日付チェック
            If IsDate(arr廃棄(i, col廃_日付)) Then
                日付 = arr廃棄(i, col廃_日付)

                If 日付 >= 開始日 And 日付 <= 終了日 Then
                    Dim 工程 As String, 件数 As Long
                    工程 = Trim(arr廃棄(i, col廃_工程) & "")
                    件数 = 0
                    If IsNumeric(arr廃棄(i, col廃_件数)) Then
                        件数 = CLng(arr廃棄(i, col廃_件数))
                    End If

                    ' 条件6: 工程={成形} → 成形の廃棄
                    If 工程 = "成形" Then
                        成形_廃棄 = 成形_廃棄 + 件数
                    End If

                    ' 条件7: 工程={塗装} → 塗装の廃棄
                    If 工程 = "塗装" Then
                        塗装_廃棄 = 塗装_廃棄 + 件数
                    End If

                    ' 条件8: 工程={成形,塗装} → 加工流出総数の廃棄
                    If 工程 = "成形" Or 工程 = "塗装" Then
                        加工流出総数_廃棄 = 加工流出総数_廃棄 + 件数
                    End If
                End If
            End If
        Next i
    End If

    ' === 出力テーブルに転記 ===
    Application.StatusBar = "結果を転記中..."

    Dim lo出力 As ListObject
    Set lo出力 = ws期間.ListObjects("_入力")

    If Not lo出力.DataBodyRange Is Nothing Then
        Dim arr出力 As Variant
        arr出力 = lo出力.DataBodyRange.Value

        ' 列インデックス取得
        Dim col出_工程 As Long, col出_流出 As Long, col出_廃棄 As Long
        col出_工程 = lo出力.ListColumns("工程").Index
        col出_流出 = lo出力.ListColumns("流出").Index
        col出_廃棄 = lo出力.ListColumns("廃棄").Index

        ' 工程別に該当行を探して転記
        For i = 1 To UBound(arr出力, 1)
            Dim 工程名 As String
            工程名 = Trim(arr出力(i, col出_工程) & "")

            Select Case 工程名
                Case "成形"
                    arr出力(i, col出_流出) = 成形_流出
                    arr出力(i, col出_廃棄) = 成形_廃棄
                Case "塗装発見"
                    arr出力(i, col出_流出) = 塗装発見_流出
                Case "塗装"
                    arr出力(i, col出_流出) = 塗装_流出
                    arr出力(i, col出_廃棄) = 塗装_廃棄
                Case "加工流出総数"
                    arr出力(i, col出_流出) = 加工流出総数_流出
                    arr出力(i, col出_廃棄) = 加工流出総数_廃棄
            End Select
        Next i

        ' 一括書き込み
        lo出力.DataBodyRange.Value = arr出力
    End If

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
End Sub
