Attribute VB_Name = "m転記_見せる表_加工"
Option Explicit

' ========================================
' マクロ名: 転記_見せる表_加工
' 処理概要: 工程内不良データを統合し、期間別の見せる表シートを作成
' ソーステーブル:
'   - シート「加工G」テーブル「_流出G_加工_期間1」（流出）
'   - シート「加工T」テーブル「_手直しT_加工_期間1」（手直し）
'   - シート「加工H」テーブル「_廃棄H_加工_期間1」（廃棄）
' 出力: 原紙加工シートをコピーして新シート作成（シート名: 加工M.d‾M.d）
'
' 【処理の特徴】
' - 工程内不良データを1枚のシートに統合
' - 動的行削除: 不良項目数に応じて余分な行を自動削除
' - ワースト強調: 上位3項目を太字+2ポイント
' - 全体不良数の自動計算: 手直し+廃棄の合計
' ========================================

Sub 転記_見せる表_加工()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.StatusBar = "見せる表を作成中..."

    ' ============================================
    ' 1. 期間情報の取得
    ' ============================================
    ' 理由: 加工Gの期間1を基準にシート作成
    ' ポイント: 加工T/H/Gは一括転記マクロで期間同期済み

    Dim 期間Tbl As ListObject
    Set 期間Tbl = Sheets("加工G").ListObjects("_集計期間加工G")

    Dim 開始日 As Date, 終了日 As Date
    Dim i As Long

    ' 期間="期間1"の行を検索（加工G）
    For i = 1 To 期間Tbl.ListRows.Count
        If 期間Tbl.ListColumns("期間").DataBodyRange(i, 1).Value = "期間1" Then
            開始日 = 期間Tbl.ListColumns("開始日").DataBodyRange(i, 1).Value
            終了日 = 期間Tbl.ListColumns("終了日").DataBodyRange(i, 1).Value
            Exit For
        End If
    Next i

    ' ============================================
    ' 2. シート名決定と既存シート削除
    ' ============================================
    ' 理由: 同名シートが存在する場合は上書きする仕様
    ' ポイント: 確認メッセージなしで削除（業務効率優先）

    Dim newName As String
    newName = "加工" & Format(開始日, "M.d") & "‾" & Format(終了日, "M.d")

    ' 同名シートが存在する場合は削除（確認なし）
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = Sheets(newName)
    On Error GoTo ErrorHandler
    If Not sh Is Nothing Then
        Application.DisplayAlerts = False
        sh.Delete
        Application.DisplayAlerts = True
        Set sh = Nothing
    End If

    ' ============================================
    ' 3. 原紙シートのコピーと期間入力
    ' ============================================
    ' 理由: 原紙加工シートにレイアウトとセル結合が定義済み
    ' ポイント: B2セルに期間を入力してタイトル表示

    Sheets("原紙加工").Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = newName
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' B2セルに期間を入力
    ws.Range("B2").Value = "加工 " & Format(開始日, "M/d") & "‾" & Format(終了日, "M/d")

    ' ============================================
    ' 4. ソーステーブル参照
    ' ============================================
    ' 理由: 各テーブルから不良データを取得
    ' ポイント: 加工は工程内不良の概念がない

    Dim 流出Tbl As ListObject
    Set 流出Tbl = Sheets("加工G").ListObjects("_流出G_加工_期間1")

    Dim 手直しTbl As ListObject
    Set 手直しTbl = Sheets("加工T").ListObjects("_手直しT_加工_期間1")

    Dim 廃棄Tbl As ListObject
    Set 廃棄Tbl = Sheets("加工H").ListObjects("_廃棄H_加工_期間1")

    ' ============================================
    ' 5. 不良項目数のカウント
    ' ============================================
    ' 理由: 動的行削除のために項目数を事前把握
    ' ポイント: 「ショット数」「不良数」など集計行を除外してカウント

    Dim 不良項目数 As Long
    不良項目数 = 0
    Dim item As String

    For i = 1 To 流出Tbl.ListRows.Count
        item = 流出Tbl.ListColumns("項目").DataBodyRange(i, 1).Value
        If Trim(item) <> "ショット数" And Trim(item) <> "不良数" Then
            不良項目数 = 不良項目数 + 1
        End If
    Next i

    Application.StatusBar = "不良項目数: " & 不良項目数 & "件"

    ' ============================================
    ' 6. 工程内不良処理
    ' ============================================

    ' --------------------------------------------
    ' 6-1. 余分な行を削除
    ' --------------------------------------------
    ' 理由: 原紙は最大20項目分の行があるため、実際の項目数に合わせて削除
    ' ポイント: 16行目から削除開始（15行目は不良項目の開始行）

    Application.StatusBar = "行を調整中..."
    Dim deleteCount As Long
    deleteCount = 20 - 不良項目数

    If deleteCount > 0 Then
        ws.Rows("16:" & (15 + deleteCount)).Delete Shift:=xlUp
    End If

    ' --------------------------------------------
    ' 6-2. 列マッピング設定
    ' --------------------------------------------
    ' 理由: セル結合状態が行ごとに異なるため、複数の列マッピングが必要
    ' ポイント: 手直しは補給品なし（8分類）、廃棄・流出は補給品あり（9分類）

    ' ショット数用（7行目）- G列開始、補給品あり
    Dim colMapショット数 As Object
    Set colMapショット数 = CreateObject("Scripting.Dictionary")
    colMapショット数("合計") = "G"
    colMapショット数("58050FrLH") = "K"
    colMapショット数("58050FrRH") = "N"
    colMapショット数("58050RrLH") = "Q"
    colMapショット数("58050RrRH") = "T"
    colMapショット数("28050FrLH") = "W"
    colMapショット数("28050FrRH") = "Z"
    colMapショット数("28050RrLH") = "AC"
    colMapショット数("28050RrRH") = "AF"
    colMapショット数("補給品") = "AI"

    ' 全体不良数用（8行目）- G列開始、補給品あり
    Dim colMap全体不良数 As Object
    Set colMap全体不良数 = CreateObject("Scripting.Dictionary")
    colMap全体不良数("合計") = "G"
    colMap全体不良数("58050FrLH") = "K"
    colMap全体不良数("58050FrRH") = "N"
    colMap全体不良数("58050RrLH") = "Q"
    colMap全体不良数("58050RrRH") = "T"
    colMap全体不良数("28050FrLH") = "W"
    colMap全体不良数("28050FrRH") = "Z"
    colMap全体不良数("28050RrLH") = "AC"
    colMap全体不良数("28050RrRH") = "AF"
    colMap全体不良数("補給品") = "AI"

    ' 不良数行専用（11行目・14行目）- D列開始
    Dim colMap不良数 As Object
    Set colMap不良数 = CreateObject("Scripting.Dictionary")
    colMap不良数("合計") = "D"
    colMap不良数("58050FrLH") = "K"
    colMap不良数("58050FrRH") = "N"
    colMap不良数("58050RrLH") = "Q"
    colMap不良数("58050RrRH") = "T"
    colMap不良数("28050FrLH") = "W"
    colMap不良数("28050FrRH") = "Z"
    colMap不良数("28050RrLH") = "AC"
    colMap不良数("28050RrRH") = "AF"
    colMap不良数("補給品") = "AI"

    ' 12行目用（手直し）- G列開始、補給品なし
    Dim colMap手直し As Object
    Set colMap手直し = CreateObject("Scripting.Dictionary")
    colMap手直し("合計") = "G"
    colMap手直し("58050FrLH") = "L"
    colMap手直し("58050FrRH") = "O"
    colMap手直し("58050RrLH") = "R"
    colMap手直し("58050RrRH") = "U"
    colMap手直し("28050FrLH") = "X"
    colMap手直し("28050FrRH") = "AA"
    colMap手直し("28050RrLH") = "AD"
    colMap手直し("28050RrRH") = "AG"

    ' 13行目用（廃棄）- G列開始、補給品あり
    Dim colMap廃棄 As Object
    Set colMap廃棄 = CreateObject("Scripting.Dictionary")
    colMap廃棄("合計") = "G"
    colMap廃棄("58050FrLH") = "L"
    colMap廃棄("58050FrRH") = "O"
    colMap廃棄("58050RrLH") = "R"
    colMap廃棄("58050RrRH") = "U"
    colMap廃棄("28050FrLH") = "X"
    colMap廃棄("28050FrRH") = "AA"
    colMap廃棄("28050RrLH") = "AD"
    colMap廃棄("28050RrRH") = "AG"
    colMap廃棄("補給品") = "AJ"

    ' 不良項目用（15行目以降）- G列開始、補給品あり
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    colMap("合計") = "G"
    colMap("58050FrLH") = "L"
    colMap("58050FrRH") = "O"
    colMap("58050RrLH") = "R"
    colMap("58050RrRH") = "U"
    colMap("28050FrLH") = "X"
    colMap("28050FrRH") = "AA"
    colMap("28050RrLH") = "AD"
    colMap("28050RrRH") = "AG"
    colMap("補給品") = "AJ"

    ' --------------------------------------------
    ' 6-3. データ転記処理
    ' --------------------------------------------
    ' 理由: ショット数、手直し、廃棄、流出データを原紙の指定行に転記
    ' ポイント: 手直し・廃棄不良数を辞書に保存（全体不良数計算用）

    Application.StatusBar = "データを転記中..."

    ' 手直し・廃棄不良数を保存する辞書（全体不良数計算用）
    Dim 手直し不良数 As Object
    Set 手直し不良数 = CreateObject("Scripting.Dictionary")

    Dim 廃棄不良数 As Object
    Set 廃棄不良数 = CreateObject("Scripting.Dictionary")

    ' 手直しテーブルから不良数を12行目に転記
    For i = 1 To 手直しTbl.ListRows.Count
        item = 手直しTbl.ListColumns("項目").DataBodyRange(i, 1).Value
        If Trim(item) = "不良数" Then
            Debug.Print "→ 手直し不良数を12行目に転記"
            Call 転記行データ(ws, 12, 手直しTbl, i, colMap手直し)
            ' 手直し不良数を保存（全体不良数計算用）
            Dim keyCol As Variant
            For Each keyCol In colMap手直し.Keys
                手直し不良数(CStr(keyCol)) = 手直しTbl.ListColumns(CStr(keyCol)).DataBodyRange(i, 1).Value
            Next keyCol
            Exit For
        End If
    Next i

    ' 廃棄テーブルから不良数を13行目に転記
    For i = 1 To 廃棄Tbl.ListRows.Count
        item = 廃棄Tbl.ListColumns("項目").DataBodyRange(i, 1).Value
        If Trim(item) = "不良数" Then
            Debug.Print "→ 廃棄不良数を13行目に転記"
            Call 転記行データ(ws, 13, 廃棄Tbl, i, colMap廃棄)
            ' 廃棄不良数を保存（全体不良数計算用）
            Dim keyCol2 As Variant
            For Each keyCol2 In colMap廃棄.Keys
                廃棄不良数(CStr(keyCol2)) = 廃棄Tbl.ListColumns(CStr(keyCol2)).DataBodyRange(i, 1).Value
            Next keyCol2
            Exit For
        End If
    Next i

    ' 流出テーブルからデータ転記
    Application.StatusBar = "流出データを転記中..."

    ' 全項目行の全体最大値を事前計算
    ' 理由: ワースト強調の判定基準を統一（全項目の中での相対評価）
    ' ポイント: 各項目行で個別判定せず、全体基準で判定
    Dim 全体最大値 As Double
    全体最大値 = Get全体最大値(流出Tbl, colMap, "ショット数,不良数,その他")
    Debug.Print "全体最大値: " & 全体最大値

    Dim targetRow As Long
    targetRow = 15
    Dim 不良数行Index As Long
    不良数行Index = 0

    For i = 1 To 流出Tbl.ListRows.Count
        item = 流出Tbl.ListColumns("項目").DataBodyRange(i, 1).Value
        Debug.Print "行" & i & ": [" & item & "] (長さ:" & Len(item) & ")"

        If Trim(item) = "ショット数" Then
            Debug.Print "→ ショット数を7行目に転記"
            ' 7行目に転記
            Call 転記行データ(ws, 7, 流出Tbl, i, colMapショット数)
        ElseIf Trim(item) = "不良数" Then
            不良数行Index = i
            Debug.Print "→ 不良数を11行目・14行目に転記"
            ' 11行目に転記
            Call 転記行データ(ws, 11, 流出Tbl, i, colMap不良数)
            ' 14行目に転記
            Call 転記行データ(ws, 14, 流出Tbl, i, colMap不良数)
            ' フォント強調を適用
            Dim 強調キー不良数 As Collection
            Set 強調キー不良数 = Get強調対象キー(流出Tbl, i, colMap不良数)
            Call フォント強調適用(ws, Array(11, 14), 強調キー不良数, colMap不良数)
        Else
            Debug.Print "→ 不良項目を" & targetRow & "行目に転記"
            ' 不良項目を15行目以降に転記
            ws.Range("E" & targetRow).Value = item
            Call 転記行データ(ws, targetRow, 流出Tbl, i, colMap)
            ' フォント強調を適用（全体最大値基準）
            Dim 強調キー項目 As Collection
            Set 強調キー項目 = Get強調対象キー_全体判定(流出Tbl, i, colMap, 全体最大値)
            Call フォント強調適用(ws, Array(targetRow), 強調キー項目, colMap)
            targetRow = targetRow + 1
        End If
    Next i

    ' ============================================
    ' 7. 全体不良数を計算して8行目に転記
    ' ============================================
    ' 理由: 手直し不良数と廃棄不良数の合計を表示
    ' ポイント: 辞書に保存した値を使って計算

    Application.StatusBar = "全体不良数を計算中..."
    Dim keyCol3 As Variant
    Dim 手直し値 As Variant, 廃棄値 As Variant
    Dim 合計値 As Double

    For Each keyCol3 In colMap全体不良数.Keys
        ' 手直し不良数と廃棄不良数を取得
        手直し値 = 0
        廃棄値 = 0

        If 手直し不良数.Exists(CStr(keyCol3)) Then
            手直し値 = 手直し不良数(CStr(keyCol3))
        End If

        If 廃棄不良数.Exists(CStr(keyCol3)) Then
            廃棄値 = 廃棄不良数(CStr(keyCol3))
        End If

        ' 合計を計算
        合計値 = CDbl(手直し値) + CDbl(廃棄値)

        ' 8行目に転記
        ws.Range(colMap全体不良数(keyCol3) & "8").Value = 合計値
        Debug.Print "  全体不良数: " & keyCol3 & " = " & 手直し値 & " + " & 廃棄値 & " = " & 合計値
    Next keyCol3

    ' ============================================
    ' 8. 処理完了（最終セル選択と画面更新復元）
    ' ============================================
    ' 理由: A1セルを選択して見やすい位置で終了
    ' ポイント: ScreenUpdatingを戻して画面を正常化

    Application.StatusBar = False
    Application.ScreenUpdating = True

    ' A1セルを選択して終了
    ws.Range("A1").Select

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
End Sub

' ============================================
' Private Sub: 転記行データ
' ============================================
' 概要: colMapに従ってソーステーブルから指定行にデータを転記
' 引数:
'   - ws: 転記先ワークシート
'   - targetRow: 転記先行番号
'   - srcTbl: ソーステーブル
'   - srcRowIndex: ソーステーブルの行インデックス
'   - colMap: 列マッピング辞書
'
Private Sub 転記行データ(ws As Worksheet, targetRow As Long, _
                        srcTbl As ListObject, srcRowIndex As Long, _
                        colMap As Object)
    ' colMapに従ってデータを転記
    Dim key As Variant
    Dim val As Variant
    Debug.Print "  転記行データ開始: 行" & targetRow
    For Each key In colMap.Keys
        val = srcTbl.ListColumns(key).DataBodyRange(srcRowIndex, 1).Value
        ws.Range(colMap(key) & targetRow).Value = val
        Debug.Print "    " & key & " → " & colMap(key) & targetRow & " = " & val
    Next key
End Sub

' ============================================
' Private Function: Get全体最大値
' ============================================
' 概要: 全項目行の全セルから最大値を取得（除外項目を除く）
' 引数:
'   - srcTbl: ソーステーブル
'   - colMap: 列マッピング辞書
'   - 除外項目: カンマ区切りの除外項目リスト
' 戻り値: 全体最大値（Double）
'
Private Function Get全体最大値(srcTbl As ListObject, colMap As Object, _
                            除外項目 As String) As Double
    ' 全項目行の全セルから最大値を取得
    ' 除外項目：カンマ区切り（例: "ショット数,不良数"）

    Get全体最大値 = 0

    ' 除外項目リストを配列化
    Dim 除外Array As Variant
    除外Array = Split(除外項目, ",")

    ' 合計列を除く9分類のキー
    Dim 対象キー As Variant
    対象キー = Array("58050FrLH", "58050FrRH", "58050RrLH", "58050RrRH", _
                    "28050FrLH", "28050FrRH", "28050RrLH", "28050RrRH", "補給品")

    Dim i As Long
    Dim item As String
    Dim k As Variant
    Dim v As Variant
    Dim 除外フラグ As Boolean
    Dim j As Long

    For i = 1 To srcTbl.ListRows.Count
        item = srcTbl.ListColumns("項目").DataBodyRange(i, 1).Value

        ' 除外項目チェック
        除外フラグ = False
        For j = LBound(除外Array) To UBound(除外Array)
            If Trim(item) = Trim(除外Array(j)) Then
                除外フラグ = True
                Exit For
            End If
        Next j

        If Not 除外フラグ Then
            ' 各列の値をチェック
            For Each k In 対象キー
                If colMap.Exists(CStr(k)) Then
                    v = srcTbl.ListColumns(CStr(k)).DataBodyRange(i, 1).Value
                    If IsNumeric(v) Then
                        If CDbl(v) > 0 And CDbl(v) > Get全体最大値 Then
                            Get全体最大値 = CDbl(v)
                        End If
                    End If
                End If
            Next k
        End If
    Next i

End Function

' ============================================
' Private Function: Get強調対象キー
' ============================================
' 概要: ワースト上位のキーを判定して返す（不良数行用）
' 条件: 最大値の70%以上、0除外、最大値5未満は対象外、上限3つ
' 引数:
'   - srcTbl: ソーステーブル
'   - srcRowIndex: ソーステーブルの行インデックス
'   - colMap: 列マッピング辞書
' 戻り値: 強調対象キーのCollection
'
Private Function Get強調対象キー(srcTbl As ListObject, srcRowIndex As Long, _
                              colMap As Object) As Collection
    ' ワースト上位のキーを判定して返す（不良数行用）
    ' 条件：最大値の70%以上、0除外、最大値5未満は対象外、上限3つ

    Set Get強調対象キー = New Collection

    ' 合計列を除く9分類のキー
    Dim 対象キー As Variant
    対象キー = Array("58050FrLH", "58050FrRH", "58050RrLH", "58050RrRH", _
                    "28050FrLH", "28050FrRH", "28050RrLH", "28050RrRH", "補給品")

    ' 各キーの値を取得
    Dim 値リスト As Object
    Set 値リスト = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    Dim v As Variant
    Dim 最大値 As Double
    最大値 = 0

    For Each k In 対象キー
        If colMap.Exists(CStr(k)) Then
            v = srcTbl.ListColumns(CStr(k)).DataBodyRange(srcRowIndex, 1).Value
            If IsNumeric(v) Then
                If CDbl(v) > 0 Then  ' 0を除外
                    値リスト(CStr(k)) = CDbl(v)
                    If CDbl(v) > 最大値 Then
                        最大値 = CDbl(v)
                    End If
                End If
            End If
        End If
    Next k

    ' 最大値が5未満なら対象外
    If 最大値 < 5 Then
        Exit Function
    End If

    ' 70%閾値を計算
    Dim 閾値 As Double
    閾値 = 最大値 * 0.7

    ' 閾値以上の値を持つキーを抽出して降順ソート
    Dim 候補 As Object
    Set 候補 = CreateObject("Scripting.Dictionary")

    For Each k In 値リスト.Keys
        If 値リスト(k) >= 閾値 Then
            候補(CStr(k)) = 値リスト(k)
        End If
    Next k

    ' 降順ソート（簡易版：配列に変換して比較）
    Dim ソート済み As Object
    Set ソート済み = CreateObject("Scripting.Dictionary")

    Dim maxKey As String
    Dim maxVal As Double
    Dim Count As Long
    Count = 0

    Do While 候補.Count > 0 And Count < 3  ' 上限3つ
        maxVal = -1
        maxKey = ""
        For Each k In 候補.Keys
            If 候補(k) > maxVal Then
                maxVal = 候補(k)
                maxKey = CStr(k)
            End If
        Next k

        If maxKey <> "" Then
            Get強調対象キー.Add maxKey
            候補.Remove maxKey
            Count = Count + 1
        Else
            Exit Do
        End If
    Loop

End Function

' ============================================
' Private Function: Get強調対象キー_全体判定
' ============================================
' 概要: ワースト上位のキーを判定して返す（全項目行での最大値基準）
' 条件: 全体最大値の70%以上、0除外、全体最大値5未満は対象外、上限3つ
' 引数:
'   - srcTbl: ソーステーブル
'   - srcRowIndex: ソーステーブルの行インデックス
'   - colMap: 列マッピング辞書
'   - 全体最大値: 全項目行での最大値
' 戻り値: 強調対象キーのCollection
'
Private Function Get強調対象キー_全体判定(srcTbl As ListObject, srcRowIndex As Long, _
                                      colMap As Object, 全体最大値 As Double) As Collection
    ' ワースト上位のキーを判定して返す（全項目行での最大値基準）
    ' 条件：全体最大値の70%以上、0除外、全体最大値5未満は対象外、上限3つ

    Set Get強調対象キー_全体判定 = New Collection

    ' 全体最大値が5未満なら対象外
    If 全体最大値 < 5 Then
        Exit Function
    End If

    ' 合計列を除く9分類のキー
    Dim 対象キー As Variant
    対象キー = Array("58050FrLH", "58050FrRH", "58050RrLH", "58050RrRH", _
                    "28050FrLH", "28050FrRH", "28050RrLH", "28050RrRH", "補給品")

    ' 70%閾値を計算
    Dim 閾値 As Double
    閾値 = 全体最大値 * 0.7

    ' 各キーの値を取得して閾値以上のものを候補に
    Dim 候補 As Object
    Set 候補 = CreateObject("Scripting.Dictionary")

    Dim k As Variant
    Dim v As Variant

    For Each k In 対象キー
        If colMap.Exists(CStr(k)) Then
            v = srcTbl.ListColumns(CStr(k)).DataBodyRange(srcRowIndex, 1).Value
            If IsNumeric(v) Then
                If CDbl(v) > 0 And CDbl(v) >= 閾値 Then  ' 0を除外、閾値以上
                    候補(CStr(k)) = CDbl(v)
                End If
            End If
        End If
    Next k

    ' 降順ソート（簡易版：配列に変換して比較）
    Dim maxKey As String
    Dim maxVal As Double
    Dim Count As Long
    Count = 0

    Do While 候補.Count > 0 And Count < 3  ' 上限3つ
        maxVal = -1
        maxKey = ""
        For Each k In 候補.Keys
            If 候補(k) > maxVal Then
                maxVal = 候補(k)
                maxKey = CStr(k)
            End If
        Next k

        If maxKey <> "" Then
            Get強調対象キー_全体判定.Add maxKey
            候補.Remove maxKey
            Count = Count + 1
        Else
            Exit Do
        End If
    Loop

End Function

' ============================================
' Private Sub: フォント強調適用
' ============================================
' 概要: 指定された行とキーのセルのフォントサイズを+2ポイント、太字
' 引数:
'   - ws: ワークシート
'   - targetRows: 対象行番号の配列
'   - 強調キー: 強調対象キーのCollection
'   - colMap: 列マッピング辞書
'
Private Sub フォント強調適用(ws As Worksheet, targetRows As Variant, _
                          強調キー As Collection, colMap As Object)
    ' 指定された行とキーのセルのフォントサイズを+2ポイント、太字

    If 強調キー.Count = 0 Then
        Exit Sub
    End If

    Dim r As Variant
    Dim k As Variant
    Dim cellAddr As String
    Dim currentSize As Double

    For Each r In targetRows
        For Each k In 強調キー
            If colMap.Exists(CStr(k)) Then
                cellAddr = colMap(CStr(k)) & CStr(r)
                currentSize = ws.Range(cellAddr).Font.Size
                ws.Range(cellAddr).Font.Size = currentSize + 2
                ws.Range(cellAddr).Font.Bold = True
                Debug.Print "    フォント強調: " & cellAddr & " (" & currentSize & " → " & (currentSize + 2) & ", 太字)"
            End If
        Next k
    Next r

End Sub
