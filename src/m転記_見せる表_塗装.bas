Attribute VB_Name = "m転記_見せる表_塗装"
Option Explicit

Sub 転記_見せる表_塗装()
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Application.StatusBar = "見せる表を作成中..."

    ' 1. 期間情報の取得
    Dim 期間Tbl As ListObject
    Set 期間Tbl = Sheets("塗装G").ListObjects("_集計期間塗装G")

    Dim 開始日 As Date, 終了日 As Date
    Dim i As Long

    ' 期間="期間1"の行を検索（塗装G）
    For i = 1 To 期間Tbl.ListRows.Count
        If 期間Tbl.ListColumns("期間").DataBodyRange(i, 1).Value = "期間1" Then
            開始日 = 期間Tbl.ListColumns("開始日").DataBodyRange(i, 1).Value
            終了日 = 期間Tbl.ListColumns("終了日").DataBodyRange(i, 1).Value
            Exit For
        End If
    Next i

    ' 塗装NWの期間情報と照合
    Dim 期間TblNW As ListObject
    Set 期間TblNW = Sheets("塗装NW").ListObjects("_集計期間日報塗装W")

    Dim 開始日NW As Date, 終了日NW As Date
    Dim 期間1見つかったNW As Boolean
    期間1見つかったNW = False

    For i = 1 To 期間TblNW.ListRows.Count
        If 期間TblNW.ListColumns("期間").DataBodyRange(i, 1).Value = "期間1" Then
            開始日NW = 期間TblNW.ListColumns("開始日").DataBodyRange(i, 1).Value
            終了日NW = 期間TblNW.ListColumns("終了日").DataBodyRange(i, 1).Value
            期間1見つかったNW = True
            Exit For
        End If
    Next i

    If Not 期間1見つかったNW Then
        MsgBox "エラー: 塗装NWシートに「期間1」が見つかりません。", vbCritical
        GoTo ErrorHandler
    End If

    ' 期間一致チェック
    If 開始日NW <> 開始日 Or 終了日NW <> 終了日 Then
        MsgBox "エラー: 塗装NWと塗装Gの期間1が一致しません。" & vbCrLf & _
               "塗装NW: " & Format(開始日NW, "yyyy/mm/dd") & " ～ " & Format(終了日NW, "yyyy/mm/dd") & vbCrLf & _
               "塗装G: " & Format(開始日, "yyyy/mm/dd") & " ～ " & Format(終了日, "yyyy/mm/dd"), vbCritical
        GoTo ErrorHandler
    End If

    ' 2. シート名決定と既存シート削除
    Dim newName As String
    newName = "塗装" & Format(開始日, "M.d") & "‾" & Format(終了日, "M.d")

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

    ' 原紙シートをコピー
    Sheets("原紙塗装").Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = newName
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' B2セルに期間を入力
    ws.Range("B2").Value = "塗装 " & Format(開始日, "M/d") & "～" & Format(終了日, "M/d")

    ' 3. ソーステーブル参照
    Dim 流出Tbl As ListObject
    Set 流出Tbl = Sheets("塗装G").ListObjects("_流出G_塗装_期間1")

    Dim 手直しTbl As ListObject
    Set 手直しTbl = Sheets("塗装T").ListObjects("_手直しT_塗装_期間1")

    Dim 廃棄Tbl As ListObject
    Set 廃棄Tbl = Sheets("塗装H").ListObjects("_廃棄H_塗装_期間1")

    ' 工程内不良用（塗装NWソース）
    Dim 工程内Tbl As ListObject
    Set 工程内Tbl = Sheets("塗装NW").ListObjects("_日報W_塗装_期間1")

    If 工程内Tbl Is Nothing Then
        MsgBox "エラー: シート「塗装NW」にテーブル「_日報W_塗装_期間1」が見つかりません。", vbCritical
        GoTo ErrorHandler
    End If

    If 工程内Tbl.DataBodyRange Is Nothing Then
        MsgBox "エラー: テーブル「_日報W_塗装_期間1」にデータがありません。", vbCritical
        GoTo ErrorHandler
    End If

    ' 4. 不良項目数のカウント（後工程流出）
    Dim 不良項目数 As Long
    不良項目数 = 0
    Dim item As String

    For i = 1 To 流出Tbl.ListRows.Count
        item = 流出Tbl.ListColumns("項目").DataBodyRange(i, 1).Value
        If Trim(item) <> "ショット数" And Trim(item) <> "不良数" Then
            不良項目数 = 不良項目数 + 1
        End If
    Next i

    ' 工程内不良項目数のカウント（塗装NWソース）
    Dim 不良項目数NW As Long
    不良項目数NW = 0
    Dim itemNW As String

    For i = 1 To 工程内Tbl.ListRows.Count
        itemNW = 工程内Tbl.ListColumns("項目").DataBodyRange(i, 1).Value
        ' ショット数、不良数、リコート、廃棄を除外
        If Trim(itemNW) <> "ショット数" And Trim(itemNW) <> "不良数" _
           And Trim(itemNW) <> "リコート" And Trim(itemNW) <> "廃棄" Then
            不良項目数NW = 不良項目数NW + 1
        End If
    Next i

    Application.StatusBar = "工程内不良項目数: " & 不良項目数NW & "件、後工程流出項目数: " & 不良項目数 & "件"

    ' ============================================
    ' 5. 後工程流出処理（既存機能）
    ' ============================================

    ' 5-1. 後工程流出の余分な行を削除（41行目から）
    Application.StatusBar = "後工程流出の行を調整中..."
    Dim deleteCount As Long
    deleteCount = 20 - 不良項目数

    If deleteCount > 0 Then
        ws.Rows("41:" & (40 + deleteCount)).Delete Shift:=xlUp
    End If

    ' 5-2. 後工程流出用の列マッピング設定
    ' 不良項目用（40行目以降）
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

    ' 不良数行専用（39行目・36行目）- セル結合が異なるため
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

    ' 37行目用（手直し）- G列開始、補給品なし
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

    ' 38行目用（廃棄）- G列開始、補給品あり
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

    ' 5-3. 後工程流出データの転記処理
    Application.StatusBar = "後工程流出データを転記中..."

    ' 後工程不良数を保存する辞書（全体不良数計算用）
    Dim 後工程不良数 As Object
    Set 後工程不良数 = CreateObject("Scripting.Dictionary")

    ' 工程内不良数を保存する辞書（全体不良数計算用）
    Dim 工程内不良数 As Object
    Set 工程内不良数 = CreateObject("Scripting.Dictionary")

    ' 手直しテーブルから不良数を37行目に転記
    For i = 1 To 手直しTbl.ListRows.Count
        item = 手直しTbl.ListColumns("項目").DataBodyRange(i, 1).Value
        If Trim(item) = "不良数" Then
            Debug.Print "→ 手直し不良数を37行目に転記"
            Call 転記行データ(ws, 37, 手直しTbl, i, colMap手直し)
            Exit For
        End If
    Next i

    ' 廃棄テーブルから不良数を38行目に転記
    For i = 1 To 廃棄Tbl.ListRows.Count
        item = 廃棄Tbl.ListColumns("項目").DataBodyRange(i, 1).Value
        If Trim(item) = "不良数" Then
            Debug.Print "→ 廃棄不良数を38行目に転記"
            Call 転記行データ(ws, 38, 廃棄Tbl, i, colMap廃棄)
            Exit For
        End If
    Next i

    ' 8. 流出テーブルからデータ転記
    Application.StatusBar = "流出データを転記中..."

    ' 全項目行の全体最大値を事前計算
    Dim 後工程全体最大値 As Double
    後工程全体最大値 = Get全体最大値(流出Tbl, colMap, "ショット数,不良数,その他")
    Debug.Print "後工程全体最大値: " & 後工程全体最大値

    Dim targetRow As Long
    targetRow = 40
    Dim 不良数行Index As Long
    不良数行Index = 0

    For i = 1 To 流出Tbl.ListRows.Count
        item = 流出Tbl.ListColumns("項目").DataBodyRange(i, 1).Value
        Debug.Print "行" & i & ": [" & item & "] (長さ:" & Len(item) & ")"

        If Trim(item) = "不良数" Then
            不良数行Index = i
            Debug.Print "→ 不良数を36行目・39行目に転記"
            ' 36行目に転記（追加機能）
            Call 転記行データ(ws, 36, 流出Tbl, i, colMap不良数)
            ' 39行目に転記（既存機能）
            Call 転記行データ(ws, 39, 流出Tbl, i, colMap不良数)
            ' 後工程不良数を保存（全体不良数計算用）
            Dim keyCol As Variant
            For Each keyCol In colMap不良数.Keys
                後工程不良数(CStr(keyCol)) = 流出Tbl.ListColumns(CStr(keyCol)).DataBodyRange(i, 1).Value
            Next keyCol
            ' フォント強調を適用
            Dim 強調キー後工程不良数 As Collection
            Set 強調キー後工程不良数 = Get強調対象キー(流出Tbl, i, colMap不良数)
            Call フォント強調適用(ws, Array(36, 39), 強調キー後工程不良数, colMap不良数)
        ElseIf Trim(item) <> "ショット数" And Trim(item) <> "不良数" Then
            Debug.Print "→ 不良項目を" & targetRow & "行目に転記"
            ' 不良項目を40行目以降に転記
            ws.Range("E" & targetRow).Value = item
            Call 転記行データ(ws, targetRow, 流出Tbl, i, colMap)
            ' フォント強調を適用（全体最大値基準）
            Dim 強調キー後工程項目 As Collection
            Set 強調キー後工程項目 = Get強調対象キー_全体判定(流出Tbl, i, colMap, 後工程全体最大値)
            Call フォント強調適用(ws, Array(targetRow), 強調キー後工程項目, colMap)
            targetRow = targetRow + 1
        Else
            Debug.Print "→ スキップ"
        End If
    Next i

    ' ============================================
    ' 6. 工程内不良処理（新規機能）
    ' ============================================

    ' 6-1. 工程内不良の余分な行を削除（16行目から）
    Application.StatusBar = "工程内不良の行を調整中..."
    Dim deleteCountKouteiNai As Long
    deleteCountKouteiNai = 20 - 不良項目数NW

    If deleteCountKouteiNai > 0 Then
        ws.Rows("16:" & (15 + deleteCountKouteiNai)).Delete Shift:=xlUp
    End If

    ' 6-2. 工程内不良用の列マッピング設定

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

    ' 全体不良数用（8行目）- ショット数と同じ列マッピング
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

    ' リコート用（12行目）- G列開始、補給品あり
    Dim colMapリコート As Object
    Set colMapリコート = CreateObject("Scripting.Dictionary")
    colMapリコート("合計") = "G"
    colMapリコート("58050FrLH") = "L"
    colMapリコート("58050FrRH") = "O"
    colMapリコート("58050RrLH") = "R"
    colMapリコート("58050RrRH") = "U"
    colMapリコート("28050FrLH") = "X"
    colMapリコート("28050FrRH") = "AA"
    colMapリコート("28050RrLH") = "AD"
    colMapリコート("28050RrRH") = "AG"
    colMapリコート("補給品") = "AJ"

    ' 廃棄用（13行目）- G列開始、補給品あり
    Dim colMap廃棄工程内 As Object
    Set colMap廃棄工程内 = CreateObject("Scripting.Dictionary")
    colMap廃棄工程内("合計") = "G"
    colMap廃棄工程内("58050FrLH") = "L"
    colMap廃棄工程内("58050FrRH") = "O"
    colMap廃棄工程内("58050RrLH") = "R"
    colMap廃棄工程内("58050RrRH") = "U"
    colMap廃棄工程内("28050FrLH") = "X"
    colMap廃棄工程内("28050FrRH") = "AA"
    colMap廃棄工程内("28050RrLH") = "AD"
    colMap廃棄工程内("28050RrRH") = "AG"
    colMap廃棄工程内("補給品") = "AJ"

    ' 不良項目用（15行目以降）- G列開始、補給品あり
    Dim colMap工程内 As Object
    Set colMap工程内 = CreateObject("Scripting.Dictionary")
    colMap工程内("合計") = "G"
    colMap工程内("58050FrLH") = "L"
    colMap工程内("58050FrRH") = "O"
    colMap工程内("58050RrLH") = "R"
    colMap工程内("58050RrRH") = "U"
    colMap工程内("28050FrLH") = "X"
    colMap工程内("28050FrRH") = "AA"
    colMap工程内("28050RrLH") = "AD"
    colMap工程内("28050RrRH") = "AG"
    colMap工程内("補給品") = "AJ"

    ' 6-3. 工程内不良データの転記処理
    Application.StatusBar = "工程内不良データを転記中..."

    ' 全項目行の全体最大値を事前計算
    Dim 工程内全体最大値 As Double
    工程内全体最大値 = Get全体最大値(工程内Tbl, colMap工程内, "ショット数,不良数,リコート,廃棄,その他")
    Debug.Print "工程内全体最大値: " & 工程内全体最大値

    Dim targetRow工程内 As Long
    targetRow工程内 = 15  ' 不良項目の開始行（削除後）

    For i = 1 To 工程内Tbl.ListRows.Count
        itemNW = 工程内Tbl.ListColumns("項目").DataBodyRange(i, 1).Value
        Debug.Print "工程内行" & i & ": [" & itemNW & "]"

        If Trim(itemNW) = "ショット数" Then
            Debug.Print "→ ショット数を7行目に転記"
            ' 7行目に転記
            Call 転記行データ(ws, 7, 工程内Tbl, i, colMapショット数)
        ElseIf Trim(itemNW) = "不良数" Then
            Debug.Print "→ 不良数を11行目・14行目に転記"
            ' 11行目に転記
            Call 転記行データ(ws, 11, 工程内Tbl, i, colMap不良数)
            ' 14行目に転記
            Call 転記行データ(ws, 14, 工程内Tbl, i, colMap不良数)
            ' 工程内不良数を保存（全体不良数計算用）
            Dim keyCol2 As Variant
            For Each keyCol2 In colMap不良数.Keys
                工程内不良数(CStr(keyCol2)) = 工程内Tbl.ListColumns(CStr(keyCol2)).DataBodyRange(i, 1).Value
            Next keyCol2
            ' フォント強調を適用
            Dim 強調キー工程内不良数 As Collection
            Set 強調キー工程内不良数 = Get強調対象キー(工程内Tbl, i, colMap不良数)
            Call フォント強調適用(ws, Array(11, 14), 強調キー工程内不良数, colMap不良数)
        ElseIf Trim(itemNW) = "リコート" Then
            Debug.Print "→ リコートを12行目に転記"
            ' 12行目に転記
            Call 転記行データ(ws, 12, 工程内Tbl, i, colMapリコート)
        ElseIf Trim(itemNW) = "廃棄" Then
            Debug.Print "→ 廃棄を13行目に転記"
            ' 13行目に転記
            Call 転記行データ(ws, 13, 工程内Tbl, i, colMap廃棄工程内)
        Else
            ' 不良項目を15行目以降に転記
            Debug.Print "→ 不良項目を" & targetRow工程内 & "行目に転記"
            ' 項目名を転記
            ws.Range("E" & targetRow工程内).Value = itemNW
            ' データを転記
            Call 転記行データ(ws, targetRow工程内, 工程内Tbl, i, colMap工程内)
            ' フォント強調を適用（全体最大値基準）
            Dim 強調キー工程内項目 As Collection
            Set 強調キー工程内項目 = Get強調対象キー_全体判定(工程内Tbl, i, colMap工程内, 工程内全体最大値)
            Call フォント強調適用(ws, Array(targetRow工程内), 強調キー工程内項目, colMap工程内)
            targetRow工程内 = targetRow工程内 + 1
        End If
    Next i

    ' 全体不良数を計算して8行目に転記
    Application.StatusBar = "全体不良数を計算中..."
    Dim keyCol3 As Variant
    Dim 工程内値 As Variant, 後工程値 As Variant
    Dim 合計値 As Double

    For Each keyCol3 In colMap全体不良数.Keys
        ' 工程内不良数と後工程不良数を取得
        工程内値 = 0
        後工程値 = 0

        If 工程内不良数.Exists(CStr(keyCol3)) Then
            工程内値 = 工程内不良数(CStr(keyCol3))
        End If

        If 後工程不良数.Exists(CStr(keyCol3)) Then
            後工程値 = 後工程不良数(CStr(keyCol3))
        End If

        ' 合計を計算
        合計値 = CDbl(工程内値) + CDbl(後工程値)

        ' 8行目に転記
        ws.Range(colMap全体不良数(keyCol3) & "8").Value = 合計値
        Debug.Print "  全体不良数: " & keyCol3 & " = " & 工程内値 & " + " & 後工程値 & " = " & 合計値
    Next keyCol3

    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical
End Sub

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

Private Function Get全体最大値(srcTbl As ListObject, colMap As Object, _
                            除外項目 As String) As Double
    ' 全項目行の全セルから最大値を取得
    ' 除外項目：カンマ区切り（例: "ショット数,不良数,リコート,廃棄"）

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
    Dim count As Long
    count = 0

    Do While 候補.count > 0 And count < 3  ' 上限3つ
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
            count = count + 1
        Else
            Exit Do
        End If
    Loop

End Function

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
    Dim count As Long
    count = 0

    Do While 候補.count > 0 And count < 3  ' 上限3つ
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
            count = count + 1
        Else
            Exit Do
        End If
    Loop

End Function

Private Sub フォント強調適用(ws As Worksheet, targetRows As Variant, _
                          強調キー As Collection, colMap As Object)
    ' 指定された行とキーのセルのフォントサイズを+2ポイント、太字

    If 強調キー.count = 0 Then
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
