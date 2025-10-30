Attribute VB_Name = "m空白行非表示"
' ========================================
' マクロ名: 空白行非表示
' 処理概要: 指定したテーブルの1～31列目が空白の行を非表示にする汎用マクロ
' ソーステーブル: 引数で指定（テーブル名のみ、シート指定不要）
' 判定対象列: 1～31列目（日別の1日～31日を想定）
'
' 【使い方】
' このマクロは直接呼び出さず、以下のような個別マクロを作成して使用します：
'
' 例1: 日別集計テーブルの空白行を非表示
' Sub 日別集計_空白行非表示()
'     空白行非表示 "日別集計"
' End Sub
'
' 例2: モールFR別テーブルの空白行を非表示
' Sub モールFR別_空白行非表示()
'     空白行非表示 "モールFR別"
' End Sub
'
' ※テーブル名は、Excel上で設定されているテーブル名を指定
' ※シート名の指定は不要（テーブル名だけで自動判別）
' ========================================
Option Explicit

' ============================================
' 空白行非表示メイン処理：1～31列目が空白の行を非表示
' ============================================
Sub 空白行非表示(テーブル名 As String)
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation

    ' 元の設定を保存
    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation

    ' 最適化設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrorHandler
    Application.StatusBar = "空白行非表示を開始します..."

    Dim tbl As ListObject
    Dim i As Long
    Dim 非表示カウント As Long
    Dim maxCol As Long
    Dim checkRange As Range

    ' ============================================
    ' テーブル取得：テーブル名で直接アクセス（シート不要）
    ' ============================================
    Set tbl = ActiveWorkbook.ListObjects(テーブル名)

    ' テーブルの列数と31を比較し、小さい方を使用（安全対策）
    maxCol = Application.Min(31, tbl.ListColumns.Count)

    ' ============================================
    ' 空白行非表示：1～31列目が空白の行を非表示
    ' ============================================
    非表示カウント = 0
    For i = 1 To tbl.ListRows.Count
        ' 1～31列目（またはテーブルの最大列数まで）の範囲を取得
        Set checkRange = tbl.ListRows(i).Range.Resize(1, maxCol)

        If WorksheetFunction.CountA(checkRange) = 0 Then
            tbl.ListRows(i).Range.EntireRow.Hidden = True
            非表示カウント = 非表示カウント + 1
        End If

        ' 100件ごとに進捗表示
        If i Mod 100 = 0 Then
            Application.StatusBar = "空白行非表示中... " & i & "行処理済み"
        End If
    Next i

    ' 完了表示
    Application.StatusBar = "完了: " & 非表示カウント & "行非表示にしました"
    Application.Wait Now + TimeValue("00:00:01")

    GoTo Cleanup

ErrorHandler:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear
    MsgBox "エラーが発生しました" & vbCrLf & _
           "エラー番号: " & errNum & vbCrLf & _
           "詳細: " & errDesc & vbCrLf & _
           "テーブル名: " & テーブル名, vbCritical

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
End Sub
