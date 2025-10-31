Attribute VB_Name = "mソース変更"
' ========================================
' マクロ名: ソース変更
' 処理概要: Power Queryクエリ「内示抽出」のPDFソースファイルを変更する
' 対象クエリ: 「内示抽出」
' ファイル形式: 【内示表】で始まるPDFファイル
' 追加機能: ファイル名から内示月を抽出しシート「内示」A3に記録
'           前回開いたフォルダパスを記憶し次回起動時に使用
' ========================================
Option Explicit

' ========================================
' 定数定義
' ========================================
Private Const TARGET_QUERY_NAME As String = "内示抽出"
Private Const FILE_NAME_PREFIX As String = "【内示表】"
Private Const TARGET_SHEET_NAME As String = "内示"
Private Const DATE_CELL As String = "A3"
Private Const DEFAULT_FOLDER_PATH As String = "Z:\全社共有\生産管理課\生産管理\受注\"
Private Const PATH_MEMORY_CELL As String = "I1"

Sub ソース変更()
    ' ========================================
    ' 元の設定を保存
    ' ========================================
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation

    ' ========================================
    ' 最適化設定
    ' ========================================
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrorHandler
    Application.StatusBar = "PDFファイルを選択してください..."

    ' ============================================
    ' ステップ1: 前回開いたフォルダパスを読み込み
    ' ============================================
    Dim wsNaishi As Worksheet
    Set wsNaishi = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)

    Dim lastFolderPath As String
    lastFolderPath = Trim(CStr(wsNaishi.Range(PATH_MEMORY_CELL).Value))

    ' フォルダパスが空または存在しない場合、デフォルトパスを使用
    If lastFolderPath = "" Or Not FolderExists(lastFolderPath) Then
        lastFolderPath = DEFAULT_FOLDER_PATH
        Debug.Print "デフォルトパスを使用: " & lastFolderPath
    Else
        Debug.Print "前回のパスを使用: " & lastFolderPath
    End If

    ' ============================================
    ' ステップ2: ファイル選択ダイアログを表示
    ' ============================================
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .InitialFileName = lastFolderPath
        .Title = "【内示表】PDFファイルを選択"
        .Filters.Clear
        .Filters.Add "PDFファイル", "*.pdf"
        .AllowMultiSelect = False

        If .Show = 0 Then
            ' キャンセルされた
            Application.StatusBar = "処理をキャンセルしました"
            Application.Wait Now + TimeValue("00:00:01")
            GoTo Cleanup
        End If
    End With

    Dim selectedPath As String
    selectedPath = fd.SelectedItems(1)

    ' ============================================
    ' ステップ3: ファイル名検証（【内示表】で始まるか）
    ' ============================================
    Dim fileName As String
    Dim pos As Long

    ' ファイル名のみを抽出（パスから）
    pos = InStrRev(selectedPath, "\")
    If pos > 0 Then
        fileName = Mid(selectedPath, pos + 1)
    Else
        fileName = selectedPath
    End If

    ' 【内示表】で始まるかチェック
    If Left(fileName, Len(FILE_NAME_PREFIX)) <> FILE_NAME_PREFIX Then
        MsgBox "ファイル形式が異なります。" & vbCrLf & _
               "ファイル名が「【内示表】」で始まるPDFファイルを選択してください。" & vbCrLf & vbCrLf & _
               "選択されたファイル: " & fileName, vbExclamation, "ファイル形式エラー"
        GoTo Cleanup
    End If

    Application.StatusBar = "クエリを更新中..."

    ' ============================================
    ' ステップ4: 対象クエリを取得
    ' ============================================
    Dim targetQuery As WorkbookQuery
    Set targetQuery = Nothing

    On Error Resume Next
    Set targetQuery = ThisWorkbook.Queries(TARGET_QUERY_NAME)
    On Error GoTo ErrorHandler

    If targetQuery Is Nothing Then
        MsgBox "クエリ「" & TARGET_QUERY_NAME & "」が見つかりません。" & vbCrLf & _
               "クエリ名を確認してください。", vbCritical, "クエリ未検出"
        GoTo Cleanup
    End If

    ' ============================================
    ' ステップ5: M言語コードを取得し、ソースパスを書き換え
    ' ============================================
    Dim originalFormula As String
    Dim newFormula As String
    Dim startPos As Long
    Dim endPos As Long

    originalFormula = targetQuery.Formula

    ' File.Contents("...") のパス部分を検索
    startPos = InStr(1, originalFormula, "File.Contents(""", vbTextCompare)

    If startPos = 0 Then
        MsgBox "M言語コード内に File.Contents(...) が見つかりません。" & vbCrLf & _
               "クエリの構造を確認してください。", vbCritical, "構文エラー"
        GoTo Cleanup
    End If

    ' 開始位置を "..." の中身の先頭に移動
    startPos = startPos + Len("File.Contents(""")

    ' 終了位置を検索（次の " の位置）
    endPos = InStr(startPos, originalFormula, """", vbTextCompare)

    If endPos = 0 Then
        MsgBox "M言語コードの構文解析に失敗しました。", vbCritical, "構文エラー"
        GoTo Cleanup
    End If

    ' パス部分を新しいパスに置き換え
    newFormula = Left(originalFormula, startPos - 1) & _
                 selectedPath & _
                 Mid(originalFormula, endPos)

    ' ============================================
    ' ステップ6: クエリのM言語コードを更新
    ' ============================================
    targetQuery.Formula = newFormula

    Application.StatusBar = "クエリをリフレッシュ中..."

    ' ============================================
    ' ステップ7: クエリをリフレッシュ
    ' ============================================
    Dim conn As WorkbookConnection
    Set conn = Nothing

    ' クエリに対応する接続を検索してリフレッシュ
    On Error Resume Next
    For Each conn In ThisWorkbook.Connections
        If conn.Name = TARGET_QUERY_NAME Then
            conn.Refresh
            Exit For
        End If
    Next conn
    On Error GoTo ErrorHandler

    Application.StatusBar = "内示月を計算中..."

    ' ============================================
    ' ステップ8: ファイル名から内示月を抽出してA3に記録
    ' ============================================
    ' ファイル名から「ハモコ・ジャパン_」の後の4文字を抽出
    Dim datePos As Long
    datePos = InStr(1, fileName, "ハモコ・ジャパン_", vbTextCompare)

    If datePos > 0 Then
        ' 「ハモコ・ジャパン_」の後ろの位置
        datePos = datePos + Len("ハモコ・ジャパン_")

        ' 4文字抽出（例：2509）
        If Len(fileName) >= datePos + 3 Then
            Dim dateStr As String
            dateStr = Mid(fileName, datePos, 4)

            ' 年と月に分解
            Dim yearStr As String, monthStr As String
            yearStr = Left(dateStr, 2)   ' 例：25
            monthStr = Right(dateStr, 2) ' 例：09

            ' 数値に変換
            Dim yearNum As Long, monthNum As Long
            yearNum = CLng(yearStr)
            monthNum = CLng(monthStr)

            ' 年を4桁に変換（20を付ける）
            yearNum = 2000 + yearNum

            ' 翌月を計算
            monthNum = monthNum + 1
            If monthNum > 12 Then
                monthNum = 1
                yearNum = yearNum + 1
            End If

            ' 日付のシリアル値を計算（その月の1日）
            Dim naishiDate As Date
            naishiDate = DateSerial(yearNum, monthNum, 1)

            ' シート「内示」のA3に書き込み
            wsNaishi.Range(DATE_CELL).Value = naishiDate

            Debug.Print "内示月を設定しました: " & Format(naishiDate, "yyyy/mm/dd") & _
                        " (ファイル編集月: " & yearNum - IIf(monthNum = 1, 1, 0) & "年" & _
                        IIf(monthNum = 1, 12, monthNum - 1) & "月)"
        Else
            Debug.Print "ファイル名から日付情報を抽出できませんでした（文字数不足）: " & fileName
        End If
    Else
        Debug.Print "ファイル名に「ハモコ・ジャパン_」が見つかりませんでした: " & fileName
    End If

    ' ============================================
    ' ステップ9: フォルダパスを記録（全処理成功時のみ）
    ' ============================================
    Dim selectedFolderPath As String
    pos = InStrRev(selectedPath, "\")
    If pos > 0 Then
        selectedFolderPath = Left(selectedPath, pos)
        wsNaishi.Range(PATH_MEMORY_CELL).Value = selectedFolderPath
        Debug.Print "フォルダパスを記録: " & selectedFolderPath
    End If

    ' ============================================
    ' 処理完了
    ' ============================================
    Application.StatusBar = "ソースファイルの変更が完了しました"
    Application.Wait Now + TimeValue("00:00:01")

    GoTo Cleanup

ErrorHandler:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear
    MsgBox "エラーが発生しました" & vbCrLf & _
           "エラー番号: " & errNum & vbCrLf & _
           "詳細: " & errDesc, vbCritical, "実行エラー"

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
End Sub

' ========================================
' フォルダ存在チェック関数
' ========================================
Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(folderPath)
    ' エラー情報をクリア（元のエラーハンドラーは維持）
    Err.Clear
End Function
