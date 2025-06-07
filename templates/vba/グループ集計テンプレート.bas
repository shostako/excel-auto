Attribute VB_Name = "mグループ集計テンプレート"
Sub グループ集計テンプレート()
    '==========================================
    ' グループ集計マクロテンプレート
    ' 用途: テーブルデータをグループ化して集計
    ' 
    ' カスタマイズ箇所:
    ' 1. シート名とテーブル名
    ' 2. グループ化するキー項目
    ' 3. 集計する項目と集計方法
    '==========================================
    
    ' 変数宣言
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim dict As Object
    Dim dataRange As Range
    Dim i As Long
    Dim key As String
    Dim value As Double
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' 高速化設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Dictionary作成（グループ化用）
    Set dict = CreateObject("Scripting.Dictionary")
    
    '==========================================
    ' カスタマイズ箇所1: シートとテーブル
    '==========================================
    Set ws = ThisWorkbook.Worksheets("データシート")  ' ← シート名を変更
    Set tbl = ws.ListObjects("テーブル名")            ' ← テーブル名を変更
    
    ' データ範囲取得
    If tbl.DataBodyRange Is Nothing Then
        MsgBox "テーブルにデータがありません。", vbCritical
        GoTo CleanupAndExit
    End If
    Set dataRange = tbl.DataBodyRange
    
    ' 進捗表示
    Application.StatusBar = "集計処理を開始します..."
    
    '==========================================
    ' カスタマイズ箇所2: グループ化と集計
    '==========================================
    ' データをループしてグループ化
    For i = 1 To dataRange.Rows.Count
        ' グループキーの作成（例: 品番と日付の組み合わせ）
        key = dataRange.Cells(i, 1).Value & "_" & _
              Format(dataRange.Cells(i, 2).Value, "yyyy-mm-dd")  ' ← 列番号を調整
        
        ' 集計値の取得（例: 数量）
        value = Val(dataRange.Cells(i, 3).Value)  ' ← 列番号を調整
        
        ' Dictionaryに追加または加算
        If dict.Exists(key) Then
            dict(key) = dict(key) + value  ' 合計
        Else
            dict(key) = value
        End If
        
        ' 進捗表示（1000件ごと）
        If i Mod 1000 = 0 Then
            Application.StatusBar = "集計処理中... " & i & "/" & dataRange.Rows.Count
        End If
    Next i
    
    '==========================================
    ' カスタマイズ箇所3: 結果出力
    '==========================================
    ' 結果を別シートに出力（例）
    Dim wsResult As Worksheet
    Set wsResult = ThisWorkbook.Worksheets("集計結果")  ' ← シート名を変更
    
    ' ヘッダー設定
    wsResult.Range("A1").Value = "キー項目"
    wsResult.Range("B1").Value = "集計値"
    
    ' 結果出力
    Dim row As Long
    row = 2
    Dim k As Variant
    For Each k In dict.Keys
        wsResult.Cells(row, 1).Value = k
        wsResult.Cells(row, 2).Value = dict(k)
        row = row + 1
    Next k
    
    ' 完了メッセージ（デバッグ用）
    Debug.Print "集計完了: " & dict.Count & "グループ"
    
CleanupAndExit:
    ' 設定を元に戻す
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    ' オブジェクト解放
    Set dict = Nothing
    Exit Sub
    
ErrorHandler:
    ' エラー処理
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    MsgBox "エラーが発生しました: " & Err.Description & vbCrLf & _
           "エラー番号: " & Err.Number, vbCritical, "エラー"
End Sub

'==========================================
' 応用例: 複数項目の集計
'==========================================
Sub グループ集計_複数項目版()
    ' カスタムクラスを使った高度な集計例
    
    ' 集計データを保持するクラス
    ' ※VBAではクラスモジュールを別途作成する必要があります
    
    ' 以下は疑似コード：
    ' Dim dict As Object
    ' Set dict = CreateObject("Scripting.Dictionary")
    ' 
    ' For Each row In データ
    '     key = グループキー作成
    '     
    '     If Not dict.Exists(key) Then
    '         ' 新規作成
    '         Dim item As New 集計アイテム
    '         item.Count = 1
    '         item.Sum = value
    '         item.Max = value
    '         item.Min = value
    '         dict.Add key, item
    '     Else
    '         ' 既存更新
    '         With dict(key)
    '             .Count = .Count + 1
    '             .Sum = .Sum + value
    '             If value > .Max Then .Max = value
    '             If value < .Min Then .Min = value
    '         End With
    '     End If
    ' Next
End Sub