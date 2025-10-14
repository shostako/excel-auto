# ピボットテーブルのClearAllFiltersが効かない問題と段階的解決策

## 問題の概要
VBAでピボットテーブルのフィールドフィルターを解除しようとした際、`ClearAllFilters`メソッドが効かない場合がある。特に列フィールド（Fr/Rr、RH/LH等）で発生しやすい。

## 発生日時
2025-01-26

## 症状
- `PivotField.ClearAllFilters`を実行してもフィルターが解除されない
- エラーは発生しないが、フィルター状態が変わらない
- 特に列フィールド（xlColumnField）で発生

## 初期の失敗アプローチ
```vba
' 失敗例：単純な個別選択
With pt.PivotFields("Fr/Rr")
    For Each pi In .PivotItems
        If pi.Name = "Fr" Or pi.Name = "Rr" Then
            pi.Visible = True
        Else
            pi.Visible = False
        End If
    Next pi
End With
```

## 正しい解決策：段階的フォールバック戦略

### 1. ヘルパー関数：柔軟なフィールド検索
```vba
Function GetPivotField(pt As PivotTable, fieldName As String) As PivotField
    Dim pf As PivotField
    Dim n1 As String, n2 As String

    ' 比較用に正規化（小文字・スペース除去）
    n1 = LCase(Replace(fieldName, " ", ""))

    For Each pf In pt.PivotFields
        n2 = LCase(Replace(CStr(pf.Name), " ", ""))
        If n1 = n2 Then
            Set GetPivotField = pf
            Exit Function
        End If
    Next pf

    Set GetPivotField = Nothing
End Function
```

### 2. 段階的フォールバック実装
```vba
Sub EnsurePivotFieldSelectAll(pt As PivotTable, fieldName As String)
    Dim pf As PivotField
    Dim pi As PivotItem
    Dim prevManual As Boolean

    On Error GoTo EH

    Set pf = GetPivotField(pt, fieldName)
    If pf Is Nothing Then Exit Sub

    ' 手段1: 標準的なClearAllFilters
    On Error Resume Next
    pf.ClearAllFilters
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    On Error GoTo EH

    ' 手段2: PivotCache更新後の再試行
    On Error Resume Next
    If Not pt.PivotCache Is Nothing Then
        pt.PivotCache.Refresh
    End If
    pt.RefreshTable
    On Error GoTo EH

    On Error Resume Next
    pf.ClearAllFilters
    If Err.Number = 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    On Error GoTo EH

    ' 手段3: ページフィールドの特殊処理
    If pf.Orientation = xlPageField Then
        On Error Resume Next
        pf.EnableMultiplePageItems = True
        pf.CurrentPage = "(All)"
        If Err.Number = 0 Then
            Err.Clear
            On Error GoTo 0
            Exit Sub
        End If
        Err.Clear
        On Error GoTo EH
    End If

    ' 手段4: 最終手段 - 全PivotItemを個別に表示化
    prevManual = pt.ManualUpdate
    On Error Resume Next
    pt.ManualUpdate = True

    For Each pi In pf.PivotItems
        On Error Resume Next
        pi.Visible = True
        If Err.Number <> 0 Then
            Err.Clear
        End If
    Next pi

    pt.ManualUpdate = prevManual
    On Error GoTo EH

    Exit Sub

EH:
    Debug.Print "EnsurePivotFieldSelectAll エラー: " & Err.Number & " - " & Err.Description
    On Error GoTo 0
End Sub
```

## 重要な学習ポイント

### 1. 段階的フォールバック戦略の重要性
- 一つの方法に固執せず、複数の代替手段を用意する
- 簡単な方法から複雑な方法へ順番に試す
- 各手段の成功/失敗を適切に判定する

### 2. フィールド名の柔軟な検索
- 大文字小文字の違いを吸収
- スペースの有無を吸収
- 実運用でのロバスト性向上

### 3. PivotCacheの重要性
- フィルター問題の多くはPivotCacheのリフレッシュで解決
- RefreshTableと組み合わせることで効果的

### 4. ManualUpdateの活用
- 大量のPivotItem操作時は必須
- パフォーマンスとエラー回避の両立

## 教訓
- **表面的な対処療法ではなく、根本原因を考える**
- **複数の解決策を段階的に試すアーキテクチャを設計する**
- **実運用での様々なケースを想定したロバストな実装を心がける**

## 関連ファイル
- `/home/shostako/ClaudeCode/excel-auto/src/mグラフ表示_ゾーン別改訂.bas`