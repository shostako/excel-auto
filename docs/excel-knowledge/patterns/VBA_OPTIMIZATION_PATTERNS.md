# VBA最適化パターン集

症状から原因を特定し、適切な対処を行うためのパターン集。

## パターン1: 画面がちらつく・もたもた動く

### 症状
- セルが一つずつ更新される様子が見える
- シート間を行き来する画面が表示される
- 「アニメーション」のような動き

### 原因（優先順位順）
1. **Activateメソッドの使用**（90%これ）
2. ScreenUpdatingが有効のまま
3. Selectメソッドの多用

### 診断コード
```vba
' 問題のあるコード例
ws.Activate  ' ← 真犯人
Range("A1").Select  ' ← これも悪い
Selection.Value = "データ"  ' ← 最悪
```

### 解決策
```vba
' 正しいコード
' 1. 画面更新を停止（最優先）
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

' 2. Activateを使わない
' ws.Activate ← 削除！

' 3. オブジェクト参照で直接操作
ws.Range("A1").Value = "データ"  ' Selectなし

' 4. 最後に設定を戻す
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
```

## パターン2: 処理が異常に遅い

### 症状
- 単純な処理なのに数分かかる
- Excelが応答なしになる
- プログレスバーが進まない

### 原因チェックリスト
1. ScreenUpdating確認（まずこれ）
2. 計算モード確認
3. セル単位の処理をしていないか
4. 無駄なループがないか

### 診断と解決
```vba
' 遅いコード
For i = 1 To 10000
    Cells(i, 1).Value = i  ' セル単位は遅い
Next i

' 速いコード - 配列使用
Dim arr(1 To 10000, 1 To 1) As Variant
For i = 1 To 10000
    arr(i, 1) = i
Next i
Range("A1:A10000").Value = arr  ' 一括書き込み
```

## パターン3: CommandButton一括実行での問題

### 症状
- 個別実行は問題ないが、一括実行でちらつく
- 途中で画面更新が再開される

### 原因
個別マクロが勝手に設定を変更している

### 解決策
```vba
' CommandButtonのコード
Sub 一括実行()
    Application.ScreenUpdating = False
    
    Call マクロ1
    Call マクロ2
    Call マクロ3
    
    Application.ScreenUpdating = True
End Sub

' 個別マクロ（修正後）
Sub マクロ1()
    ' ScreenUpdating設定はしない！
    ' 処理のみ記述
End Sub
```

## パターン4: メモリ不足エラー

### 症状
- 「メモリが不足しています」エラー
- 大量データで落ちる

### 原因と対策
```vba
' 問題：Range全体をコピー
Range("A:A").Copy Range("B:B")  ' 100万行コピー

' 解決：必要な範囲のみ
Dim lastRow As Long
lastRow = Cells(Rows.Count, "A").End(xlUp).Row
Range("A1:A" & lastRow).Copy Range("B1")
```

## 黄金律

### 最適化の優先順位
1. **画面制御**（ScreenUpdating = False）
2. **計算制御**（Calculation = Manual）
3. **Activate/Select排除**
4. **配列処理**
5. **アルゴリズム改善**

### 症状診断フロー
```
もたもた・ちらつき
    ↓
ScreenUpdating確認
    ↓
Activate探す
    ↓
それでもダメならアルゴリズム
```

## Mondayの格言

> 「賢い解決策より、正しい解決策を選べ」
> 
> 「Dictionaryより先にScreenUpdating」
> 
> 「Activateを見たら即削除」

---

*注：このパターン集は、実際の失敗から生まれた。理論より実践、完璧より実用性を重視する。*