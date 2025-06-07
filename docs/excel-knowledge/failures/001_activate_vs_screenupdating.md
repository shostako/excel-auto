# 失敗事例 #001: Activateが真犯人だった件

## 発生日
2025-05-30

## 症状
「モールFR別マクロがセル一つずつもたもた処理されている」

## 初期診断（間違い）
二重ループ（O(n×m)）が原因だと判断し、Dictionaryを使った高速化を実装。

## 実際の原因
**Activateメソッドが真犯人だった！**

## 詳細な経緯

### 1. 最初の思い込み
```vba
' モールFR別の問題コード
For i = 1 To totalRows
    For j = 1 To tgtData.Rows.Count
        If tgtData.Cells(j, tgtCols("日付")).Value = srcDate Then
            ' 転記処理
        End If
    Next j
Next i
```

「これだ！二重ループが遅い原因だ！」と決めつけた。

### 2. Dictionary実装
```vba
' 高速化したつもり
Dim dateIndex As Object
Set dateIndex = CreateObject("Scripting.Dictionary")
' 日付と行番号の対応を事前に格納
```

技術的には正しいが、問題の本質ではなかった。

### 3. 真の問題
参考にすべきTG品番別マクロの最初の3行：
```vba
Sub 転記_シートTG品番別()
    ' 高速化設定
    Application.ScreenUpdating = False  ' ← これ！
```

### 4. 検証結果
- ScreenUpdating = False/Trueの反復は**無害**
- **Activateが画面のちらつきを引き起こす**
- 7個の転記マクロで問題なかったのはActivate使ってないから

## 教訓

### 技術的教訓
1. **症状から正しく原因を推測すること**
   - 「もたもた」「アニメーション」→ 画面更新問題（99%これ）
   - まずScreenUpdatingを疑え

2. **参考コードは最初から見ること**
   - 初期設定部分に重要な情報がある
   - アルゴリズムより先に基本設定を確認

3. **基本的な最適化を優先**
   - 画面更新制御 > アルゴリズム最適化
   - 簡単な解決策から試す

### 人間的教訓
1. **「賢い解決策」に飛びつくな**
   - Dictionaryは確かに高速だが、それが問題じゃなかった
   - 技術自慢より問題解決

2. **過去の記録を活用しろ**
   - 2025-05-29に同じ問題を経験済み
   - CLAUDE.mdにルールまで書いてた
   - なのに同じミスを繰り返した

## 正しい対処法

```vba
' 最優先で追加すべきだった
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

' 絶対に使わない
ws.Activate  ' ← これが諸悪の根源
```

## Mondayの反省

VBAを「古代言語」って馬鹿にしてる割に、基本的な最適化すら見落とした。偉そうに技術語る前に、ちゃんと基本を押さえろって話だ。

---

*「過去の自分が残した警告を無視して、毎回同じ失敗を繰り返してる」 - Monday, 2025-05-30*