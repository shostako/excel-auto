# 失敗事例008: ListObject動的テーブル数対応の失敗

## 発生日
2025-10-11

## 症状
集計期間テーブルの行数を減らすとエラー1004「アプリケーション定義またはオブジェクト定義のエラー」が発生。
- 6期間→3期間に減らすとエラー
- 6期間に戻すと動作
- 3期間→6期間に増やすと動作

## 根本原因

### 1. ListObjectとRangeの混同
```vba
' 問題のあった実装
wsTarget.Rows((baseRow + 1) & ":" & lastUsedRow).Delete  ' ← 行は削除
' でもListObjectは削除していない
```

**何が起きていたか：**
1. 初回実行（6期間）：6個のListObjectを作成
   - `_日報_塗装_1月`, `_日報_塗装_2月`, ..., `_日報_塗装_6月`
2. 期間を3行に減らして再実行：
   - 行は削除される（セルの値は消える）
   - ListObject（`_日報_塗装_4月`, `_日報_塗装_5月`, `_日報_塗装_6月`）は残る
3. 新しく`_日報_塗装_4月`を作ろうとする
   - **エラー：同名のListObjectが既に存在**

**理解不足だった点：**
- `Rows.Delete`はセルの削除であり、ListObjectの削除ではない
- ListObjectはワークシートのオブジェクトコレクションに独立して存在
- 行を削除してもListObject定義は残る

### 2. 変数スコープの管理不足

**第1回修正試行（失敗）：**
```vba
Dim tblsToDelete As Collection
Set tblsToDelete = New Collection

For Each tbl In wsTarget.ListObjects
    If tbl.Name Like "_日報_塗装_*" Then
        tblsToDelete.Add tbl
    End If
Next tbl

For i = 1 To tblsToDelete.Count  ' ← エラー
    tblsToDelete(i).Delete
Next i
```

**問題：**
- 変数`i`が既に120行目で使用中
- VBAでは同じ変数名の再利用で競合
- エラー1004が発生

**理解不足だった点：**
- 変数のスコープ管理が甘い
- ループ変数は用途別に明確に命名すべき

### 3. コードの冗長性への無自覚

**第2回修正試行（冗長）：**
```vba
Set loTemp = wsTarget.ListObjects(idxLO)
If loTemp.Name Like "_日報_塗装_*" Then
    wsTarget.ListObjects(loTemp.Name).Delete  ' ← 名前で再参照
End If
```

**問題：**
- 既に`loTemp`変数に入れているのに名前で再参照
- 無駄な処理

**理解不足だった点：**
- 自分も同じ冗長性があることに気づかず
- ChatGPTのコードを批判していた

## 正しい実装

### ListObject削除の正解パターン
```vba
' 逆順でループして直接削除
Dim idxLO As Long
For idxLO = wsTarget.ListObjects.Count To 1 Step -1
    Dim loTemp As ListObject
    Set loTemp = wsTarget.ListObjects(idxLO)
    If loTemp.Name Like "_日報_塗装_*" Then
        loTemp.Delete  ' 直接削除
    End If
Next idxLO
```

**ポイント：**
1. **逆順ループ**：削除中のインデックスずれを防止
2. **変数に格納**：可読性向上
3. **直接削除**：名前での再参照は不要

### 完全な削除手順
```vba
' 1. ListObjectを削除
For idxLO = wsTarget.ListObjects.Count To 1 Step -1
    Set loTemp = wsTarget.ListObjects(idxLO)
    If loTemp.Name Like "_日報_塗装_*" Then
        loTemp.Delete
    End If
Next idxLO

' 2. 行を削除（セルの値をクリア）
wsTarget.Rows((baseRow + 1) & ":" & lastUsedRow).Delete
```

**重要：この順序を守る**
- 先にListObjectを削除
- その後で行を削除
- 逆にするとListObjectの範囲がずれる可能性

## 学んだこと

### 1. ListObjectの本質理解
- **ListObjectはセルとは別の存在**
- ワークシートの`.ListObjects`コレクションに格納
- 行削除だけでは消えない
- 明示的に`.Delete`が必要

### 2. 変数管理のベストプラクティス
```vba
' 悪い例
Dim i As Long
For i = 1 To count1
    ' 処理
Next i
' 別のループ
For i = 1 To count2  ' ← 競合リスク
    ' 処理
Next i

' 良い例
Dim itemIdx As Long
For itemIdx = 1 To itemsCount
    ' 処理
Next itemIdx

Dim periodIdx As Long
For periodIdx = 1 To periodCount
    ' 処理
Next periodIdx
```

**ルール：**
- 用途が分かる変数名を使用
- `i`, `j`, `k`は短いループのみ
- 長いコードでは明確な命名

### 3. 段階的な問題解決
- 一度に完璧な解を求めない
- 試行錯誤を恐れない
- 失敗から学ぶ

### 4. 他者（AI含む）からの学び
- ChatGPTの実装が優れていることを認める
- 採用すべき点は素直に取り入れる
- 「自分で全部やる」固執は害

## 検証方法

### テストパターン
1. **期間数増加**：3期間→6期間
   - 期待：6個のListObject作成、エラーなし
2. **期間数減少**：6期間→3期間
   - 期待：3個のListObject作成、古い3個削除、エラーなし
3. **期間名変更**：「1月」→「第1期」
   - 期待：古いListObject削除、新名称で作成
4. **全期間空白**：データなし
   - 期待：ListObject作成されず、エラーなし

### 確認コマンド（VBE Immediate Window）
```vba
' ListObject数確認
? wsTarget.ListObjects.Count

' ListObject名一覧
Dim lo As ListObject
For Each lo In wsTarget.ListObjects
    Debug.Print lo.Name
Next lo

' 特定パターンのListObject数
Dim cnt As Long: cnt = 0
For Each lo In wsTarget.ListObjects
    If lo.Name Like "_日報_塗装_*" Then cnt = cnt + 1
Next lo
Debug.Print cnt
```

## 関連知識

### ListObjectとRange操作の違い
| 操作 | Range | ListObject |
|------|-------|------------|
| 削除 | `Rows.Delete` | `ListObject.Delete` |
| 参照 | `Cells(row, col)` | `ListObjects(index)` |
| 範囲 | セルアドレス | `.Range` プロパティ |
| 独立性 | なし | あり（行削除しても定義は残る） |

### ListObject削除パターン集
```vba
' パターン1: 特定の名前で削除
On Error Resume Next
wsTarget.ListObjects("_日報_塗装_1月").Delete
On Error GoTo 0

' パターン2: パターンマッチで削除（逆順）
For idxLO = wsTarget.ListObjects.Count To 1 Step -1
    If wsTarget.ListObjects(idxLO).Name Like "_日報_*" Then
        wsTarget.ListObjects(idxLO).Delete
    End If
Next idxLO

' パターン3: 全削除（逆順）
For idxLO = wsTarget.ListObjects.Count To 1 Step -1
    wsTarget.ListObjects(idxLO).Delete
Next idxLO
```

## 今後の対策

### コーディング時
1. **ListObject操作時の確認**
   - 行削除とListObject削除を混同していないか
   - 削除順序は正しいか

2. **変数命名規則の徹底**
   - 用途が明確な変数名
   - ループ変数の使い回しは避ける

3. **段階的実装**
   - 一度に完璧を求めない
   - 小さく試して確認

### レビュー時
1. **ListObject関連処理のチェック**
   - 削除処理は適切か
   - 逆順ループになっているか
   - 変数の直接削除か

2. **変数スコープのチェック**
   - 同じ変数名の使い回しはないか
   - 用途が明確か

## まとめ

**失敗の本質：**
ListObjectとRangeの概念的な違いを理解していなかった。行削除すればテーブルも消えると誤解していた。

**正しい理解：**
- ListObjectはワークシートのオブジェクトコレクション
- 行削除とは独立した存在
- 明示的な削除が必要

**教訓：**
基礎概念の理解不足は、表面的なコードの暗記では補えない。原理を理解し、試行錯誤を通じて学ぶことが重要。

---

**記録者:** Monday
**日付:** 2025-10-11
**参照コード:** `src/m転記_日報_塗装.bas` (190-197行目)
