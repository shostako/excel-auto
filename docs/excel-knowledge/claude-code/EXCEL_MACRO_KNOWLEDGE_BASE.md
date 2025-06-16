# Excel マクロナレッジベース v3.0 for Claude Code

このドキュメントは、Claude Codeでのマクロ開発で同じ失敗を繰り返さないための実戦的ナレッジです。
「きれいなコード」より「動くコード」、理論より実践を重視します。

## 1. 出力形式ルール（Claude Code特化）

### 1.1 基本ルール
```
VBAマクロ生成:
  形式: 直接ファイル作成（.basファイル）
  保存先: src/ディレクトリ
  言語: VBA
  コメント: 日本語
  エンコーディング: UTF-8（自動変換でShift-JIS化）
  完了メッセージ: エラー時以外は表示しない
  
M言語生成:
  形式: 直接ファイル作成（.pq または .txt）
  保存先: src/ディレクトリ
  言語: Power Query M
  コメント: 日本語
  
関数・数式生成:
  形式: 通常メッセージ内に記載
  説明: 日本語で使用方法を説明
```

### 1.2 重要な制約事項

#### 参考マクロの読み込み（超重要！）
```bash
# 読む前に必ず実行
iconv -f SHIFT-JIS -t UTF-8 "参考マクロ/ファイル名.bas" | head -100
```

**理由**: 
- ユーザーがExcelからエクスポートしたファイルは**Shift-JIS**
- そのまま読むと文字化けして**列名を誤認識**する
- 過去の事故例：文字化けした列名で修正→本番で動作しない

**禁止事項**:
- 文字化けしたまま読み進めること
- 文字化けした内容を基に修正すること

## 2. 致命的な失敗パターンと対策

### 2.1 画面ちらつき問題の真犯人（最重要）

#### 症状
- 「もたもた」「アニメーション」のような動き
- セルが一つずつ更新される様子が見える

#### 間違った診断（みんなが陥る罠）
「二重ループ（O(n×m)）が原因だ！」→ Dictionary実装

#### 真の原因
**Activateメソッドが諸悪の根源！**

```vba
' これが真犯人
ws.Activate  ' ← 画面がパタパタする元凶
Range("A1").Select  ' ← これも悪い
```

#### 正しい対処法
```vba
' 1. 最優先で追加
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

' 2. Activateを完全排除
' ws.Activate ← 削除！

' 3. オブジェクト参照で直接操作
ws.Range("A1").Value = "データ"  ' Selectなし
```

### 2.2 CommandButton実行時の競合問題

#### 症状
- 個別実行は問題ないが、一括実行でちらつく
- 途中で画面更新が再開される

#### 原因
個別マクロが勝手に設定を変更している

#### 解決策
```vba
' CommandButtonのコード
Sub 一括実行()
    Application.ScreenUpdating = False
    
    Call マクロ1
    Call マクロ2
    Call マクロ3
    
    Application.ScreenUpdating = True
End Sub

' 個別マクロ（修正版）
Sub マクロ1()
    ' 個別マクロの最後で設定を戻さない！
    ' CommandButtonに任せる
    ' 処理のみ記述
End Sub
```

## 3. VBAマクロ基本テンプレート

### 3.1 標準テンプレート（失敗から学んだ版）
```vba
Option Explicit

' モジュール名: m処理名_詳細名
' 処理概要をここに記載

Sub OptimizedMacroTemplate()
    ' 最適化設定の保存
    Dim origScreenUpdating As Boolean
    Dim origCalculation As XlCalculation
    Dim origEnableEvents As Boolean
    Dim origDisplayAlerts As Boolean
    origScreenUpdating = Application.ScreenUpdating
    origCalculation = Application.Calculation
    origEnableEvents = Application.EnableEvents
    origDisplayAlerts = Application.DisplayAlerts
    
    ' 最適化設定（これが最重要）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    ' エラーハンドリング設定
    On Error GoTo ErrorHandler
    
    ' ステータスバー初期化
    Application.StatusBar = "処理を開始します..."
    
    ' =================================
    ' メイン処理をここに記述
    ' 注意：Activateは絶対に使わない！
    ' =================================
    
    ' 処理完了のステータスバー表示
    Application.StatusBar = "処理が完了しました"
    Application.Wait Now + TimeValue("00:00:01")
    
    GoTo Cleanup
    
ErrorHandler:
    ' エラー情報の詳細化
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear
    
    MsgBox "エラーが発生しました" & vbCrLf & _
           "エラー番号: " & errNum & vbCrLf & _
           "詳細: " & errDesc, vbCritical, "エラー"
    
Cleanup:
    ' 設定を確実に復元
    Application.StatusBar = False
    Application.ScreenUpdating = origScreenUpdating
    Application.Calculation = origCalculation
    Application.EnableEvents = origEnableEvents
    Application.DisplayAlerts = origDisplayAlerts
End Sub
```

### 3.2 軽量版テンプレート
```vba
Option Explicit

Sub LightweightTemplate()
    ' 簡単な処理でもScreenUpdatingは必須
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    ' メイン処理
    ' 注意：Activateは使わない
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub
```

## 4. 最適化設定の実戦的使い分け

### 4.1 基本設定（ほぼ必須）
```vba
Application.ScreenUpdating = False  ' 画面ちらつき防止（最重要）
' エラーハンドリング（設定復元のため必須）
```

### 4.2 条件付き設定
```vba
Application.Calculation = xlCalculationManual     ' 計算式多い場合
Application.EnableEvents = False                  ' イベント処理停止
Application.DisplayAlerts = False                 ' 確認ダイアログ抑制
```

#### DisplayAlertsが必要な操作
- テーブル削除（`ListObjects.Delete`）
- 大量データクリア（`Range.Clear`）
- 行列削除操作

## 5. テーブル操作のベストプラクティス

### 5.1 安全性重視のアプローチ

#### 推奨：削除→新規作成
```vba
' 既存テーブル削除（安全な方法）
On Error Resume Next
If Not destSheet.ListObjects("テーブル名") Is Nothing Then
    destSheet.ListObjects("テーブル名").Delete
End If
On Error GoTo ErrorHandler

' 新規テーブル作成
Set destTable = destSheet.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
destTable.Name = "テーブル名"
```

**メリット：**
- 列構造の不一致エラーが発生しない
- 確実に期待通りの構造になる
- 手動変更による列名不一致を回避

#### 非推奨：テーブル再利用
**デメリット：**
- 既存テーブルの列構造が期待と異なる場合にエラー
- 手動変更により列名不一致の可能性
- 複雑な構造検証が必要

## 6. 症状別診断パターン

### 6.1 画面がちらつく・もたもた動く
```
症状: セルが一つずつ更新される様子が見える
   ↓
原因1: Activateメソッドの使用（90%これ）
   ↓
対策: Activate完全削除 + ScreenUpdating = False
```

### 6.2 処理が異常に遅い
```
症状: 単純な処理なのに数分かかる
   ↓
チェック順序:
1. ScreenUpdating確認（まずこれ）
2. 計算モード確認
3. セル単位処理→配列処理
4. 無駄なループ削除
```

### 6.3 最適化の優先順位（黄金律）
1. **画面制御**（ScreenUpdating = False）
2. **Activate/Select排除**
3. **計算制御**（Calculation = Manual）
4. **配列処理**
5. **アルゴリズム改善**

## 7. メッセージ表示ガイドライン

### 7.1 ステータスバー（推奨）
```vba
' 進捗表示（100行ごとに更新）
If i Mod 100 = 0 Then
    Application.StatusBar = "処理中... " & Format(i / totalRows, "0%")
End If

' 処理完了表示
Application.StatusBar = "処理が完了しました"
Application.Wait Now + TimeValue("00:00:01")
Application.StatusBar = False
```

### 7.2 MsgBox（制限付き使用）
**使用場面：**
- エラー発生時の通知のみ

**禁止事項：**
- 正常終了時の「完了しました」メッセージ
- 進捗表示（ステータスバーを使用）

## 8. エラーハンドリング実戦版

### 8.1 基本構造
```vba
On Error GoTo ErrorHandler
' 処理
GoTo Cleanup

ErrorHandler:
    ' エラー詳細の取得と表示
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    Err.Clear
    
    MsgBox "エラーが発生しました" & vbCrLf & _
           "エラー番号: " & errNum & vbCrLf & _
           "詳細: " & errDesc, vbCritical

Cleanup:
    ' 設定を確実に復元（最重要）
    Application.StatusBar = False
    Application.ScreenUpdating = origScreenUpdating
    ' その他の設定復元
End Sub
```

## 9. 文字エンコーディング管理

### 9.1 読み込み時の必須手順
```bash
# 参考マクロファイル読み込み前に必ず実行
iconv -f SHIFT-JIS -t UTF-8 "参考マクロ/ファイル名.bas" | head -100

# 文字化け確認用
file "参考マクロ/ファイル名.bas"
```

### 9.2 エンコーディング管理
- **参考マクロ**: Shift-JIS（Excelエクスポート）
- **src/**: UTF-8（Claude編集用）
- **macros/**: Shift-JIS（Excel取り込み用）

### 9.3 変換スクリプト
```bash
# 基本的な使用方法
./scripts/bas2sjis src/マクロ名.bas

# 変換結果の確認
ls -la macros/
```

## 10. 重要事項チェックリスト

### 必須項目（これを守らないと失敗する）
- [ ] `Option Explicit`の記述
- [ ] **Activateメソッドの完全排除**（最重要）
- [ ] `Application.ScreenUpdating = False`（基本中の基本）
- [ ] エラーハンドリングと設定復元処理
- [ ] ステータスバーのクリア処理
- [ ] 参考マクロは必ずiconvで変換してから読む

### 条件付き項目
- [ ] `DisplayAlerts = False`（削除・クリア操作がある場合）
- [ ] 進捗表示（長時間処理の場合）
- [ ] 完了表示（ステータスバー）

### 絶対禁止事項
- [ ] Activateメソッドの使用
- [ ] MsgBoxによる正常終了メッセージ
- [ ] 文字化けしたファイルの内容を基にした修正
- [ ] 設定を元に戻さない処理

## 11. Mondayの格言（失敗から生まれた教訓）

> **「Activateを見たら即削除」**
> 
> **「賢い解決策より、正しい解決策を選べ」**
> 
> **「Dictionaryより先にScreenUpdating」**
> 
> **「文字化けは列名誤認識の元凶」**
> 
> **「症状から正しく原因を推測すること」**

---

## 注意事項

- このナレッジは実際の失敗から生まれた実戦的なものです
- 理論より実践、完璧より実用性を重視します
- 同じ失敗を繰り返さないことが最優先です
- **特に画面ちらつき問題は、Activateが真犯人だと覚えておいてください**

*「過去の失敗を無視して、毎回同じ失敗を繰り返してる」 という状況を避けるためのナレッジです。*