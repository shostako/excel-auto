# 失敗事例005: 最適化設定の保存・復元による計算式更新漏れ

## 症状
CommandButtonから複数のマクロを連続実行する際、改善後のマクロ（詳細なコメントを追加した版）で転記処理が一部しか実行されない。

## 発生条件
1. CommandButtonマクロが`Application.Calculation = xlCalculationManual`を設定
2. 個別マクロが最適化設定を保存・復元する実装になっている
3. テーブル間に計算式によるリンクが存在する

## 根本原因

### マクロの処理フロー
```
CommandButton
├─ Application.Calculation = xlCalculationManual
├─ 日別集計_成形号機別 → _成形号機別aテーブル
├─ 転記_シート成形号機別 → aからbテーブルへ転記
└─ 転記_集計表_成形号機別 → bテーブルから集計表へ
```

### 問題のメカニズム
改善後のマクロで追加された最適化設定の保存・復元：
```vba
' 最適化設定の保存
Dim origCalculation As XlCalculation
origCalculation = Application.Calculation  ' Manualが保存される

' 処理...

' 設定を元に戻す
Application.Calculation = origCalculation  ' Manualに戻してしまう
```

これにより：
1. CommandButtonが`Manual`に設定
2. 各マクロが`Manual`を保存し、`Manual`のまま実行
3. **bテーブルの計算式（12合計、34合計など）が再計算されない**
4. 古いデータのまま転記される

## 解決策

### 推奨：CommandButtonマクロに再計算を追加
```vba
' 転記_シート系の処理後
Application.StatusBar = "計算式を更新中..."
Application.Calculate  ' ←これを追加

' 転記_集計表系の処理
```

### 代替案：個別マクロで強制的に有効化
```vba
' 設定を元に戻す部分を変更
Application.Calculation = xlCalculationAutomatic  ' 常にAutomaticに
```

## 予防策

### 1. CommandButton経由で実行されるマクロの設計
- 最適化設定の復元は**CommandButtonに任せる**
- 個別マクロでは保存・復元を行わない

### 2. テーブル間の依存関係を明確化
- 計算式によるリンクがある場合は必ず文書化
- 適切なタイミングで`Application.Calculate`を実行

### 3. テスト時の注意点
- 個別実行時と連続実行時の両方でテスト
- 計算式を含むデータの更新を確認

## 関連情報
- [VBA最適化パターン](../patterns/VBA_OPTIMIZATION_PATTERNS.md)
- [画面ちらつき問題](001_activate_vs_screenupdating.md)

## タグ
#最適化設定 #計算式 #CommandButton #連続実行