# VBAマクロ開発クイックリファレンス

急いでいる時、最低限これだけ確認すれば大丈夫。詳細は[メインナレッジベース](EXCEL_MACRO_KNOWLEDGE_BASE.md)を参照。

## 🚨 絶対守るべき6つのルール

```vba
□ Option Explicit                        ' 変数宣言を強制
□ Application.ScreenUpdating = False     ' 画面ちらつき防止
□ Activateメソッド完全排除               ' ws.Activate禁止！
□ On Error GoTo ErrorHandler             ' エラー処理必須
□ 参考マクロ: iconv -f SHIFT-JIS -t UTF-8  ' 文字化け防止
□ モジュール名は「_改X」で終わらせる      ' VBA文字数制限対応
```

## 📊 優先度別ルール

### 🔴 最優先（破ると動作不良・本番事故）

#### 1. 画面ちらつき対策
```vba
' 必須設定
Application.ScreenUpdating = False
' Activateは絶対ダメ
' ws.Activate  ← 削除！
' Range("A1").Select  ← これも削除！
```

#### 2. 文字化け対策（列名誤認識事故防止）
```bash
# 参考マクロを読む前に必ず実行
iconv -f SHIFT-JIS -t UTF-8 "参考マクロ/ファイル名.bas" | head -100
```

#### 3. 変数管理（タイポ・重複エラー防止）
```vba
Option Explicit  ' 最上部に必須
Dim rowIdx As Long  ' 明確な変数名（iやjは避ける）
```

### 🟡 推奨（品質・保守性向上）

#### 1. エラーハンドリング基本構造
```vba
On Error GoTo ErrorHandler
' 処理
Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True  ' 設定復元
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub
```

#### 2. 進捗表示（長時間処理）
```vba
Application.StatusBar = "処理中... " & i & "/" & totalCount
' 完了時
Application.StatusBar = False
```

#### 3. 実用的コメント書式（推奨）
```vba
' ========================================
' マクロ名: m日別集計_モールFR別
' 処理概要: 日別データをモールFR別に集計して出力
' ソーステーブル: シート「sysdata」テーブル「_sysdata」
' ターゲットテーブル: シート「日別集計」テーブル「_日別FR別」
' ========================================
```

### 🟢 状況次第（必要な時だけ）

- `Application.Calculation = xlCalculationManual` （計算式が多い場合）
- `Application.DisplayAlerts = False` （削除処理がある場合）
- ユーザーフォーム実装（大量エラー表示が必要な場合）

## 🔍 症状別1分診断

| 症状 | 原因（90%これ） | 対処法 |
|------|------------------|---------|
| 画面がちらつく・もたもた動く | Activateメソッド使用 | Activate削除 + ScreenUpdating = False |
| 処理が異常に遅い | ScreenUpdating忘れ | 最初に必ず設定 |
| 文字化けで動かない | Shift-JIS未変換 | iconv変換してから読む |
| 変数エラー | Option Explicit なし | 最上部に追加 |
| エラー時に画面固まる | 設定復元忘れ | ErrorHandlerで必ず復元 |
| VBAで「名前が無効」エラー | モジュール名が長すぎ | 「_改X」形式に短縮 |

## 🚀 最速スタートテンプレート

```vba
Option Explicit

Sub マクロ名()
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    
    ' ここに処理を書く
    ' Activateは使わない！
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub
```

## 📚 詳細情報へのリンク

- **なぜActivateがダメなのか？** → [画面ちらつき問題の真犯人](EXCEL_MACRO_KNOWLEDGE_BASE.md#21-画面ちらつき問題の真犯人最重要)
- **変数管理の失敗パターン** → [Mondayがよくやらかすやつ](EXCEL_MACRO_KNOWLEDGE_BASE.md#3-変数管理の失敗パターンmondayがよくやらかすやつ)
- **コード例が欲しい** → [実装例集](IMPLEMENTATION_EXAMPLES.md)

---

**緊急時はこのページだけ見れば大丈夫。** 詳細が必要になったらメインナレッジベースへ。