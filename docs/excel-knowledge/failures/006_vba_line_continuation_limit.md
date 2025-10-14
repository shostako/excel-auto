# VBA行継続文字（アンダースコア）の制限エラー

## 発生日
2025-09-23

## エラー内容
```
Microsoft Visual Basic for Applications
行継続文字（_）を使いすぎています。
```

## 問題の詳細
Power Query（M言語）のクエリ文字列を生成するVBA関数で、長大な文字列を`_`で継続して連結しようとしたところ、VBAの制限に抵触。

### 制限事項
- **VBAの行継続文字制限**: 1つの論理行に最大**24個**まで
- 25個目以降を使用するとコンパイルエラー

## 失敗したコード例
```vba
' NG: 行継続文字が多すぎる
newFormula = "let" & vbCrLf & _
            "    前月変更された型 = Table.TransformColumnTypes(前月集計1_Table,{{""日付"", type date}, {""開始時間"", type datetime}, {""終了時間"", type datetime}, " & _
            "{""所要時間"", type number}, {""型替"", type number}, {""号機"", Int64.Type}, " & _
            ' ... 20行以上続く ...
            "{""補助1"", Int64.Type}})," & vbCrLf & _
            ' さらに続く...
```

## 解決方法

### 方法1: 変数分割方式（推奨）
```vba
' OK: 長い文字列を変数に分割
Dim typeList As String
Dim columnList As String

' 型定義リストを変数に格納
typeList = "{{""日付"", type date}, {""開始時間"", type datetime}, "
typeList = typeList & "{""終了時間"", type datetime}, "
typeList = typeList & "{""所要時間"", type number}, "
' ... 必要なだけ連結

' クエリ本体で変数を使用
newFormula = "let" & vbCrLf
newFormula = newFormula & "    前月変更された型 = Table.TransformColumnTypes(前月集計1_Table," & typeList & ")," & vbCrLf
```

### 方法2: 文字列連結方式
```vba
' OK: &演算子で連結（行継続文字不使用）
newFormula = "let" & vbCrLf
newFormula = newFormula & "    // 前月データ" & vbCrLf
newFormula = newFormula & "    前月ソース = Excel.Workbook(..." & vbCrLf
' 行継続文字を使わずに連結
```

## ベストプラクティス

### 1. 長い文字列は変数に分割
- 可読性が向上
- 再利用可能
- デバッグしやすい

### 2. 共通部分を定数化
```vba
' 共通部分は定数や変数に
Const BASE_PATH = "Z:\全社共有\オート事業部\日報\"
Dim fullPath As String
fullPath = BASE_PATH & folderName & "\" & fileName
```

### 3. 配列とJoinの活用
```vba
' 大量の項目がある場合
Dim columns() As Variant
columns = Array("日付", "品番", "数量", "不良数")
Dim columnList As String
columnList = "{""" & Join(columns, """, """) & """}"
```

## 教訓
- VBAには見えない制限が多い
- 長大な文字列生成時は分割を検討
- Power QueryのM言語は特に長くなりがちなので注意

## 関連ナレッジ
- [VBA最適化パターン](../patterns/VBA_OPTIMIZATION_PATTERNS.md)
- [VBA基本テクニック](../techniques/VBA_BASIC_TECHNIQUES.md)