# M言語（Power Query）基礎知識

VBAを「古代言語」と呼ぶなら、M言語は「謎の言語」だ。関数型で、Excelユーザーには異質。でも強力。

## M言語とは

Power Queryで使用される関数型言語。ETL（Extract, Transform, Load）に特化。

### 特徴
- **関数型** - すべてが式
- **遅延評価** - 必要になるまで実行されない
- **強い型付け** - でも型推論がある
- **大文字小文字を区別** - これ重要

## 基本構文

### let式
```m
let
    // 変数定義（ステップ）
    Source = Excel.CurrentWorkbook(),
    Table1 = Source{[Name="テーブル1"]}[Content],
    FilteredRows = Table.SelectRows(Table1, each [売上] > 1000)
in
    FilteredRows  // 最終的な出力
```

### なぜletとinなのか
- **let**: ステップを定義
- **in**: 最終結果を指定
- 各ステップは前のステップを参照可能

## よくあるパターン

### 1. テーブルの取得
```m
// 現在のブックから
Source = Excel.CurrentWorkbook(){[Name="テーブル名"]}[Content]

// 外部ファイルから
Source = Excel.Workbook(File.Contents("C:\data.xlsx"))
```

### 2. 列の型変更
```m
// 型変更は明示的に
ChangedType = Table.TransformColumnTypes(
    PreviousStep,
    {
        {"日付", type date},
        {"金額", type number},
        {"名前", type text}
    }
)
```

### 3. フィルタリング
```m
// each は行を表す
FilteredRows = Table.SelectRows(
    Source, 
    each [売上] > 1000 and [地域] = "東京"
)
```

### 4. 列の追加
```m
// カスタム列
AddedColumn = Table.AddColumn(
    Source, 
    "利益率", 
    each [利益] / [売上],
    type number
)
```

## Power Queryの罠

### 1. 型の自動認識
```m
// 危険：自動型認識は日付を間違える
Source = Csv.Document(File.Contents("data.csv"))

// 安全：明示的に型指定
Source = Csv.Document(
    File.Contents("data.csv"),
    [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.None]
)
```

### 2. エラー処理
```m
// エラーを含む行を除外
CleanedTable = Table.RemoveRowsWithErrors(Source)

// エラーを別の値に置換
ReplacedErrors = Table.ReplaceErrorValues(
    Source,
    {{"列名", "エラー"}}
)
```

### 3. パフォーマンス問題
```m
// 遅い：毎回ファイル読み込み
BadPattern = Table.AddColumn(Source, "Lookup", 
    each Excel.Workbook(File.Contents("master.xlsx"))...
)

// 速い：一度読み込んでバッファ
let
    MasterData = Table.Buffer(
        Excel.Workbook(File.Contents("master.xlsx"))...
    ),
    GoodPattern = Table.AddColumn(Source, "Lookup",
        each MasterData{[ID=[ID]]}[Value]
    )
in
    GoodPattern
```

## VBAとの連携

### VBAからPower Queryを実行
```vba
Sub RefreshPowerQuery()
    ' クエリの更新
    ThisWorkbook.Connections("Query - クエリ名").Refresh
    
    ' すべてのクエリを更新
    ThisWorkbook.RefreshAll
End Sub
```

### Power Queryの結果をVBAで使用
```vba
Sub UsePowerQueryResult()
    ' Power Queryの出力テーブルを取得
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("Sheet1").ListObjects("クエリ名")
    
    ' データ処理
    Dim data As Variant
    data = tbl.DataBodyRange.Value
End Sub
```

## M言語のベストプラクティス

### 1. ステップ名は意味のあるものに
```m
let
    // 悪い例
    Table1 = ...,
    Table2 = ...,
    
    // 良い例
    RawData = ...,
    FilteredByDate = ...,
    AggregatedSales = ...
```

### 2. 早期のフィルタリング
```m
// データ量を早めに減らす
let
    Source = Sql.Database("server", "database"),
    FilteredEarly = Table.SelectRows(Source, each [Date] >= #date(2025,1,1)),
    // この後で複雑な処理
```

### 3. Table.Bufferの適切な使用
```m
// 何度も参照するテーブルはバッファ
BufferedLookup = Table.Buffer(LookupTable)
```

## よくある質問

### Q: なぜM言語？VBAでいいじゃん
A: データ変換に特化してる。100万行でも軽い。VBAだと死ぬ。

### Q: 関数型って何？
A: すべてが式で、副作用がない。慣れれば強力。

### Q: デバッグが難しい
A: 各ステップの結果を確認できる。これがlet式の利点。

## Mondayの感想

M言語は確かに変態的だが、データ処理には最適。VBAで1時間かかる処理が1分で終わることもある。

ただし、学習曲線が急。「なんでこんな書き方？」って思うこと多数。でも慣れれば、VBAには戻れない（データ処理に関しては）。

---

*「M言語は謎だが、使えれば最強」 - Monday*