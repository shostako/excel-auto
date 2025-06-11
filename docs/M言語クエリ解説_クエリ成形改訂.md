# M言語クエリ「クエリ成形改訂」の解説

## 1. 概要

製造業の成形工程における生産実績データを、日別・品番別・機械別に集計・整形するM言語クエリです。生産管理システムから出力された詳細なトランザクションデータを、分析可能な形式に変換し、特に夜勤対応や不良品情報の統合などの製造現場特有の要件に対応しています。

## 2. 業務フロー

```
生産管理システム
    ↓ エクスポート（月次データ）
Excelファイル（生データ）
    ↓ Power Query実行
本クエリによる整形処理
    ├─ 夜勤時間帯の日付調整
    ├─ SS機械の重複削除
    ├─ グループ化による集計
    └─ 実績数量0行の統合
    ↓
整形済みデータ
    ↓ Excelマクロ処理
日別集計表・グラフ・分析レポート
```

## 3. 主要処理の詳細解説

### 3.1 データソースの読み込み（行1-5）
```m
ソース = Excel.Workbook(File.Contents("Z:\全社共有\生産管理課\生産実績データ\2506.xlsx"), null, true),
Sheet1_Sheet = ソース{[Item="Sheet1",Kind="Sheet"]}[Data],
昇格されたヘッダー数 = Table.PromoteHeaders(Sheet1_Sheet, [PromoteAllScalars=true]),
変更された型 = Table.TransformColumnTypes(...)
```

**処理内容**：
- 共有フォルダ上の月次実績ファイルを読み込み
- 22列のデータを適切な型（日付、数値、テキスト）に変換
- ヘッダー行の設定により、列名での参照を可能に

### 3.2 夜勤対応の日付調整（行8-19）
```m
日付列追加 = Table.AddColumn(削除された空白行, "日付", each
    let
        dateTimeValue = try DateTime.From([時刻]) otherwise null,
        作業日の日付 = [作業日],
        時刻部分 = if dateTimeValue <> null then DateTime.Time(dateTimeValue) else null
    in
    // 時刻部分が取得でき、かつ0時より大きく8時以下の場合
    if 時刻部分 <> null and 時刻部分 > #time(0, 0, 0) and 時刻部分 <= #time(8, 0, 0) then
        Date.AddDays(作業日の日付, -1)
    else
        作業日の日付,
type date)
```

**処理内容**：
- 午前0時～8時の作業実績は前日の日付として扱う
- 例：6月2日午前3時の作業 → 6月1日の実績として記録
- 24時間稼働の製造現場での夜勤シフトに対応

### 3.3 SS機械の重複削除処理（行24-40）
```m
重複削除済み行 = let
    SS行 = Table.SelectRows(フィルターされた行, each ([機械コード] = "SS01" or ... or [機械コード] = "SS05")),
    非SS行 = Table.SelectRows(フィルターされた行, each ([機械コード] <> "SS01" and ... and [機械コード] <> "SS05")),
    
    重複削除済みSS行 = Table.FromRecords(
        List.Combine(
            List.Transform(
                Table.Group(SS行, {"実績数量", "時刻"}, {{"Rows", each _, type table}})[Rows],
                each Table.ToRecords(Table.FirstN(Table.Sort(_, {{"品番・図番", Order.Ascending}}), 1))
            )
        )
    ),
    
    結合結果 = Table.Combine({非SS行, 重複削除済みSS行})
in
    結合結果
```

**処理内容**：
- SS01～SS05の射出成形機では同じ実績が複数記録される場合がある
- 実績数量と時刻が同じレコードをグループ化し、品番順で最初の1件のみを残す
- 他の機械のデータはそのまま保持

### 3.4 統合グループ化処理（行42-71）
```m
統合グループ化 = Table.Group(重複削除済み行, {"日付", "品番・図番", "機械コード"}, {
    {"実績数量", each List.Sum([実績数量]), type nullable number},
    {"不良数量", each List.Sum([不良数量]), type nullable number},
    {"段取時間", each 
        let
            段取時間計算値 = List.Sum(List.Transform(Table.SelectRows(_, each [作業区分] = "段取完了")[加工時間], each _ / 60))
        in
            if 段取時間計算値 = null then 0 else 段取時間計算値, 
        type nullable number},
    {"稼働時間", each List.Sum(List.Transform(Table.SelectRows(_, each [作業区分] = "加工完了")[加工時間], each _ / 60)), type nullable number},
    {"不良区分", each
        let
            ValidRowsTable = Table.SelectRows(_,
                (row) => row[不良区分略称] <> null and row[不良区分略称] <> ""
            ),
            CombinedTextList = List.Transform(
                Table.ToRecords(ValidRowsTable),
                (record) => record[不良区分略称] & Text.From(record[不良数量])
            )
        in
            Text.Combine(CombinedTextList, ","),
        type nullable text
    }
})
```

**処理内容**：
- 日付・品番・機械コードの組み合わせごとに集計
- 実績数量と不良数量を合計
- 段取時間：「段取完了」レコードの加工時間を分単位で集計（nullの場合は0）
- 稼働時間：「加工完了」レコードの加工時間を分単位で集計
- 不良区分：「不良区分略称+不良数量」の形式でカンマ区切りのテキストに統合

### 3.5 実績数量0行の統合処理（行73-116）
```m
処理済みグループ = Table.Group(統合グループ化, {"品番・図番", "機械コード"}, {
    {"ProcessedData", (currentGroupTable as table) =>
        let
            zero実績行 = Table.SelectRows(currentGroupTable, each [実績数量] = 0),
            nonZero実績行 = Table.SelectRows(currentGroupTable, each [実績数量] <> 0),
            resultTable = if Table.IsEmpty(zero実績行) or Table.IsEmpty(nonZero実績行) then
                            nonZero実績行
                        else
                            let
                                zero不良数量合計 = List.Sum(List.RemoveNulls(zero実績行[不良数量])),
                                // 最も実績数量が多い行に不良情報を統合
                                targetRowRecord = Table.First(Table.Sort(indexedNonZero実績行, {{"実績数量", Order.Descending}, {"__SortIndex", Order.Ascending}})),
                                // ... 不良数量と不良区分を統合 ...
                            in
                                result
        in
            resultTable, type table}
})
```

**処理内容**：
- 実績数量が0の行（不良品のみ発生）を、同じ品番・機械の実績がある行に統合
- 統合先は実績数量が最も多い行を選択
- 不良数量は加算、不良区分はカンマ区切りで結合
- これにより、不良情報が分散することを防ぐ

### 3.6 最終整形処理（行118-128）
```m
最終結果テーブル展開前 = Table.ExpandTableColumn(処理済みグループ, "ProcessedData", columnsToExpand),
フィルターされた行1 = Table.SelectRows(最終結果テーブル展開前, each ([日付] <> null)),
並べ替えられた列 = Table.ReorderColumns(...),
列名変更 = Table.RenameColumns(並べ替えられた列,{{"品番・図番", "品番"}, {"機械コード", "機械"}, {"実績数量", "実績"}, {"不良数量", "不良"}})
```

**処理内容**：
- グループ化したデータを展開
- null日付の行を除外
- 列の並び順を整理
- 列名を短縮形に変更（後続処理での参照を簡潔に）

## 4. 技術的なポイント

### 4.1 エラーハンドリング
- `try ... otherwise` 構文による安全な型変換
- null値チェックによる不正データの除外
- 条件付き処理による柔軟なデータ処理

### 4.2 パフォーマンス最適化
- グループ化処理の統合により、複数回のテーブルスキャンを回避
- 早期のフィルタリングで処理対象データを削減
- List関数の活用による効率的な集計

### 4.3 データ品質の確保
- 型変換による一貫性のあるデータ処理
- 重複削除による正確な実績カウント
- 不良情報の統合による完全性の維持

### 4.4 M言語の高度な機能活用
- カスタム関数（each式）による複雑な集計ロジック
- レコード操作によるデータ変換
- ネストしたlet式による読みやすいコード構造

## 5. 後続処理との連携

### 5.1 出力データ形式
最終的に以下の8列を持つテーブルが出力されます：
- **日付**: 夜勤調整済みの作業日
- **品番**: 製品識別番号
- **機械**: 加工機械コード
- **実績**: 生産数量の合計
- **不良**: 不良数量の合計
- **稼働時間**: 実際の加工時間（分）
- **段取時間**: 準備・調整時間（分）
- **不良区分**: 不良理由と数量の一覧

### 5.2 Excelマクロとの連携
このクエリの出力は、以下のようなExcelマクロで利用されます：
- **日別集計マクロ**: 日付ごとの生産実績を集計
- **品番別分析マクロ**: 品番ごとの不良率や稼働率を分析
- **機械別稼働率マクロ**: 機械ごとの稼働効率を算出
- **グラフ作成マクロ**: 視覚的なレポート生成

### 5.3 更新と保守
- ファイルパスの変更：2行目の`File.Contents`のパスを変更
- 夜勤時間の調整：15行目の時間判定ロジックを変更
- 新しい機械コードの追加：SS機械の判定条件に追加
- 集計項目の追加：統合グループ化処理に新しい列を追加

このクエリは、製造現場の複雑な要件に対応しながら、保守性と拡張性を確保した設計となっています。