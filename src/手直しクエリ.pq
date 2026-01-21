let
    ソース = Access.Database(File.Contents("Z:\全社共有\オート事業部\日報\不良集計\不良集計表\2026年\不良調査表DB-2026.accdb"), [CreateNavigationProperties=true]),
    テーブル = ソース{[Schema="",Item="_不良集計ゾーン別"]}[Data],
    削除された他の列 = Table.SelectColumns(テーブル,{"ID", "日付", "品番", "ロット", "発見", "ゾーン", "番号", "数量", "差戻し"}),
    変更された型 = Table.TransformColumnTypes(削除された他の列,{{"数量", Int64.Type}, {"差戻し", Int64.Type}})
in
    変更された型
