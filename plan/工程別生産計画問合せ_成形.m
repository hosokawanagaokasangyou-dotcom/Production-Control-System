// Encoding: UTF-8 with BOM. Open as UTF-8 if characters look wrong.
// Source workbook: 工程別生産計画問合せ問合せ_*.xlsx
// (1) Skip first 4 rows  (2) Header = Excel row6 text + row5 text per column
// (3) Drop columns containing 加工時間 or 加工速度
// (4) Remove substring 加工数量 from headers EXCEPT when the header is exactly "加工数量" only (keep that column name)
// (5) Prefix year from first data row of 受注日 for M/D headers -> e.g. 2026/04/01(水)
// (6) Normalize to YYYY/MM/DD (strip weekday parens from display path via BeforeParen)
//
// Edit Path and SheetName for your environment.

let
    Path =
        "C:\\工程管理AIプロジェクト_JAVA\\Production-Control-System\\plan\\工程別生産計画問合せ問合せ_20260501_153807.xlsx",
    SheetName = "工程別生産計画問合せ",

    Binary = File.Contents(Path),
    Wb = Excel.Workbook(Binary, null, true),
    SheetTbl =
        try Wb{[Item = SheetName, Kind = "Sheet"]}[Data]
        otherwise Wb{0}[Data],

    AfterSkip4 = Table.Skip(SheetTbl, 4),
    ColNames0 = Table.ColumnNames(AfterSkip4),

    NullBlankRow =
        Table.TransformColumns(
            AfterSkip4,
            List.Transform(ColNames0, (col) => {col, each if _ = "" then null else _})
        ),

    RecRow5 = Table.First(NullBlankRow),
    RecRow6 = Table.First(Table.Skip(NullBlankRow, 1)),
    DataRows = Table.Skip(NullBlankRow, 2),

    CombinedHeaders =
        List.Transform(
            ColNames0,
            each
                let
                    s6 = try Text.From(Record.Field(RecRow6, _)) otherwise "",
                    s5 = try Text.From(Record.Field(RecRow5, _)) otherwise ""
                in
                    s6 & s5
        ),

    WithHeaders = Table.RenameColumns(DataRows, List.Zip({ColNames0, CombinedHeaders})),

    NamesA = Table.ColumnNames(WithHeaders),
    KeepCols =
        List.Select(
            NamesA,
            each (not Text.Contains(_, "加工時間")) and (not Text.Contains(_, "加工速度"))
        ),
    AfterDropSpeed = Table.SelectColumns(WithHeaders, KeepCols, MissingField.Ignore),

    NamesB = Table.ColumnNames(AfterDropSpeed),
    NamesStripped =
        List.Transform(
            NamesB,
            each
                let
                    t = Text.Trim(_)
                in
                    if t = "加工数量" then
                        t
                    else
                        Text.Replace(t, "加工数量", "", Replacer.ReplaceText)
        ),
    AfterStripQty = Table.RenameColumns(AfterDropSpeed, List.Zip({NamesB, NamesStripped})),

    ToDate =
        (v as any) as nullable date =>
            if v = null then
                null
            else if Value.Type(v) = type date then
                v
            else if Value.Type(v) = type datetime then
                Date.From(v)
            else if Value.Type(v) = type number then
                Date.AddDays(#date(1899, 12, 30), Int64.From(v))
            else
                try Date.FromText(Text.Trim(Text.From(v))) otherwise null,

    OrderDatesRaw = try Table.Column(AfterStripQty, "受注日") otherwise {},
    FirstRowOrderDate =
        if List.Count(OrderDatesRaw) > 0 then
            ToDate(OrderDatesRaw{0})
        else
            null,
    ValidOrderDates = List.RemoveNulls(List.Transform(OrderDatesRaw, each ToDate(_))),

    BaseDate =
        if FirstRowOrderDate <> null then
            FirstRowOrderDate
        else if List.Count(ValidOrderDates) > 0 then
            ValidOrderDates{0}
        else
            Date.From(DateTime.LocalNow()),
    BaseYear = Date.Year(BaseDate),
    BaseMonth = Date.Month(BaseDate),

    NamesC = Table.ColumnNames(AfterStripQty),
    NamesFinal =
        List.Transform(
            NamesC,
            each
                let
                    hRaw = Text.Trim(_),
                    BeforeParen =
                        if Text.Contains(hRaw, "(") then
                            Text.Trim(Text.BeforeDelimiter(hRaw, "("))
                        else
                            hRaw,
                    Parts = Text.Split(BeforeParen, "/"),
                    Out =
                        if List.Count(Parts) = 3
                            and (try Number.From(Parts{0}) <> null otherwise false)
                            and (try Number.From(Parts{1}) <> null otherwise false)
                            and (try Number.From(Parts{2}) <> null otherwise false)
                        then
                            let
                                yy = Number.From(Parts{0}),
                                mm = Number.From(Parts{1}),
                                dd = Number.From(Parts{2})
                            in
                                Text.From(yy)
                                & "/"
                                & Text.PadStart(Text.From(mm), 2, "0")
                                & "/"
                                & Text.PadStart(Text.From(dd), 2, "0")
                        else if List.Count(Parts) = 2
                            and (try Number.From(Parts{0}) <> null otherwise false)
                            and (try Number.From(Parts{1}) <> null otherwise false)
                        then
                            let
                                Mo = Number.From(Parts{0}),
                                Dy = Number.From(Parts{1}),
                                Yy =
                                    if BaseMonth >= 11 and Mo <= 2 then
                                        BaseYear + 1
                                    else if BaseMonth <= 2 and Mo >= 11 then
                                        BaseYear - 1
                                    else
                                        BaseYear
                            in
                                Text.From(Yy)
                                & "/"
                                & Text.PadStart(Text.From(Mo), 2, "0")
                                & "/"
                                & Text.PadStart(Text.From(Dy), 2, "0")
                        else
                            hRaw
                in
                    Out
        ),

    Result = Table.RenameColumns(AfterStripQty, List.Zip({NamesC, NamesFinal}))
in
    Result
