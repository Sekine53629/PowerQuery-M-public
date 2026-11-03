// ================================================================================
// ⑫ クエリ名：T_ABC月次（データモデルに読み込む）
//    月×需要のまとまり の ABC 区分。その月を終わりとする直近「ABCの月数」ヶ月の払出で付ける
//    ・まとまり：薬品グループの「需要のまとまり」。未登録の薬は薬品名そのもの
//    ・区分：払出の多い順に累積構成比を取り、≦ Aの線 → A、≦ Bの線 → B、それ以外 → C
//    ・前月の区分：1つ前の出納の月の区分（暦の前月ではなく、出納のある月の並びで1つ前）
//    ・変更：「B→A」のように表示。変わっていなければ空欄
//    ・線は設定テーブルで変える（ABCの月数 3／ABCのAの線 0.7／ABCのBの線 0.9）
// ================================================================================
let
    Prm  = #"設定値",
    Num  = (key as text, dflt as number) as number =>
               let v = Record.FieldOrDefault(Prm, key, null)
               in  if v = null or (v is text and Text.Trim(v) = "") then dflt else Number.From(v),
    N     = Number.RoundDown(Num("ABCの月数", 3)),
    ALine = Num("ABCのAの線", 0.7),
    BLine = Num("ABCのBの線", 0.9),

    Mon    = Table.SelectRows(#"T_月次", each [出納データあり]),
    Grp    = Table.Buffer(Table.SelectColumns(#"M_薬品", {"薬品名", "需要のまとまり"})),
    Keyed  = Table.AddColumn(
                 Table.ExpandTableColumn(
                     Table.NestedJoin(Mon, {"薬品名"}, Grp, {"薬品名"}, "g", JoinKind.LeftOuter), "g", {"需要のまとまり"}),
                 "まとまり", each [需要のまとまり] ?? [薬品名], type text),
    KeyedB = Table.Buffer(Table.SelectColumns(Keyed, {"月", "まとまり", "払出数量"})),
    Months = List.Buffer(List.Sort(List.Distinct(Mon[月]))),

    // ---- 1つの月の区分表（その月を終わりとする N ヶ月の払出） ----
    OneMonth = (m as date) as table =>
        let
            i      = List.PositionOf(Months, m),
            win    = List.LastN(List.FirstN(Months, i + 1), N),
            rows   = Table.SelectRows(KeyedB, each List.Contains(win, [月])),
            sums   = Table.Group(rows, {"まとまり"}, {{"払出", each List.Sum([払出数量]) ?? 0, type number}}),
            sorted = Table.Sort(sums, {{"払出", Order.Descending}, {"まとまり", Order.Ascending}}),
            total  = List.Sum(sorted[払出]),
            outs   = List.Buffer(sorted[払出]),
            idx    = Table.AddIndexColumn(sorted, "i", 0, 1),
            cum    = Table.AddColumn(idx, "累積", each
                         if total = null or total = 0 then 0 else List.Sum(List.FirstN(outs, [i] + 1)) / total, type number),
            share  = Table.AddColumn(cum, "構成比", each if total = null or total = 0 then 0 else [払出] / total, type number),
            cls    = Table.AddColumn(share, "区分", each
                         if [累積] <= ALine then "A" else if [累積] <= BLine then "B" else "C", type text),
            withM  = Table.AddColumn(cls, "月", (x) => m, type date)
        in
            Table.SelectColumns(withM, {"月", "まとまり", "払出", "構成比", "累積", "区分"}),

    EmptyT = #table(type table [月 = date, まとまり = text, 払出 = number, 構成比 = number, 累積 = number, 区分 = text], {}),
    All    = if List.IsEmpty(Months) then EmptyT else Table.Combine(List.Transform(Months, OneMonth)),

    // ---- 前月（出納のある月の並びで1つ前）の区分 ----
    NextOf = (m as date) as nullable date =>
        let i = List.PositionOf(Months, m)
        in  if i < 0 or i >= List.Count(Months) - 1 then null else Months{i + 1},
    Prev   = Table.SelectRows(
                 Table.AddColumn(Table.SelectColumns(All, {"月", "まとまり", "区分"}), "次の月", each NextOf([月]), type nullable date),
                 each [次の月] <> null),
    PrevR  = Table.RenameColumns(Table.SelectColumns(Prev, {"次の月", "まとまり", "区分"}), {{"区分", "前月の区分"}}),
    J      = Table.ExpandTableColumn(
                 Table.NestedJoin(All, {"月", "まとまり"}, PrevR, {"次の月", "まとまり"}, "p", JoinKind.LeftOuter),
                 "p", {"前月の区分"}),
    Chg    = Table.AddColumn(J, "変更", each
                 if [前月の区分] = null or [前月の区分] = [区分] then "" else [前月の区分] & "→" & [区分], type text),
    Ren    = Table.RenameColumns(Chg, {{"まとまり", "需要のまとまり"}}),
    Typed  = Table.TransformColumnTypes(Ren, {
                 {"月", type date}, {"需要のまとまり", type text}, {"払出", type number},
                 {"構成比", Percentage.Type}, {"累積", Percentage.Type}, {"区分", type text},
                 {"前月の区分", type text}, {"変更", type text}}),
    Result = Table.Sort(Typed, {{"月", Order.Ascending}, {"累積", Order.Ascending}})
in
    Result
