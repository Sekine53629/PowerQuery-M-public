// ================================================================================
// ⑬ クエリ名：T_まとまり月次（データモデルに読み込む）
//    月×薬品の1行。同じ「需要のまとまり」（屋号が2つ以上）の中での払出の構成比と、出荷調整の疑い
//    ・構成比：その月のまとまりの払出に対する、その薬の払出の割合
//    ・前月：出納のある月の並びで1つ前
//    ・G01：構成比が前月から設定「出荷調整の構成比の変化」以上動き、まとまり全体の払出の変化が
//           設定「出荷調整の合計の許容」以内（＝需要は変わらないのに屋号が入れ替わった）
//    ・G02：仕入 0 が2ヶ月続き、当月の払出 ≧ 設定「極小需要の月の払出」、当月の在庫 < 当月の払出
//           （ほとんど出ない薬は除く：箱単位の発注では仕入 0 の月が普通にある）
//    ・理由：G01・G02 の文言（数字入り）
// ================================================================================
let
    Prm  = #"設定値",
    Num  = (key as text, dflt as number) as number =>
               let v = Record.FieldOrDefault(Prm, key, null)
               in  if v = null or (v is text and Text.Trim(v) = "") then dflt else Number.From(v),
    PT      = Num("出荷調整の構成比の変化", 0.3),
    TOL     = Num("出荷調整の合計の許容", 0.3),
    TinyOut = Num("極小需要の月の払出", 100),
    F       = (n as nullable number) as text => if n = null then "" else Number.ToText(Number.Round(n, 0), "N0", "ja-JP"),
    Signed  = (n as number) as text => (if n >= 0 then "+" else "") & Text.From(Number.Round(n, 0)),

    Mon    = Table.SelectRows(#"T_月次", each [出納データあり]),
    Drugs  = Table.Buffer(Table.SelectColumns(#"M_薬品", {"薬品名", "需要のまとまり"})),
    J0     = Table.ExpandTableColumn(
                 Table.NestedJoin(Mon, {"薬品名"}, Drugs, {"薬品名"}, "g", JoinKind.LeftOuter), "g", {"需要のまとまり"}),
    Rows   = Table.SelectRows(J0, each [需要のまとまり] <> null),

    // 屋号が2つ以上あるまとまりだけ
    Cnt    = Table.Group(Rows, {"需要のまとまり"}, {{"屋号数", each List.Count(List.Distinct([薬品名])), Int64.Type}}),
    Multi  = List.Buffer(Table.SelectRows(Cnt, each [屋号数] >= 2)[需要のまとまり]),
    R2     = Table.SelectRows(Rows, each List.Contains(Multi, [需要のまとまり])),

    Base   = Table.Group(R2, {"需要のまとまり", "月", "薬品名"}, {
                 {"払出",     each List.Sum([払出数量]) ?? 0, type number},
                 {"仕入",     each List.Sum([仕入数量]) ?? 0, type number},
                 {"在庫",     each List.Max([在庫数量]), type nullable number}}),
    Tot    = Table.Group(Base, {"需要のまとまり", "月"}, {{"まとまりの払出", each List.Sum([払出]) ?? 0, type number}}),
    WithT  = Table.ExpandTableColumn(
                 Table.NestedJoin(Base, {"需要のまとまり", "月"}, Tot, {"需要のまとまり", "月"}, "t", JoinKind.LeftOuter),
                 "t", {"まとまりの払出"}),
    Share  = Table.AddColumn(WithT, "構成比", each if [まとまりの払出] = 0 then 0 else [払出] / [まとまりの払出], type number),

    // ---- 前月（出納のある月の並びで1つ前）の値をつなぐ ----
    Months = List.Buffer(List.Sort(List.Distinct(Mon[月]))),
    NextOf = (m as date) as nullable date =>
        let i = List.PositionOf(Months, m)
        in  if i < 0 or i >= List.Count(Months) - 1 then null else Months{i + 1},
    Prev   = Table.SelectRows(
                 Table.AddColumn(Table.SelectColumns(Share, {"月", "薬品名", "払出", "仕入", "構成比", "まとまりの払出"}),
                                 "次の月", each NextOf([月]), type nullable date),
                 each [次の月] <> null),
    PrevR  = Table.RenameColumns(Table.SelectColumns(Prev, {"次の月", "薬品名", "払出", "仕入", "構成比", "まとまりの払出"}),
                 {{"払出", "前月の払出"}, {"仕入", "前月の仕入"}, {"構成比", "前月の構成比"}, {"まとまりの払出", "まとまりの払出_前月"}}),
    J      = Table.ExpandTableColumn(
                 Table.NestedJoin(Share, {"月", "薬品名"}, PrevR, {"次の月", "薬品名"}, "p", JoinKind.LeftOuter),
                 "p", {"前月の払出", "前月の仕入", "前月の構成比", "まとまりの払出_前月"}),

    Calc   = Table.AddColumn(J, "計算", each
        let
            chg   = if [前月の構成比] = null then null else [構成比] - [前月の構成比],
            totCh = if [まとまりの払出_前月] = null or [まとまりの払出_前月] = 0 then null
                    else ([まとまりの払出] - [まとまりの払出_前月]) / [まとまりの払出_前月],
            g01   = chg <> null and totCh <> null and Number.Abs(chg) >= PT and Number.Abs(totCh) <= TOL,
            g02   = [仕入] = 0 and ([前月の仕入] ?? 0) = 0 and [払出] >= TinyOut and ([前月の払出] ?? 0) > 0
                    and [在庫] <> null and [在庫] < [払出],
            why   = if g01 then "構成比が " & Signed(chg * 100) & " ポイント動きました（まとまり全体は " & Signed(totCh * 100) & "%）"
                    else if g02 then "仕入がないまま払出が続き、在庫 " & F([在庫]) & "（月の払出 " & F([払出]) & "）"
                    else ""
        in
            [変化 = chg, まとまりの変化 = totCh, G01 = g01, G02 = g02, 理由 = why]),
    E      = Table.ExpandRecordColumn(Calc, "計算", {"変化", "まとまりの変化", "G01", "G02", "理由"}),
    Out    = Table.SelectColumns(E, {"月", "薬品名", "需要のまとまり", "払出", "仕入", "在庫", "まとまりの払出", "構成比",
                 "前月の構成比", "変化", "まとまりの払出_前月", "まとまりの変化", "G01", "G02", "理由"}),
    Typed  = Table.TransformColumnTypes(Out, {
                 {"月", type date}, {"薬品名", type text}, {"需要のまとまり", type text},
                 {"払出", type number}, {"仕入", type number}, {"在庫", type number}, {"まとまりの払出", type number},
                 {"構成比", Percentage.Type}, {"前月の構成比", Percentage.Type}, {"変化", Percentage.Type},
                 {"まとまりの払出_前月", type number}, {"まとまりの変化", Percentage.Type},
                 {"G01", type logical}, {"G02", type logical}, {"理由", type text}}),
    Result = Table.Sort(Typed, {{"月", Order.Ascending}, {"需要のまとまり", Order.Ascending}, {"薬品名", Order.Ascending}})
in
    Result
