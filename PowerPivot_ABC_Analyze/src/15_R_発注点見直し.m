// ================================================================================
// ⑮ クエリ名：R_発注点見直し（シート「A_発注点」に表として読み込む）
//    A 発注点の切り替え提案：区分が前月から変わった、または直近 N ヶ月の払出が前の N ヶ月から
//    設定「発注点見直しの変化」（初期値 0.2＝±20%）以上動いた「需要のまとまり」だけを載せる
//    ・代表の薬：まとまりの中で直近の払出が一番多い薬（切替先に指定された薬を優先、供給状況が「通常」以外と切替元は後回し）
//    ・発注点の目安 ＝ 直近の1日平均払出 × 設定「発注点の日数」（切り上げ）
//    ・発注量の目安 ＝ 直近の1日平均払出 × 設定「発注量の日数」（切り上げ）
//    ・持ち方：回転（直近の1日平均払出 × 30 ≧ 設定「極小需要の月の払出」）× 予測できる（ずれの大きさ < 予測できるずれの線）
//        回る×予測できる → 自動発注で持つ／回る×できない → 最低在庫を持つ／
//        回らない×できる → 持たずに取り寄せ／回らない×できない → 持たない
//    ・月01～月12：直近12ヶ月（出納のある月）の払出。月12 が一番新しい月（スパークライン用）
//    ・区分変更：変わったら「○」（印刷して読む表なので、真偽ではなく印にする）
//    ・今月の区分：⑫ の区分（A／B／C／D）。量の区分・月平均払出・月平均患者数 は区分の根拠（0.5.0）
//    ・傾向・年間比・患者の傾向：M_薬品 の まとまりの傾向・まとまりの年間比・患者の傾向（代表の薬の行。0.10.0）。
//        直近 N ヶ月 vs その前 N ヶ月（変化）に加えて 年間 と比べ、増加／減少 なら理由に書く。発注量の調整は使う人に委ねる
// ================================================================================
let
    Prm  = #"設定値",
    Num  = (key as text, dflt as number) as number =>
               let v = Record.FieldOrDefault(Prm, key, null)
               in  if v = null or (v is text and Text.Trim(v) = "") then dflt else Number.From(v),
    N        = Number.RoundDown(Num("ABCの月数", 3)),
    RopDays  = Num("発注点の日数", 7),
    OqDays   = Num("発注量の日数", 21),
    TinyOut  = Num("極小需要の月の払出", 100),
    GsLine   = Num("予測できるずれの線", 0.5),
    ChgLine  = Num("発注点見直しの変化", 0.2),
    MonthCols = List.Transform({1 .. 12}, (i) => "月" & Text.PadStart(Text.From(i), 2, "0")),

    Abc    = #"T_ABC月次",
    Months = List.Buffer(List.Sort(List.Distinct(Abc[月]))),
    Cur    = List.Last(Months),
    AbcCur = Table.SelectRows(Abc, each [月] = Cur),

    // ---- 12ヶ月の払出（まとまりごと） ----
    Mon    = Table.SelectRows(#"T_月次", each [出納データあり]),
    Drugs  = Table.Buffer(#"M_薬品"),
    DrugK  = Table.AddColumn(Drugs, "まとまり", each [需要のまとまり] ?? [薬品名], type text),
    KeyMap = Table.Buffer(Table.SelectColumns(DrugK, {"薬品名", "まとまり"})),
    Keyed  = Table.ExpandTableColumn(
                 Table.NestedJoin(Mon, {"薬品名"}, KeyMap, {"薬品名"}, "k", JoinKind.LeftOuter), "k", {"まとまり"}),
    Last12 = List.LastN(Months, 12),
    Ser    = Table.Group(Table.SelectRows(Keyed, each List.Contains(Last12, [月])), {"まとまり", "月"},
                 {{"払出", each List.Sum([払出数量]) ?? 0, type number}}),
    PerGrp = Table.Group(Ser, {"まとまり"}, {{"s", each
                 let t = _
                 in  List.Transform(Last12, (m) =>
                         let r = Table.SelectRows(t, each [月] = m)
                         in  if Table.IsEmpty(r) then 0 else r{0}[払出]), type list}}),

    // ---- 代表の薬（供給状況が通常以外は後回し、直近の払出が多い順） ----
    Rep    = Table.Group(Table.AddColumn(DrugK, "k", each
                 if [切替元] <> null then 0
                 else if [切替先] <> null then 3
                 else if [供給状況] = null or [供給状況] = "通常" then 1 else 2, Int64.Type),
                 {"まとまり"}, {
                 {"屋号数", each Table.RowCount(_), Int64.Type},
                 {"r", each Table.First(Table.Sort(_, {{"k", Order.Ascending}, {"直近の払出合計", Order.Descending}})), type record}}),
    RepE   = Table.ExpandRecordColumn(Rep, "r",
                 {"薬品名", "直近の1日平均払出", "ずれの大きさ", "まとまりの傾向", "まとまりの年間比", "患者の傾向"},
                 {"代表の薬", "直近の1日平均払出", "ずれの大きさ", "傾向", "年間比", "患者の傾向"}),

    J1 = Table.ExpandTableColumn(Table.NestedJoin(AbcCur, {"需要のまとまり"}, PerGrp, {"まとまり"}, "sj", JoinKind.LeftOuter), "sj", {"s"}),
    J2 = Table.ExpandTableColumn(Table.NestedJoin(J1, {"需要のまとまり"}, RepE, {"まとまり"}, "r", JoinKind.LeftOuter), "r",
             {"屋号数", "代表の薬", "直近の1日平均払出", "ずれの大きさ", "傾向", "年間比", "患者の傾向"}),

    Calc = Table.AddColumn(J2, "計算", each
        let
            s0     = [s] ?? {},
            s      = List.Repeat({0}, List.Max({0, 12 - List.Count(s0)})) & s0,
            n12    = List.Count(s0),
            recent = List.Sum(List.LastN(s, N)) ?? 0,
            before = if n12 >= 2 * N then List.Sum(List.FirstN(List.LastN(s, 2 * N), N)) else null,
            chg    = if before = null or before = 0 then null else (recent - before) / before,
            perDay = [直近の1日平均払出] ?? 0,
            rop    = Number.RoundUp(perDay * RopDays),
            oq     = Number.RoundUp(perDay * OqDays),
            turns  = perDay * 30 >= TinyOut,
            pred   = [ずれの大きさ] <> null and [ずれの大きさ] < GsLine,
            policy = if perDay <= 0 then "持たない（払出なし）"
                     else if turns and pred then "自動発注で持つ"
                     else if turns then "最低在庫を持つ"
                     else if pred then "持たずに取り寄せ"
                     else "持たない",
            changed = [前月の区分] <> null and [前月の区分] <> [区分],
            moved   = chg <> null and Number.Abs(chg) >= ChgLine,
            reason  = Text.Combine(List.RemoveNulls({
                          if changed then "区分 " & [前月の区分] & "→" & [区分] else null,
                          if moved then "直近" & Text.From(N) & "ヶ月の払出が前の" & Text.From(N) & "ヶ月より "
                                        & (if chg >= 0 then "+" else "") & Text.From(Number.Round(chg * 100, 0)) & "%" else null,
                          if [傾向] = "増加" or [傾向] = "減少" then
                              [傾向] & "傾向（年間の " & (if [年間比] = null then "－" else Text.From(Number.Round([年間比] * 100, 0)) & "%")
                              & "。患者は " & ([患者の傾向] ?? "不明") & "）" else null}), "、"),
            months  = Record.FromList(s, MonthCols)
        in
            [直近の払出 = recent, 前の払出 = before, 変化 = chg, 変化の大きさ = if chg = null then 0 else Number.Abs(chg),
             持ち方 = policy, 発注点の目安 = rop, 発注量の目安 = oq, 理由 = reason, 区分変更 = if changed then "○" else "", 対象 = changed or moved] & months),
    E    = Table.ExpandRecordColumn(Calc, "計算",
               {"直近の払出", "前の払出", "変化", "変化の大きさ", "持ち方", "発注点の目安", "発注量の目安", "理由", "区分変更", "対象"} & MonthCols),
    Sel  = Table.SelectRows(E, each [対象]),
    Per  = Table.AddColumn(Sel, "期間", (x) => Date.ToText(List.First(Last12), "yyyy/MM") & "～" & Date.ToText(Cur, "yyyy/MM"), type text),
    Out  = Table.SelectColumns(Per, {"需要のまとまり", "屋号数", "代表の薬", "前月の区分", "区分", "区分変更", "量の区分", "月平均払出", "月平均患者数",
               "直近の払出", "前の払出", "変化", "傾向", "年間比", "患者の傾向", "構成比", "持ち方", "発注点の目安", "発注量の目安", "理由", "期間"} & MonthCols & {"変化の大きさ"}),
    Ren  = Table.RenameColumns(Out, {{"区分", "今月の区分"}}),
    Typed = Table.TransformColumnTypes(Ren, {
               {"需要のまとまり", type text}, {"屋号数", Int64.Type}, {"代表の薬", type text}, {"前月の区分", type text},
               {"今月の区分", type text}, {"区分変更", type text}, {"量の区分", type text}, {"月平均払出", type number},
               {"月平均患者数", type number}, {"直近の払出", type number}, {"前の払出", type number},
               {"変化", Percentage.Type}, {"傾向", type text}, {"年間比", type number}, {"患者の傾向", type text},
               {"構成比", Percentage.Type}, {"持ち方", type text},
               {"発注点の目安", type number}, {"発注量の目安", type number}, {"理由", type text}, {"期間", type text},
               {"変化の大きさ", type number}} & List.Transform(MonthCols, (c) => {c, type number})),
    Sorted = Table.Sort(Typed, {{"区分変更", Order.Descending}, {"変化の大きさ", Order.Descending}, {"需要のまとまり", Order.Ascending}}),
    Result = Table.RemoveColumns(Sorted, {"変化の大きさ"})
in
    Result
