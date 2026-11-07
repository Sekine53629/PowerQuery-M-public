// ================================================================================
// ⑫ クエリ名：T_ABC月次（データモデルに読み込む）
//    月×需要のまとまり の区分（A／B／C／D）。その月を終わりとする直近「ABCの月数」ヶ月で付ける（0.5.0）
//    ・まとまり：薬品グループの「需要のまとまり」（手入力テーブルを直接読む）。未登録の薬は薬品名そのもの
//    ・行：その月までに一度でも出納か予定に出たまとまり（窓に払出がなくても 払出 0 の行を作る。患者も 0 なら D）
//    ・月平均払出：窓（直近 N ヶ月）の払出の合計 ÷ 窓の月数
//    ・量の区分：月平均払出を固定の線で切る
//        A：≧ 区分Aの線（初期値 1,000 錠／月。需要が多すぎて一度に全量を発注できない薬。在庫の上限を人が管理する）
//        B：≧ 区分Bの線（初期値 300。通常の発注量で問題がない薬）
//        C：それ未満（0 も C。定期処方はあるが在庫量は多くない薬）
//    ・患者の最少月：窓の各月の、まとまりの予定の延べ患者数（T_月次 の予定患者数を合計）のいちばん少ない月の値。
//        予定データのある月だけで見る。予定の行がない月は 0 人。窓に予定データの月がなければ空欄
//    ・月平均患者数：同じ月々の平均（参考）
//    ・区分：患者の最少月 ＜ 区分Dの患者数（初期値 2）なら D（患者が極端に少なく、在庫を手で管理する薬。
//        DeadStock の D。出納にしか出ない薬もここに入る）。患者の最少月が空欄（予定データが窓にない）なら 量の区分 のまま。
//        それ以外は 量の区分。D は量より優先する（「D なのに量は A」＝まとまって出る薬が 量の区分 で分かる）
//    ・構成比・累積：窓の払出の構成比と累積（参考。0.4.0 までの区分の根拠。区分には使わない）
//    ・前月の区分：1つ前の出納の月の区分（暦の前月ではなく、出納のある月の並びで1つ前）
//    ・変更：「B→A」のように表示。変わっていなければ空欄
//    ・線は設定テーブルで変える（ABCの月数 3／区分Aの線 1000／区分Bの線 300／区分Dの患者数 2）
// ================================================================================
let
    Prm  = #"設定値",
    Num  = (key as text, dflt as number) as number =>
               let v = Record.FieldOrDefault(Prm, key, null)
               in  if v = null or (v is text and Text.Trim(v) = "") then dflt else Number.From(v),
    N     = Number.RoundDown(Num("ABCの月数", 3)),
    ALine = Num("区分Aの線", 1000),
    BLine = Num("区分Bの線", 300),
    DLine = Num("区分Dの患者数", 2),

    // ---- まとまり（薬品グループの手入力テーブル。M_薬品 を通さない：⑨ がこのクエリを参照するため） ----
    Clean = (v as any) as nullable text => if v = null then null else let t = Text.Trim(Text.From(v)) in if t = "" then null else t,
    Hand  = (name as text, cols as list) as table =>
                try Table.SelectColumns(Excel.CurrentWorkbook(){[Name = name]}[Content], cols, MissingField.UseNull)
                otherwise #table(cols, {}),
    Grp   = Table.Buffer(Table.Distinct(Table.TransformColumns(Hand("薬品グループ", {"薬品名", "需要のまとまり"}), {
                {"薬品名", Clean, type nullable text}, {"需要のまとまり", Clean, type nullable text}}), {"薬品名"})),

    MonAll = #"T_月次",
    Keyed  = Table.AddColumn(
                 Table.ExpandTableColumn(
                     Table.NestedJoin(MonAll, {"薬品名"}, Grp, {"薬品名"}, "g", JoinKind.LeftOuter), "g", {"需要のまとまり"}),
                 "まとまり", each [需要のまとまり] ?? [薬品名], type text),
    // 払出：出納のある月の行（予定だけの薬も払出 0 で入っている）
    KeyedB = Table.Buffer(Table.SelectColumns(Table.SelectRows(Keyed, each [出納データあり]), {"月", "まとまり", "払出数量"})),
    Months = List.Buffer(List.Sort(List.Distinct(KeyedB[月]))),
    // 予定の延べ患者数：予定データのある月の、月×まとまりの合計
    PatB   = Table.Buffer(Table.Group(Table.SelectRows(Keyed, each [予定データあり]), {"月", "まとまり"},
                 {{"患者", each List.Sum([予定患者数]) ?? 0, type number}})),
    PlanMonths = List.Buffer(List.Distinct(PatB[月])),

    // ---- 1つの月の区分表（その月を終わりとする N ヶ月） ----
    OneMonth = (m as date) as table =>
        let
            i      = List.PositionOf(Months, m),
            win    = List.LastN(List.FirstN(Months, i + 1), N),
            nWin   = List.Count(win),
            rows   = Table.SelectRows(KeyedB, each List.Contains(win, [月])),
            sums0  = Table.Group(rows, {"まとまり"}, {{"払出", each List.Sum([払出数量]) ?? 0, type number}}),
            // その月までに出たまとまり（窓に行がなければ払出 0）
            seen   = Table.FromColumns({List.Distinct(
                         Table.SelectRows(KeyedB, each [月] <= m)[まとまり] & Table.SelectRows(PatB, each [月] <= m)[まとまり])},
                         type table [まとまり = text]),
            sums   = Table.TransformColumns(
                         Table.ExpandTableColumn(Table.NestedJoin(seen, {"まとまり"}, sums0, {"まとまり"}, "s", JoinKind.LeftOuter), "s", {"払出"}),
                         {{"払出", each _ ?? 0, type number}}),
            sorted = Table.Sort(sums, {{"払出", Order.Descending}, {"まとまり", Order.Ascending}}),
            total  = List.Sum(sorted[払出]),
            outs   = List.Buffer(sorted[払出]),
            idx    = Table.AddIndexColumn(sorted, "i", 0, 1),
            cum    = Table.AddColumn(idx, "累積", each
                         if total = null or total = 0 then 0 else List.Sum(List.FirstN(outs, [i] + 1)) / total, type number),
            share  = Table.AddColumn(cum, "構成比", each if total = null or total = 0 then 0 else [払出] / total, type number),
            avg    = Table.AddColumn(share, "月平均払出", each if nWin > 0 then [払出] / nWin else 0, type number),
            qcls   = Table.AddColumn(avg, "量の区分", each
                         if [月平均払出] >= ALine then "A" else if [月平均払出] >= BLine then "B" else "C", type text),
            // 患者：窓のうち予定データのある月だけ。予定の行がない月は 0 人
            winP   = List.Select(win, each List.Contains(PlanMonths, _)),
            nP     = List.Count(winP),
            pats   = Table.Group(Table.SelectRows(PatB, each List.Contains(winP, [月])), {"まとまり"}, {
                         {"最少", each List.Min([患者]), type number},
                         {"合計", each List.Sum([患者]), type number},
                         {"n",    each Table.RowCount(_), Int64.Type}}),
            jp     = Table.ExpandTableColumn(
                         Table.NestedJoin(qcls, {"まとまり"}, pats, {"まとまり"}, "p", JoinKind.LeftOuter), "p", {"最少", "合計", "n"}),
            pmin   = Table.AddColumn(jp, "患者の最少月", each
                         if nP = 0 then null
                         else if [n] = null or [n] < nP then 0
                         else [最少], type nullable number),
            pavg   = Table.AddColumn(pmin, "月平均患者数", each
                         if nP = 0 then null else ([合計] ?? 0) / nP, type nullable number),
            cls    = Table.AddColumn(pavg, "区分", each
                         if [患者の最少月] <> null and [患者の最少月] < DLine then "D" else [量の区分], type text),
            withM  = Table.AddColumn(cls, "月", (x) => m, type date)
        in
            Table.SelectColumns(withM, {"月", "まとまり", "払出", "構成比", "累積", "月平均払出", "量の区分",
                                        "患者の最少月", "月平均患者数", "区分"}),

    EmptyT = #table(type table [月 = date, まとまり = text, 払出 = number, 構成比 = number, 累積 = number,
                                月平均払出 = number, 量の区分 = text, 患者の最少月 = nullable number,
                                月平均患者数 = nullable number, 区分 = text], {}),
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
                 {"構成比", Percentage.Type}, {"累積", Percentage.Type}, {"月平均払出", type number},
                 {"量の区分", type text}, {"患者の最少月", type number}, {"月平均患者数", type number},
                 {"区分", type text}, {"前月の区分", type text}, {"変更", type text}}),
    Result = Table.Sort(Typed, {{"月", Order.Ascending}, {"累積", Order.Ascending}})
in
    Result
