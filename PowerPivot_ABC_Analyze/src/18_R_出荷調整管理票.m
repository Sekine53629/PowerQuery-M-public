// ================================================================================
// ⑱ クエリ名：R_出荷調整管理票（シート「D_出荷調整」に表として読み込む）
//    D 出荷調整の管理票：屋号が2つ以上ある「需要のまとまり」ごとに、役割を付けて並べる
//    ・役割：経過措置品（期限あり、在庫か直近の払出あり）を先に外し、残りの後発を直近の払出の多い順に
//        推奨ジェネリック・次点ジェネリック、先発品、その他の後発。最後に 経過措置（管理継続）
//        供給状況が「通常」以外の屋号は、推奨・次点の順位を後ろに回す
//    ・状態：⑬ T_まとまり月次 の G01（構成比の入れ替わり）・G02（入荷停止の疑い）の一番新しい月、
//        経過措置品は期限と残り日数
//    ・切替・転送の案（まとまりごとに同じ文）：
//        推奨に G02 → 次点（なければ先発）へ切替／推奨の在庫 < 発注点の日数 → 次点を併用／
//        G01 で構成比が落ちた屋号に在庫が残る → 転送・返品／
//        経過措置品がまとまりで一番出ている → 【警告】推奨を置き換え（E04）／
//        経過措置品の在庫 → 期限まで使い切れるか・余る量。在庫 0 でも期限まで管理を継続
//    ・疑い：「経過措置」「疑い」「」。先頭に経過措置、次に疑いのあるまとまり、あとは直近の払出の多い順
//    ・構成比_月01～月12：直近12ヶ月（出納のある月）の構成比。月12 が一番新しい月
// ================================================================================
let
    Prm  = #"設定値",
    Num  = (key as text, dflt as number) as number =>
               let v = Record.FieldOrDefault(Prm, key, null)
               in  if v = null or (v is text and Text.Trim(v) = "") then dflt else Number.From(v),
    RopDays  = Num("発注点の日数", 7),
    BaseDate = Prm[発注日],
    F        = (n as nullable number) as text => if n = null then "" else Number.ToText(Number.Round(n, 0), "N0", "ja-JP"),
    MonthCols = List.Transform({1 .. 12}, (i) => "構成比_月" & Text.PadStart(Text.From(i), 2, "0")),

    Drugs  = Table.Buffer(Table.SelectRows(#"M_薬品", each [需要のまとまり] <> null)),
    GM     = Table.Buffer(#"T_まとまり月次"),
    Months = List.Buffer(List.Sort(List.Distinct(GM[月]))),
    Last12 = List.LastN(Months, 12),
    Cur    = if List.IsEmpty(Months) then null else List.Last(Months),

    // ---- 薬ごと：一番新しい G01／G02 と、12ヶ月の構成比、先月の仕入 ----
    FlagT  = Table.Group(Table.SelectRows(GM, each [G01] = true or [G02] = true), {"薬品名"},
                 {{"f", each Table.Last(Table.Sort(_, {{"月", Order.Ascending}})), type record}}),
    FlagE  = Table.ExpandRecordColumn(FlagT, "f", {"月", "G01", "G02", "前月の構成比", "構成比", "変化"},
                 {"疑いの月", "G01", "G02", "疑いの前月の構成比", "疑いの当月の構成比", "疑いの変化"}),
    ShareT = Table.Group(Table.SelectRows(GM, each List.Contains(Last12, [月])), {"薬品名"}, {{"s", each
                 let t = _
                 in  List.Transform(Last12, (m) =>
                         let r = Table.SelectRows(t, each [月] = m)
                         in  if Table.IsEmpty(r) then 0 else r{0}[構成比]), type list}}),
    LastBuy = Table.RenameColumns(
                  Table.SelectColumns(Table.SelectRows(#"T_月次", each [月] = Cur), {"薬品名", "仕入数量"}), {{"仕入数量", "先月の仕入"}}),
    D1 = Table.ExpandTableColumn(Table.NestedJoin(Drugs, {"薬品名"}, FlagE, {"薬品名"}, "f", JoinKind.LeftOuter), "f",
             {"疑いの月", "G01", "G02", "疑いの前月の構成比", "疑いの当月の構成比", "疑いの変化"}),
    D2 = Table.ExpandTableColumn(Table.NestedJoin(D1, {"薬品名"}, ShareT, {"薬品名"}, "sj", JoinKind.LeftOuter), "sj", {"s"}),
    D3 = Table.ExpandTableColumn(Table.NestedJoin(D2, {"薬品名"}, LastBuy, {"薬品名"}, "b", JoinKind.LeftOuter), "b", {"先月の仕入"}),

    // ---- まとまりごとに役割を付け、案を作る ----
    AssignRoles = (g as table) as table =>
        let
            trans   = Table.SelectRows(g, each [経過措置期限] <> null and (([最終在庫] ?? 0) > 0 or ([直近の払出合計] ?? 0) > 0)),
            transN  = trans[薬品名],
            rest    = Table.AddColumn(Table.SelectRows(g, each not List.Contains(transN, [薬品名])), "k",
                          each if [供給状況] = null or [供給状況] = "通常" then 0 else 1, Int64.Type),
            gen     = Table.ToRecords(Table.Sort(Table.SelectRows(rest, each [先発後発] = "後発"),
                          {{"k", Order.Ascending}, {"直近の払出合計", Order.Descending}})),
            brand   = Table.ToRecords(Table.Sort(Table.SelectRows(rest, each [先発後発] <> "後発"),
                          {{"k", Order.Ascending}, {"直近の払出合計", Order.Descending}})),
            transL  = Table.ToRecords(Table.Sort(trans, {{"経過措置期限", Order.Ascending}})),
            roles   = (if List.Count(gen) >= 1 then {[役割 = "推奨ジェネリック", r = gen{0}]} else {})
                    & (if List.Count(gen) >= 2 then {[役割 = "次点ジェネリック", r = gen{1}]} else {})
                    & List.Transform(brand, (x) => [役割 = "先発品", r = x])
                    & List.Transform(List.Skip(gen, 2), (x) => [役割 = "その他の後発", r = x])
                    & List.Transform(transL, (x) => [役割 = "経過措置（管理継続）", r = x]),
            top     = Table.First(Table.Sort(g, {{"直近の払出合計", Order.Descending}}))[薬品名],
            rec     = if List.Count(gen) >= 1 then gen{0} else null,
            nxt     = if List.Count(gen) >= 2 then gen{1} else null,
            newRec  = if rec <> null then rec[薬品名] else if List.Count(brand) >= 1 then brand{0}[薬品名] else null,
            tot     = List.Sum(g[直近の払出合計]) ?? 0,
            // 切替・転送の案
            a1 = if rec = null then {}
                 else if rec[G02] = true then
                     {"推奨「" & rec[薬品名] & "」の入荷が止まっている疑い → "
                      & (if nxt <> null then "次点「" & nxt[薬品名] & "」へ切替" else "先発へ切替")}
                 else if rec[直近の在庫日数] <> null and rec[直近の在庫日数] < RopDays then
                     {"推奨の在庫が " & F(rec[直近の在庫日数]) & " 日分と薄い → 次点を併用"}
                 else {},
            a2 = List.Transform(
                     List.Select(roles, (x) => x[r][G01] = true and (x[r][疑いの変化] ?? 0) < 0 and (x[r][最終在庫] ?? 0) > 0),
                     (x) => "「" & x[r][薬品名] & "」は " & Date.ToText(x[r][疑いの月], "yyyy/MM") & " に構成比 "
                            & F((x[r][疑いの前月の構成比] ?? 0) * 100) & "%→" & F((x[r][疑いの当月の構成比] ?? 0) * 100)
                            & "%。残った在庫 " & F(x[r][最終在庫]) & " は転送・返品を検討"),
            a3 = List.Transform(transL, (x) =>
                     let
                         left   = Duration.Days(x[経過措置期限] - BaseDate),
                         stock  = x[最終在庫] ?? 0,
                         perDay = x[直近の1日平均払出] ?? 0,
                         usable = perDay * List.Max({0, left})
                     in
                         if x[薬品名] = top then
                             "【警告】経過措置「" & x[薬品名] & "」がまだ一番出ています（期限 " & Date.ToText(x[経過措置期限], "yyyy/MM/dd")
                             & "、残り " & Text.From(left) & " 日）。"
                             & (if newRec <> null then "推奨を「" & newRec & "」に置き換え、処方の切替を医師に依頼" else "代替の屋号がありません。登録を")
                         else if stock > 0 then
                             "経過措置「" & x[薬品名] & "」の在庫 " & F(stock) & " は "
                             & (if usable >= stock then "期限まで使い切れます（" & F(usable) & "）" else F(stock - usable) & " が余る見込み。返品・転送を")
                             & "。管理を継続"
                         else
                             "経過措置「" & x[薬品名] & "」は在庫 0。期限 " & Date.ToText(x[経過措置期限], "yyyy/MM/dd") & " まで管理を継続"),
            adviceL = a1 & a2 & a3,
            advice  = if List.IsEmpty(adviceL) then "切替の必要なし" else Text.Combine(adviceL, "#(lf)"),
            doubt   = List.AnyTrue(List.Transform(Table.ToRecords(g), (x) => x[G01] = true or x[G02] = true)),
            tag     = if List.Count(transL) > 0 then "経過措置" else if doubt then "疑い" else "",
            tagOrd  = if List.Count(transL) > 0 then 0 else if doubt then 1 else 2,
            out     = Table.FromRecords(List.Transform(List.Positions(roles), (i) =>
                          let
                              x  = roles{i},
                              r  = x[r],
                              s0 = r[s] ?? {},
                              s  = List.Repeat({0}, List.Max({0, 12 - List.Count(s0)})) & s0,
                              st = if x[役割] = "経過措置（管理継続）" then
                                       "経過措置 期限 " & Date.ToText(r[経過措置期限], "yyyy/MM/dd") & "（残り "
                                       & Text.From(Duration.Days(r[経過措置期限] - BaseDate)) & " 日）"
                                   else if r[G01] = true then "構成比の入れ替わり（" & Date.ToText(r[疑いの月], "yyyy/MM") & "）"
                                   else if r[G02] = true then "入荷停止の疑い（" & Date.ToText(r[疑いの月], "yyyy/MM") & "）"
                                   else ""
                          in
                              [行順 = i + 1, 役割 = x[役割], 薬品名 = r[薬品名], 先発後発 = r[先発後発], 供給状況 = r[供給状況],
                               直近の払出 = r[直近の払出合計] ?? 0,
                               構成比 = if tot = 0 then 0 else (r[直近の払出合計] ?? 0) / tot,
                               在庫 = r[最終在庫], 直近の在庫日数 = r[直近の在庫日数], 先月の仕入 = r[先月の仕入],
                               状態 = st, 疑い = tag, 疑いの順 = tagOrd, #"切替・転送の案" = advice, まとまりの払出 = tot]
                              & Record.FromList(s, MonthCols)))
        in
            out,

    Groups = Table.Group(D3, {"需要のまとまり"}, {{"n", each Table.RowCount(_), Int64.Type}, {"t", each AssignRoles(_), type table}}),
    Multi  = Table.SelectRows(Groups, each [n] >= 2),
    Cols   = {"行順", "役割", "薬品名", "先発後発", "供給状況", "直近の払出", "構成比", "在庫", "直近の在庫日数", "先月の仕入",
              "状態", "疑い", "疑いの順", "切替・転送の案", "まとまりの払出"} & MonthCols,
    E      = Table.ExpandTableColumn(Table.SelectColumns(Multi, {"需要のまとまり", "t"}), "t", Cols),
    Typed  = Table.TransformColumnTypes(E, {
                 {"需要のまとまり", type text}, {"行順", Int64.Type}, {"役割", type text}, {"薬品名", type text},
                 {"先発後発", type text}, {"供給状況", type text}, {"直近の払出", type number}, {"構成比", Percentage.Type},
                 {"在庫", type number}, {"直近の在庫日数", type number}, {"先月の仕入", type number}, {"状態", type text},
                 {"疑い", type text}, {"疑いの順", Int64.Type}, {"切替・転送の案", type text}, {"まとまりの払出", type number}}
                 & List.Transform(MonthCols, (c) => {c, Percentage.Type})),
    Sorted = Table.Sort(Typed, {{"疑いの順", Order.Ascending}, {"まとまりの払出", Order.Descending},
                 {"需要のまとまり", Order.Ascending}, {"行順", Order.Ascending}}),
    Result = Table.RemoveColumns(Sorted, {"疑いの順"})
in
    Result
