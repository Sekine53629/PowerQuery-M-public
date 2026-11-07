// ================================================================================
// ㉔ クエリ名：R_発注書（シート「F_発注書」に表として読み込む）
//    F 発注書（来週の分）：薬ごとに、来週（予測開始～予測終了）の予定と在庫から発注の目安を出し、
//    発注してよいか（発注する／控えめに／発注しない／不要）と注意を付ける。区分 → 略称 → 薬品名 の順に並ぶ
//    ・来週の予定・来週の患者：薬品使用量予測データ（設定「予測フォルダパス」。空欄なら実績フォルダ）の
//        予測開始（発注日の翌日）～予測終了（予測開始 ＋ 予測日数 − 1）の 予定使用量・患者数 の合計。
//        当月の予測データには月末までの予定が入っている。週が翌月にかかるときは翌月のファイルも置く
//    ・在庫：予測データの一番新しい日の 在庫数量（発注残を含む）。予測に出ない薬は M_薬品 の最終在庫（出納の最終月。参考）
//    ・ずれ：M_薬品 の ずれの割合（1 年の 払出 ÷ 予定 − 1）。見込み ＝ 来週の予定 × (1 ＋ ずれ)（0 未満にはしない）
//    ・不足 ＝ 来週の予定 − 在庫（マイナスは 0）。発注の目安 ＝ 見込み − 在庫 を切り上げ（錠。箱への切り上げは現場）
//    ・必要な最大在庫：1 年の日次の予定を「納品と納品の間」（注文は翌営業日に届くとして、同じ次の納品日を持つ日々）で
//        区切った合計のいちばん大きいもの × (1 ＋ ずれ)。連休でも不足しない在庫の上限の目安（区分 A の薬の管理用）。
//        卸の営業日は ④ fx_卸の休み（土日・祝日・年末年始・追加休業日）
//    ・判定：発注しない ＝ 発注除外／供給状況が 供給停止・薬価削除予定／切替元（薬品グループの 切替先 あり。0.7.0）（目安は 0）
//            控えめに   ＝ 区分 D／経過措置品／供給状況 限定／予定より「ずれの注意割合」以上少なく出る薬／切替先で、目安 > 0
//            発注する   ＝ それ以外で目安 > 0　　不要 ＝ 目安 0
//    ・注意：限定出荷／出荷調整の疑い（⑬ G01・G02）／期限の近いロットを先に使う（⑭ 残り日数 ≦ 期限切迫の日数 か 返す量 > 0）／
//            予定より n% 少なく出る／患者が少ない（区分 D）／経過措置品／予定にない薬／上限を超える（区分 A で 在庫＋目安 > 必要な最大在庫）
//    ・載せる薬：来週の予定か不足があるか、注意があるか、区分 A（上限の確認のため常に載せる）。在庫が足りて注意もない薬は載せない
//    ・区分の列でオートフィルターをかければ、区分ごとに印刷できる
// ================================================================================
let
    Prm  = #"設定値",
    Num  = (key as text, dflt as number) as number =>
               let v = Record.FieldOrDefault(Prm, key, null)
               in  if v = null or (v is text and Text.Trim(v) = "") then dflt else Number.From(v),
    GapRate = Num("ずれの注意割合", 0.3),
    Near    = Num("期限切迫の日数", 180),
    F       = (n as nullable number) as text => if n = null then "" else Number.ToText(Number.Round(n, 0), "N0", "ja-JP"),

    // ---- 来週の予定（予測データ） ----
    FcFolder = Prm[予測フォルダパス] ?? Prm[フォルダパス],
    FcStart  = Date.From(Prm[予測開始]),
    FcEnd    = Date.From(Prm[予測終了]),
    Fc       = if FcFolder = null then error "設定テーブルの「フォルダパス」が空欄です"
               else #"fx_使用量読込"(Text.From(FcFolder), Prm[先頭], FcStart, FcEnd, Prm[列名_予測], true),
    FcG      = Table.Group(Fc, {"薬品名"}, {
                   {"来週の予定", each List.Sum([数量]) ?? 0, type number},
                   {"来週の患者", each List.Sum([患者数]) ?? 0, type number},
                   {"在庫_予測",  each let r = Table.Max(_, "日付") in r[在庫量_発注残込み] ?? r[在庫量], type nullable number}}),

    // ---- 必要な最大在庫：納品と納品の間に出る予定の最大（1 年） ----
    S0     = Date.From(Prm[開始日]),
    E0     = Date.From(Prm[終了日]),
    Cal    = #"fx_卸の休み"(S0, Date.AddDays(E0, 14)),
    Biz    = List.Buffer(List.Sort(Table.SelectRows(Cal, each [卸の営業日])[日付])),
    NextBiz = (d as date) as nullable date => List.First(List.Select(Biz, each _ > d), null),
    Dates  = List.Dates(S0, Duration.Days(E0 - S0) + 1, #duration(1, 0, 0, 0)),
    Blocks = Table.Buffer(Table.FromColumns({Dates, List.Transform(Dates, NextBiz)},
                 type table [日付 = date, 納品日 = nullable date])),
    DayB   = Table.ExpandTableColumn(
                 Table.NestedJoin(#"T_日次", {"日付"}, Blocks, {"日付"}, "b", JoinKind.LeftOuter), "b", {"納品日"}),
    PerBlk = Table.Group(DayB, {"薬品名", "納品日"}, {
                 {"量", each List.Sum([予定数量]) ?? 0, type number},
                 {"始", each List.Min([日付]), type date},
                 {"終", each List.Max([日付]), type date}}),
    MaxBlk = Table.Group(PerBlk, {"薬品名"}, {{"r", each
                 Table.First(Table.Sort(_, {{"量", Order.Descending}, {"始", Order.Ascending}})), type record}}),
    MaxE   = Table.ExpandRecordColumn(MaxBlk, "r", {"量", "始", "終"}, {"最大の期間の量", "期間始", "期間終"}),

    // ---- 注意の材料 ----
    Adj      = #"T_まとまり月次",
    AdjLast  = if Table.IsEmpty(Adj) then Adj else Table.SelectRows(Adj, each [月] = List.Max(Adj[月])),
    AdjNames = List.Buffer(Table.SelectRows(AdjLast, each [G01] = true or [G02] = true)[薬品名]),
    Lots     = #"T_ロット",
    LotNames = List.Buffer(List.Distinct(Table.SelectRows(Lots, each
                   ([残り日数] <> null and [残り日数] <= Near) or ([返す量] ?? 0) > 0)[薬品名])),

    // ---- 薬ごとの計算 ----
    Drugs = #"M_薬品",
    J1 = Table.ExpandTableColumn(Table.NestedJoin(Drugs, {"薬品名"}, FcG, {"薬品名"}, "f", JoinKind.LeftOuter), "f",
             {"来週の予定", "来週の患者", "在庫_予測"}),
    J2 = Table.ExpandTableColumn(Table.NestedJoin(J1, {"薬品名"}, MaxE, {"薬品名"}, "m", JoinKind.LeftOuter), "m",
             {"最大の期間の量", "期間始", "期間終"}),
    Calc = Table.AddColumn(J2, "計算", (r) =>
        let
            plan    = r[来週の予定] ?? 0,
            pats    = r[来週の患者] ?? 0,
            stock   = r[在庫_予測] ?? r[最終在庫] ?? 0,
            gap     = r[ずれの割合] ?? 0,
            factor  = List.Max({0, 1 + gap}),
            exp     = plan * factor,
            short   = List.Max({0, plan - stock}),
            maxQ    = Number.RoundUp(Number.Round((r[最大の期間の量] ?? 0) * factor, 6)),
            cls     = r[区分],
            sup     = r[供給状況],
            excl    = r[発注除外] = true,
            noOrder = excl or sup = "供給停止" or sup = "薬価削除予定" or r[切替先] <> null,
            careful = cls = "D" or r[経過措置] = true or sup = "限定" or gap <= -GapRate or r[切替元] <> null,
            oq      = if noOrder then 0 else Number.RoundUp(Number.Round(List.Max({0, exp - stock}), 6)),
            overCap = cls = "A" and maxQ > 0 and stock + oq > maxQ,
            notes   = List.RemoveNulls({
                          if excl then "発注除外" else null,
                          if sup = "供給停止" or sup = "薬価削除予定" then sup else null,
                          if sup = "限定" then "限定出荷" else null,
                          if List.Contains(AdjNames, r[薬品名]) then "出荷調整の疑い" else null,
                          if List.Contains(LotNames, r[薬品名]) then "期限の近いロットを先に使う" else null,
                          if gap <= -GapRate then "予定より " & F(-gap * 100) & "% 少なく出る" else null,
                          if cls = "D" then "患者が少ない（区分 D）" else null,
                          if r[経過措置] = true then "経過措置品（切替先があれば発注しない）" else null,
                          if r[出どころ] = "出納だけ" then "予定にない薬" else null,
                          if r[切替先] <> null then "切替元（「" & r[切替先] & "」へ寄せる。在庫を使い切る）" else null,
                          if r[切替元] <> null then "切替先（切替元の在庫 " & F(r[切替元の在庫]) & " を使い切るまで控えめに）" else null,
                          if overCap then "上限を超える（在庫＋目安 " & F(stock + oq) & " ＞ " & F(maxQ) & "）" else null}),
            note    = Text.Combine(notes, "、"),
            verdict = if noOrder then "発注しない"
                      else if oq > 0 and careful then "控えめに"
                      else if oq > 0 then "発注する" else "不要",
            period  = if r[期間始] = null then ""
                      else Date.ToText(r[期間始], "yyyy/MM/dd") & "～" & Date.ToText(r[期間終], "yyyy/MM/dd")
                           & "（" & Text.From(Duration.Days(r[期間終] - r[期間始]) + 1) & "日）",
            order   = if cls = "A" then 1 else if cls = "B" then 2 else if cls = "C" then 3 else if cls = "D" then 4 else 5
        in
            [来週の予定 = plan, 来週の患者 = pats, 在庫 = stock, 不足 = short, ずれ = gap, 見込み = Number.Round(exp, 0),
             発注の目安 = oq, 必要な最大在庫 = maxQ, 最大の期間 = period, 判定 = verdict, 注意 = note,
             載せる = plan > 0 or short > 0 or note <> "" or cls = "A", 並び = order]),
    // 計算の中の 来週の予定・来週の患者 は J1 の列と同じ名前なので、先に外してから広げる
    E    = Table.ExpandRecordColumn(Table.RemoveColumns(Calc, {"来週の予定", "来週の患者", "在庫_予測"}), "計算",
               {"来週の予定", "来週の患者", "在庫", "不足", "ずれ", "見込み", "発注の目安", "必要な最大在庫", "最大の期間", "判定", "注意", "載せる", "並び"}),
    Sel  = Table.SelectRows(E, each [載せる]),
    Out  = Table.SelectColumns(Sel, {"区分", "略称", "薬品名", "来週の予定", "来週の患者", "在庫", "不足", "ずれ", "見込み",
               "発注の目安", "必要な最大在庫", "最大の期間", "判定", "注意", "並び"}),
    Typed = Table.TransformColumnTypes(Out, {
               {"区分", type text}, {"略称", type text}, {"薬品名", type text}, {"来週の予定", type number},
               {"来週の患者", type number}, {"在庫", type number}, {"不足", type number}, {"ずれ", Percentage.Type},
               {"見込み", type number}, {"発注の目安", type number}, {"必要な最大在庫", type number},
               {"最大の期間", type text}, {"判定", type text}, {"注意", type text}, {"並び", Int64.Type}}),
    Sorted = Table.Sort(Typed, {{"並び", Order.Ascending}, {"略称", Order.Ascending}, {"薬品名", Order.Ascending}}),
    Result = Table.RemoveColumns(Sorted, {"並び"})
in
    Result
