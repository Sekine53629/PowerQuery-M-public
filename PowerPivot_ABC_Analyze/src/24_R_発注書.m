// ================================================================================
// ㉔ クエリ名：R_発注書（シート「F_発注書」に表として読み込む）
//    F 発注書（来週の分）：薬ごとに、来週（予測開始～予測終了）の予定と在庫から発注の目安を出し、
//    発注してよいか（発注する／控えめに／発注しない／寄せる／不要）と注意を付ける。区分 → 略称 → 薬品名 の順に並ぶ
//    ・来週の予定・来週の患者：薬品使用量予測データ（設定「予測フォルダパス」。空欄なら実績フォルダ）の
//        予測開始（発注日の翌日）～予測終了（予測開始 ＋ 予測日数 − 1）の 予定使用量・患者数 の合計。
//        当月の予測データには月末までの予定が入っている。週が翌月にかかるときは翌月のファイルも置く
//    ・在庫：予測データの一番新しい日の 在庫数量（発注残を含む）。予測に出ない薬は M_薬品 の最終在庫（出納の最終月。参考）
//    ・ずれ：M_薬品 の ずれの割合（1 年の 払出 ÷ 予定 − 1）。見込み ＝ 来週の予定 × (1 ＋ ずれ)（0 未満にはしない）
//    ・不足 ＝ 来週の予定 − 在庫（マイナスは 0）。発注の目安 ＝ 見込み − 在庫 を切り上げ（錠。箱への切り上げは現場）
//    ・必要な最大在庫：1 年の日次の予定を「納品と納品の間」（注文は翌営業日に届くとして、同じ次の納品日を持つ日々）で
//        区切った合計のいちばん大きいもの × (1 ＋ ずれ)。連休でも不足しない在庫の上限の目安（区分 A の薬の管理用）。
//        卸の営業日は ④ fx_卸の休み（土日・祝日・年末年始・追加休業日）
//    【発注の収束】（0.9.0）同じ需要のまとまり（薬品グループ。屋号が 2 つ以上）の薬は、需要と在庫をまとまりで合算し、
//        発注を M_薬品 の「発注先」（切替先 → なければ通常供給で直近の払出が最多の屋号）の 1 薬に寄せる
//        ・まとまりの予定・まとまりの在庫：まとまり全体の 来週の予定・在庫 の合計（屋号ごとの値は 来週の予定・在庫 の列のまま）
//        ・見込み・不足・発注の目安・必要な最大在庫 は まとまりの合計 で計算し、発注先の行にだけ出す。ずれ は まとまりのずれの割合
//        ・発注先でない屋号は 判定「寄せる」（目安 0）。注意に「「発注先」に寄せる」と出る。在庫は使い切る
//        ・まとまりに登録のない薬、屋号が 1 つだけのまとまりは、これまでどおり薬ごとに計算する
//    ・判定：発注しない ＝ 発注除外／供給状況が 供給停止・薬価削除予定／切替元（薬品グループの 切替先 あり。0.7.0）（目安は 0）
//            寄せる     ＝ まとまりの発注先でない屋号（0.9.0）（目安は 0）
//            控えめに   ＝ 区分 D／経過措置品／供給状況 限定／予定より「ずれの注意割合」以上少なく出る薬／
//                         切替先（まとまりで合算していないとき）で、目安 > 0
//            発注する   ＝ それ以外で目安 > 0　　不要 ＝ 目安 0
//    ・注意：限定出荷／出荷調整の疑い（⑬ G01・G02）／期限の近いロットを先に使う（⑭ 残り日数 ≦ 期限切迫の日数 か 返す量 > 0）／
//            予定より n% 少なく出る／患者が少ない（区分 D）／経過措置品／予定にない薬／上限を超える（区分 A で 在庫＋目安 > 必要な最大在庫）／
//            まとまりの発注先（n 品目を合算）／「X」に寄せる
//    【傾向】（0.10.0）M_薬品 の 傾向（まとまりなら まとまりの傾向）が 増加／減少 の薬は、列「傾向」に出し、注意に
//        「減少傾向（直近 N ヶ月は前の N ヶ月の a%、年間の b%。患者は …）。目安を減らすか検討」のように書く。
//        発注の目安の計算は変えない（予定 × 年間のずれ のまま）。増減の判断と量の調整は使う人に委ねる
//    ・載せる薬：来週の予定か不足があるか（まとまりならその合計で見る）、注意（寄せる・傾向の注意を除く）があるか、
//        区分 A（上限の確認のため常に載せる）。在庫が足りて注意もない薬は載せない。まとまりは発注先と屋号の行が並ぶ
//    ・区分の列でオートフィルターをかければ、区分ごとに印刷できる
// ================================================================================
let
    Prm  = #"設定値",
    Num  = (key as text, dflt as number) as number =>
               let v = Record.FieldOrDefault(Prm, key, null)
               in  if v = null or (v is text and Text.Trim(v) = "") then dflt else Number.From(v),
    GapRate = Num("ずれの注意割合", 0.3),
    RecentN = Text.From(Number.RoundDown(Num("直近の月数", 3))),
    P       = (x as nullable number) as text => if x = null then "－" else Number.ToText(Number.Round(x * 100, 0), "N0", "ja-JP") & "%",
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

    // ---- 薬ごとの材料：まとまりキー（発注先のあるまとまりは需要のまとまり、それ以外は薬品名）と在庫 ----
    Drugs = #"M_薬品",
    K0 = Table.AddColumn(Drugs, "まとまりキー", each if [発注先] <> null then [需要のまとまり] else [薬品名], type text),
    J1 = Table.ExpandTableColumn(Table.NestedJoin(K0, {"薬品名"}, FcG, {"薬品名"}, "f", JoinKind.LeftOuter), "f",
             {"来週の予定", "来週の患者", "在庫_予測"}),
    J1s = Table.AddColumn(J1, "在庫_薬", each [在庫_予測] ?? [最終在庫] ?? 0, type number),
    Abbr = Table.Buffer(Table.SelectColumns(Drugs, {"薬品名", "略称"})),
    AbbrOf = (n as nullable text) as nullable text =>
                 if n = null then null else try Abbr{[薬品名 = n]}[略称] otherwise null,

    // ---- まとまりの合計（来週の予定・患者・在庫） ----
    Grp = Table.Group(J1s, {"まとまりキー"}, {
              {"まとまりの予定", each List.Sum(List.Transform([来週の予定], (v) => v ?? 0)), type number},
              {"まとまりの患者", each List.Sum(List.Transform([来週の患者], (v) => v ?? 0)), type number},
              {"まとまりの在庫", each List.Sum([在庫_薬]), type number}}),
    J1g = Table.ExpandTableColumn(Table.NestedJoin(J1s, {"まとまりキー"}, Grp, {"まとまりキー"}, "g", JoinKind.LeftOuter), "g",
              {"まとまりの予定", "まとまりの患者", "まとまりの在庫"}),

    // ---- 必要な最大在庫：納品と納品の間に出る予定の最大（1 年。まとまりは合算） ----
    S0     = Date.From(Prm[開始日]),
    E0     = Date.From(Prm[終了日]),
    Cal    = #"fx_卸の休み"(S0, Date.AddDays(E0, 14)),
    Biz    = List.Buffer(List.Sort(Table.SelectRows(Cal, each [卸の営業日])[日付])),
    NextBiz = (d as date) as nullable date => List.First(List.Select(Biz, each _ > d), null),
    Dates  = List.Dates(S0, Duration.Days(E0 - S0) + 1, #duration(1, 0, 0, 0)),
    Blocks = Table.Buffer(Table.FromColumns({Dates, List.Transform(Dates, NextBiz)},
                 type table [日付 = date, 納品日 = nullable date])),
    KeyMap = Table.Buffer(Table.SelectColumns(K0, {"薬品名", "まとまりキー"})),
    DayK   = Table.ExpandTableColumn(
                 Table.NestedJoin(#"T_日次", {"薬品名"}, KeyMap, {"薬品名"}, "k", JoinKind.LeftOuter), "k", {"まとまりキー"}),
    DayK2  = Table.ReplaceValue(DayK, null, each [薬品名], Replacer.ReplaceValue, {"まとまりキー"}),
    DayB   = Table.ExpandTableColumn(
                 Table.NestedJoin(DayK2, {"日付"}, Blocks, {"日付"}, "b", JoinKind.LeftOuter), "b", {"納品日"}),
    PerBlk = Table.Group(DayB, {"まとまりキー", "納品日"}, {
                 {"量", each List.Sum([予定数量]) ?? 0, type number},
                 {"始", each List.Min([日付]), type date},
                 {"終", each List.Max([日付]), type date}}),
    MaxBlk = Table.Group(PerBlk, {"まとまりキー"}, {{"r", each
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
    J2 = Table.ExpandTableColumn(Table.NestedJoin(J1g, {"まとまりキー"}, MaxE, {"まとまりキー"}, "m", JoinKind.LeftOuter), "m",
             {"最大の期間の量", "期間始", "期間終"}),
    Calc = Table.AddColumn(J2, "計算", (r) =>
        let
            pooled  = r[発注先] <> null,                       // まとまりで合算する薬（屋号が 2 つ以上）
            target  = not pooled or r[発注先] = r[薬品名],      // 発注の目安を出す行
            plan    = r[来週の予定] ?? 0,
            pats    = r[来週の患者] ?? 0,
            stock   = r[在庫_薬],
            gPlan   = if pooled then r[まとまりの予定] else plan,
            gStock  = if pooled then r[まとまりの在庫] else stock,
            gap     = (if pooled then r[まとまりのずれの割合] else r[ずれの割合]) ?? 0,
            factor  = List.Max({0, 1 + gap}),
            exp     = gPlan * factor,
            short   = List.Max({0, gPlan - gStock}),
            maxQ    = Number.RoundUp(Number.Round((r[最大の期間の量] ?? 0) * factor, 6)),
            cls     = r[区分],
            sup     = r[供給状況],
            excl    = r[発注除外] = true,
            noOrder = excl or sup = "供給停止" or sup = "薬価削除予定" or r[切替先] <> null,
            careful = cls = "D" or r[経過措置] = true or sup = "限定" or gap <= -GapRate or (r[切替元] <> null and not pooled),
            oq      = if noOrder or not target then 0 else Number.RoundUp(Number.Round(List.Max({0, exp - gStock}), 6)),
            overCap = target and cls = "A" and maxQ > 0 and gStock + oq > maxQ,
            grpTxt  = "予定 " & F(r[まとまりの予定]) & "・在庫 " & F(r[まとまりの在庫]) & " を合算",
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
                          if r[切替先] <> null then (if pooled then "切替元（在庫を使い切る）" else "切替元（「" & r[切替先] & "」へ寄せる。在庫を使い切る）") else null,
                          if r[切替元] <> null then (if pooled then "切替先（切替元の在庫は合算済み）"
                                                     else "切替先（切替元の在庫 " & F(r[切替元の在庫]) & " を使い切るまで控えめに）") else null,
                          if overCap then "上限を超える（在庫＋目安 " & F(gStock + oq) & " ＞ " & F(maxQ) & "）" else null}),
            trend   = if pooled then r[まとまりの傾向] else r[傾向],
            tr1     = if pooled then r[まとまりの前期比] else r[前期比],
            tr2     = if pooled then r[まとまりの年間比] else r[年間比],
            trendNotes = if trend = "増加" or trend = "減少" then
                             {trend & "傾向（直近 " & RecentN & " ヶ月は前の " & RecentN & " ヶ月の " & P(tr1) & "、年間の " & P(tr2)
                              & "。患者は " & (r[患者の傾向] ?? "不明") & "）。目安を" & (if trend = "増加" then "増やす" else "減らす") & "か検討"}
                         else {},
            poolNotes = if not pooled then {}
                        else if target then {"まとまり「" & r[需要のまとまり] & "」の発注先（" & Text.From(r[まとまりの品目数]) & " 品目の " & grpTxt & "）"}
                        else {"「" & r[発注先] & "」に寄せる（まとまり「" & r[需要のまとまり] & "」：" & grpTxt & "）"},
            noteOther = Text.Combine(notes, "、"),
            note    = Text.Combine(poolNotes & trendNotes & notes, "、"),
            verdict = if pooled and not target then "寄せる"
                      else if noOrder then "発注しない"
                      else if oq > 0 and careful then "控えめに"
                      else if oq > 0 then "発注する" else "不要",
            period  = if r[期間始] = null then ""
                      else Date.ToText(r[期間始], "yyyy/MM/dd") & "～" & Date.ToText(r[期間終], "yyyy/MM/dd")
                           & "（" & Text.From(Duration.Days(r[期間終] - r[期間始]) + 1) & "日）",
            order   = if cls = "A" then 1 else if cls = "B" then 2 else if cls = "C" then 3 else if cls = "D" then 4 else 5,
            sortAb  = (if pooled then AbbrOf(r[発注先]) else r[略称]) ?? ""
        in
            [来週の予定 = plan, 来週の患者 = pats, 在庫 = stock,
             需要のまとまり_表示 = if pooled then r[需要のまとまり] else null,
             まとまりの予定_表示 = if pooled then r[まとまりの予定] else null,
             まとまりの在庫_表示 = if pooled then r[まとまりの在庫] else null,
             不足 = if target then short else null, ずれ = gap,
             傾向 = if trend = "増加" or trend = "減少" then trend else null,
             見込み = if target then Number.Round(exp, 0) else null,
             発注の目安 = oq,
             必要な最大在庫 = if target then maxQ else null,
             最大の期間 = if target then period else "",
             判定 = verdict, 注意 = note,
             載せる = gPlan > 0 or short > 0 or noteOther <> "" or cls = "A",
             並び = order, 並び略 = sortAb, 並び先 = if target then 0 else 1]),
    // 計算の中の 来週の予定・来週の患者・傾向 は M_薬品／J1 の列と同じ名前なので、先に外してから広げる
    E    = Table.ExpandRecordColumn(Table.RemoveColumns(Calc, {"来週の予定", "来週の患者", "在庫_予測", "在庫_薬",
               "まとまりの予定", "まとまりの患者", "まとまりの在庫", "傾向"}), "計算",
               {"来週の予定", "来週の患者", "在庫", "需要のまとまり_表示", "まとまりの予定_表示", "まとまりの在庫_表示", "不足", "ずれ", "傾向", "見込み",
                "発注の目安", "必要な最大在庫", "最大の期間", "判定", "注意", "載せる", "並び", "並び略", "並び先"}),
    Sel  = Table.SelectRows(E, each [載せる]),
    Out  = Table.SelectColumns(Sel, {"区分", "略称", "薬品名", "来週の予定", "来週の患者", "在庫",
               "需要のまとまり_表示", "まとまりの予定_表示", "まとまりの在庫_表示", "不足", "ずれ", "傾向", "見込み",
               "発注の目安", "必要な最大在庫", "最大の期間", "判定", "注意", "まとまりキー", "並び", "並び略", "並び先"}),
    Ren  = Table.RenameColumns(Out, {{"需要のまとまり_表示", "需要のまとまり"}, {"まとまりの予定_表示", "まとまりの予定"},
               {"まとまりの在庫_表示", "まとまりの在庫"}}),
    Typed = Table.TransformColumnTypes(Ren, {
               {"区分", type text}, {"略称", type text}, {"薬品名", type text}, {"来週の予定", type number},
               {"来週の患者", type number}, {"在庫", type number}, {"需要のまとまり", type text},
               {"まとまりの予定", type number}, {"まとまりの在庫", type number},
               {"不足", type number}, {"ずれ", Percentage.Type}, {"傾向", type text},
               {"見込み", type number}, {"発注の目安", type number}, {"必要な最大在庫", type number},
               {"最大の期間", type text}, {"判定", type text}, {"注意", type text},
               {"まとまりキー", type text}, {"並び", Int64.Type}, {"並び略", type text}, {"並び先", Int64.Type}}),
    Sorted = Table.Sort(Typed, {{"並び", Order.Ascending}, {"並び略", Order.Ascending}, {"まとまりキー", Order.Ascending},
               {"並び先", Order.Ascending}, {"薬品名", Order.Ascending}}),
    Result = Table.RemoveColumns(Sorted, {"まとまりキー", "並び", "並び略", "並び先"})
in
    Result
