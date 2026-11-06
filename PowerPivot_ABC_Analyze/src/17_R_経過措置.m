// ================================================================================
// ⑰ クエリ名：R_経過措置（シート「C_経過措置」に表として読み込む）
//    C 経過措置・期限切迫品：経過措置期限のある薬（経過措置リスト）について、
//    期限までに直近の払出ペースで在庫を使い切れるかを見る
//    ・残り日数 ＝ 経過措置期限 － 基準日（設定の発注日）。期限切迫 ＝ 残り日数 ≦ 設定「期限切迫の日数」
//    ・期限までに使い切れる量 ＝ 1日平均払出 × 残り日数（0 未満は 0）。余る量 ＝ 在庫 － 使い切れる量
//    ・代替の屋号：同じ「需要のまとまり」で経過措置のない薬（後発を優先、直近の払出の多い順）
//    ・処置の案：期限切れ → 返品か廃棄、処方の切替を確認／余る → 返品・転送と発注停止／
//        使い切れる → 発注量を残り日数分に抑える。代替があれば「在庫がなくなるまで管理を継続」
//    ※ ロットの使用期限は ⑭ T_ロット（期限切迫リスト）で見る
// ================================================================================
let
    Prm  = #"設定値",
    Num  = (key as text, dflt as number) as number =>
               let v = Record.FieldOrDefault(Prm, key, null)
               in  if v = null or (v is text and Text.Trim(v) = "") then dflt else Number.From(v),
    Near     = Num("期限切迫の日数", 180),
    BaseDate = Prm[発注日],
    F        = (n as nullable number) as text => if n = null then "" else Number.ToText(Number.Round(n, 0), "N0", "ja-JP"),

    Drugs = Table.Buffer(#"M_薬品"),
    Trans = Table.SelectRows(Drugs, each [経過措置期限] <> null),

    // 代替の屋号：同じまとまりで経過措置のない薬。後発を優先し、直近の払出が多い順
    Alt   = Table.AddColumn(Trans, "代替の屋号", (r) =>
                if r[需要のまとまり] = null then null
                else
                    let
                        g = Table.SelectRows(Drugs, each [需要のまとまり] = r[需要のまとまり] and [薬品名] <> r[薬品名] and [経過措置期限] = null),
                        k = Table.AddColumn(g, "k", each if [先発後発] = "後発" then 0 else 1, Int64.Type),
                        s = Table.Sort(k, {{"k", Order.Ascending}, {"直近の払出合計", Order.Descending}})
                    in
                        if Table.IsEmpty(s) then null else s{0}[薬品名], type nullable text),

    Calc  = Table.AddColumn(Alt, "計算", each
        let
            left    = Duration.Days([経過措置期限] - BaseDate),
            perDay  = [#"1日平均払出"] ?? 0,
            stock   = [最終在庫] ?? 0,
            usable  = perDay * List.Max({0, left}),
            surplus = List.Max({0, stock - usable}),
            act0    = if left < 0 then "期限切れ。返品か廃棄、処方の切替を確認"
                      else if surplus > 0 then "余る " & F(surplus) & " を返品・転送。発注は停止"
                      else "使い切れる見込み。発注量を残り日数分に抑える",
            act     = if [代替の屋号] = null then act0 else act0 & "／代替は「" & [代替の屋号] & "」。在庫がなくなるまで管理を継続"
        in
            [残り日数 = left, 期限切迫 = left <= Near, 在庫 = stock, 期限までに使い切れる量 = usable, 余る量 = surplus, 処置の案 = act]),
    E     = Table.ExpandRecordColumn(Calc, "計算", {"残り日数", "期限切迫", "在庫", "期限までに使い切れる量", "余る量", "処置の案"}),
    Out   = Table.SelectColumns(E, {"薬品名", "略称", "需要のまとまり", "経過措置期限", "残り日数", "期限切迫", "在庫", "1日平均払出",
                "期限までに使い切れる量", "余る量", "代替の屋号", "処置の案"}),
    Typed = Table.TransformColumnTypes(Out, {
                {"薬品名", type text}, {"略称", type text}, {"需要のまとまり", type text}, {"経過措置期限", type date},
                {"残り日数", Int64.Type}, {"期限切迫", type logical}, {"在庫", type number}, {"1日平均払出", type number},
                {"期限までに使い切れる量", type number}, {"余る量", type number}, {"代替の屋号", type text}, {"処置の案", type text}}),
    Result = Table.Sort(Typed, {{"残り日数", Order.Ascending}, {"薬品名", Order.Ascending}})
in
    Result
