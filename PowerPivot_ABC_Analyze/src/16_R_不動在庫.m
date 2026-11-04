// ================================================================================
// ⑯ クエリ名：R_不動在庫（シート「B_不動在庫」に表として読み込む）
//    B 不動在庫と転送・返品推奨：在庫があり、不動（動きのない月数 ≧ 不動とみなす月数、または不動品リストに載っている）か、
//    在庫過多（直近の在庫日数 ≧ 在庫過多とみなす日数、または直近の払出 0）の薬
//    ・余った量 ＝ 在庫 － 直近の1日平均払出 × 設定「当面の予定日数」（0 未満は 0）
//    ・処置の案：不動 → 返品（できなければ廃棄か他店へ転送）／
//        同じまとまりの別の屋号が動いていて、直近の在庫日数 ≧ 転送候補の在庫日数 → 他店へ転送／
//        在庫日数 ≧ 転送候補の在庫日数 → 他店へ転送か返品／それ以外 → 様子見（発注を止める）
//    ※ 処置の案は目安。最終判断は現場
// ================================================================================
let
    Prm  = #"設定値",
    Num  = (key as text, dflt as number) as number =>
               let v = Record.FieldOrDefault(Prm, key, null)
               in  if v = null or (v is text and Text.Trim(v) = "") then dflt else Number.From(v),
    Still     = Num("不動とみなす月数", 3),
    Over      = Num("在庫過多とみなす日数", 60),
    Horizon   = Num("当面の予定日数", 30),
    TransferD = Num("転送候補の在庫日数", 120),

    Drugs  = Table.Buffer(#"M_薬品"),
    // 同じまとまりで今も動いている屋号（直近の払出が多い順）
    Alive  = Table.Group(
                 Table.SelectRows(Drugs, each ([直近の払出合計] ?? 0) > 0 and [需要のまとまり] <> null),
                 {"需要のまとまり"}, {{"動いている屋号", each Table.Sort(_, {{"直近の払出合計", Order.Descending}})[薬品名], type list}}),

    Cand   = Table.SelectRows(Drugs, each (([最終在庫] ?? 0) > 0 or ([システム在庫] ?? 0) > 0) and (
                 ([動きのない月数] <> null and [動きのない月数] >= Still) or [不動品リスト] = true
                 or ([直近の在庫日数] <> null and [直近の在庫日数] >= Over)
                 or ([直近の払出合計] ?? 0) = 0)),
    J      = Table.ExpandTableColumn(
                 Table.NestedJoin(Cand, {"需要のまとまり"}, Alive, {"需要のまとまり"}, "a", JoinKind.LeftOuter), "a", {"動いている屋号"}),
    Calc   = Table.AddColumn(J, "計算", each
        let
            dead    = ([動きのない月数] <> null and [動きのない月数] >= Still) or [不動品リスト] = true,
            perDay  = [直近の1日平均払出] ?? 0,
            stock   = [最終在庫] ?? [システム在庫] ?? 0,
            surplus = List.Max({0, stock - perDay * Horizon}),
            others  = List.Select([動いている屋号] ?? {}, (n) => n <> [薬品名]),
            rdays   = [直近の在庫日数],
            act     = if dead then "返品（できなければ廃棄か他店へ転送）"
                      else if not List.IsEmpty(others) and (rdays = null or rdays >= TransferD) then
                          "他店へ転送（同じ成分は " & others{0} & " が動いている）"
                      else if rdays <> null and rdays >= TransferD then "他店へ転送か返品"
                      else "様子見（発注を止める）"
        in
            [区分 = if dead then "不動" else "在庫過多", 並び = if dead then 0 else 1,
             余った量 = surplus, 処置の案 = act]),
    E      = Table.ExpandRecordColumn(Calc, "計算", {"区分", "並び", "余った量", "処置の案"}),
    Out    = Table.SelectColumns(E, {"区分", "薬品名", "略称", "棚番", "需要のまとまり", "最終在庫", "最後に動いた月", "動きのない月数",
                 "システム不動日数", "システム最終払出日", "在庫日数", "直近の在庫日数", "直近の1日平均払出", "余った量", "処置の案", "並び"}),
    Ren    = Table.RenameColumns(Out, {{"最終在庫", "在庫"}}),
    Typed  = Table.TransformColumnTypes(Ren, {
                 {"区分", type text}, {"薬品名", type text}, {"略称", type text}, {"棚番", type text}, {"需要のまとまり", type text},
                 {"在庫", type number}, {"最後に動いた月", type date}, {"動きのない月数", Int64.Type},
                 {"システム不動日数", type number}, {"システム最終払出日", type date}, {"在庫日数", type number},
                 {"直近の在庫日数", type number}, {"直近の1日平均払出", type number}, {"余った量", type number},
                 {"処置の案", type text}, {"並び", Int64.Type}}),
    Sorted = Table.Sort(Typed, {{"並び", Order.Ascending}, {"動きのない月数", Order.Descending}, {"直近の在庫日数", Order.Descending}}),
    Result = Table.RemoveColumns(Sorted, {"並び"})
in
    Result
