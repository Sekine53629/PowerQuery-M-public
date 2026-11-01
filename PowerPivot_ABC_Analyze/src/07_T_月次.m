// ================================================================================
// ⑦ クエリ名：T_月次（データモデルに読み込む）
//    月×薬品の1行に、出納（実際）と分析データ（予定）を並べる
//    ・月：その月の1日（出納は暦の月単位で出力する前提）
//    ・出納に載った：その月の出納にその薬の行がある（＝何かしら動いた）
//      出納のある月に載っていない薬は、仕入・払出・その他の増減を 0 とする（在庫は空欄）
//    ・予定データあり：設定の開始日～終了日（③ で読む期間）が、その月をまるごと含む
//    ・予定とのずれ ＝ 払出数量 － 予定数量（出納と予定の両方がある月だけ）
//        プラス：臨時処方・新患・他科・飛び処方・再来局・増量
//        マイナス：処方の終了・残薬調整で出ない・来局なし・処方日数の変更
//        ※ 月の合計どうしなので、同じ月の中のプラスとマイナスは打ち消し合う
//    ・入庫の概算 ＝ 仕入数量 ＋ その他の増減のプラス分
// ================================================================================
let
    Prm  = #"設定値",
    S0   = Date.From(Prm[開始日]),
    E0   = Date.From(Prm[終了日]),
    Sui  = #"出納",
    Day  = #"T_日次",

    // ---- 出納：月×品名 ----
    SuiM = Table.Group(Table.AddColumn(Sui, "月", each Date.StartOfMonth([期間開始]), type date), {"月", "品名"}, {
        {"繰越数量",     each List.Sum([繰越数量]), type nullable number},
        {"仕入数量",     each List.Sum([仕入数量]), type nullable number},
        {"払出数量",     each List.Sum([払出数量]), type nullable number},
        {"その他の増減", each List.Sum([その他の増減]), type nullable number},
        {"在庫数量",     each List.Last([在庫数量]), type nullable number}}),
    SuiMonths = List.Buffer(List.Distinct(SuiM[月])),

    // ---- 予定：月×薬品名 ----
    DayM = Table.Group(Table.AddColumn(Day, "月", each Date.StartOfMonth([日付]), type date), {"月", "薬品名"}, {
        {"予定数量",   each List.Sum([予定数量]), type nullable number},
        {"予定患者数", each List.Sum([予定患者数]), type nullable number},
        {"予定日数",   each List.Count(List.Select([予定数量], (q) => q <> null and q > 0)), Int64.Type}}),

    J  = Table.NestedJoin(SuiM, {"月", "品名"}, DayM, {"月", "薬品名"}, "x", JoinKind.FullOuter),
    E  = Table.ExpandTableColumn(J, "x", {"月", "薬品名", "予定数量", "予定患者数", "予定日数"},
             {"月_予定", "薬品名_予定", "予定数量", "予定患者数", "予定日数"}),
    K  = Table.AddColumn(E, "計算", each
        let
            m       = [月] ?? [月_予定],
            onSui   = [品名] <> null,
            suiMon  = List.Contains(SuiMonths, m),
            planMon = m >= S0 and Date.EndOfMonth(m) <= E0,
            zeroIf  = (v) => if v = null and suiMon then 0 else v,
            out     = zeroIf([払出数量]),
            plan    = if planMon then [予定数量] ?? 0 else null,
            gap     = if suiMon and planMon then Number.Round(out - plan, 3) else null,
            other   = zeroIf([その他の増減])
        in
            [月 = m, 薬品名 = [品名] ?? [薬品名_予定],
             出納に載った = onSui, 出納データあり = suiMon, 予定データあり = planMon,
             繰越数量 = [繰越数量], 仕入数量 = zeroIf([仕入数量]), 払出数量 = out,
             その他の増減 = other, 在庫数量 = [在庫数量],
             入庫の概算 = if suiMon then (zeroIf([仕入数量]) ?? 0) + List.Max({0, other ?? 0}) else null,
             予定数量 = plan, 予定患者数 = if planMon then [予定患者数] ?? 0 else null,
             予定日数 = if planMon then [予定日数] ?? 0 else null,
             予定とのずれ = gap,
             ずれの割合 = if gap <> null and plan > 0 then gap / plan else null]),
    Rows   = Table.FromRecords(K[計算], type table [
                 月 = date, 薬品名 = text, 出納に載った = logical, 出納データあり = logical, 予定データあり = logical,
                 繰越数量 = nullable number, 仕入数量 = nullable number, 払出数量 = nullable number,
                 その他の増減 = nullable number, 在庫数量 = nullable number, 入庫の概算 = nullable number,
                 予定数量 = nullable number, 予定患者数 = nullable number, 予定日数 = nullable number,
                 予定とのずれ = nullable number, ずれの割合 = nullable number]),
    Typed  = Table.TransformColumnTypes(Rows, {{"予定日数", Int64.Type}, {"ずれの割合", Percentage.Type}}),
    Result = Table.Sort(Typed, {{"月", Order.Ascending}, {"薬品名", Order.Ascending}})
in
    Result
