// ================================================================================
// ⑲ クエリ名：R_極小需要（シート「E_極小需要」に表として読み込む）
//    E 要管理の極小需要薬：直近 N ヶ月の払出が 設定「極小需要の月の払出」× N 以下の薬（ほとんど出ない薬）
//    ・出納だけに出て、直近の払出も在庫もない薬は載せない
//    ・処置の案：発注除外 → そのまま／払出なし＋在庫あり → 返品・転送して発注除外へ／
//        払出も在庫もなし → 発注除外に登録／在庫なし → 患者ごとに取り寄せ／
//        直近の在庫日数 ≧ 在庫過多とみなす日数 → 最小限だけ持ち余りは転送／それ以外 → 最小限の在庫で運用
// ================================================================================
let
    Prm  = #"設定値",
    Num  = (key as text, dflt as number) as number =>
               let v = Record.FieldOrDefault(Prm, key, null)
               in  if v = null or (v is text and Text.Trim(v) = "") then dflt else Number.From(v),
    TinyOut = Num("極小需要の月の払出", 100),
    N       = Number.RoundDown(Num("直近の月数", 3)),
    Over    = Num("在庫過多とみなす日数", 60),

    Drugs = #"M_薬品",
    Cand  = Table.SelectRows(Drugs, each
                let
                    recent = [直近の払出合計] ?? 0,
                    stock  = [最終在庫] ?? 0
                in
                    not ([出どころ] = "出納だけ" and recent = 0 and stock = 0) and recent <= TinyOut * N),
    Calc  = Table.AddColumn(Cand, "計算", each
        let
            recent = [直近の払出合計] ?? 0,
            stock  = [最終在庫] ?? 0,
            rdays  = [直近の在庫日数],
            excl   = [発注除外] = true,
            act    = if excl then "発注除外（取り寄せ）として運用中"
                     else if recent = 0 and stock > 0 then "払出なし。在庫を返品・転送し、発注除外に登録"
                     else if recent = 0 then "払出も在庫もなし。発注除外に登録"
                     else if stock <= 0 then "在庫なし。患者ごとに取り寄せ（発注除外に登録）"
                     else if rdays <> null and rdays >= Over then "最小限だけ持つ。余りは転送"
                     else "最小限の在庫で運用",
            warn   = (rdays <> null and rdays >= Over) or (recent = 0 and stock > 0)
        in
            [処置の案 = act, 注意 = if warn then "○" else "", 並び = if excl then 1 else 0]),
    E0    = Table.ExpandRecordColumn(Calc, "計算", {"処置の案", "注意", "並び"}),
    E     = Table.TransformColumns(E0, {{"発注除外", each if _ = true then "○" else "", type text}}),
    Out   = Table.SelectColumns(E, {"薬品名", "略称", "需要の型", "平均間隔", "予定患者数合計", "直近の払出合計", "最終在庫",
                "直近の在庫日数", "最後に動いた月", "動いた月数", "発注除外", "注意", "処置の案", "並び"}),
    Ren   = Table.RenameColumns(Out, {{"予定患者数合計", "年間の予定患者数"}, {"直近の払出合計", "直近の払出"}, {"最終在庫", "在庫"}}),
    Typed = Table.TransformColumnTypes(Ren, {
                {"薬品名", type text}, {"略称", type text}, {"需要の型", type text}, {"平均間隔", type number},
                {"年間の予定患者数", type number}, {"直近の払出", type number}, {"在庫", type number}, {"直近の在庫日数", type number},
                {"最後に動いた月", type date}, {"動いた月数", Int64.Type}, {"発注除外", type text}, {"注意", type text},
                {"処置の案", type text}, {"並び", Int64.Type}}),
    Sorted = Table.Sort(Typed, {{"並び", Order.Ascending}, {"直近の在庫日数", Order.Descending}, {"在庫", Order.Descending}}),
    Result = Table.RemoveColumns(Sorted, {"並び"})
in
    Result
