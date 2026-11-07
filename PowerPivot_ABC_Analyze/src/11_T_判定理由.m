// ================================================================================
// ⑪ クエリ名：T_判定理由（データモデルに読み込む）
//    薬品×理由コードの1行。M_薬品 の指標が条件に当てはまった薬に、コードと文言を付ける
//    ・値：その薬の数字（文言の {値}）　線：判定に使った設定の値（文言の {線}）
//    ・基準日：設定の発注日（空欄なら今日）。経過措置の残り日数に使う
//    ・線は設定テーブルで変える（下の Num の第2引数は空欄のときの仮の値）
//    ・S05：直近の在庫日数（M_薬品）で判定。E04：経過措置品がまとまりの中で一番出ている（M_薬品「まとまり内で最多」）
//    ・D06：区分 D（患者が極端に少ない。M_薬品「区分」「患者の最少月」。0.5.0）。出納だけの薬は Q01 のまま。切替元は E05 だけ
//    ・E05／E06：薬品グループの 切替先（0.7.0）。切替元には E05（使い切り日数）、切替先には E06（切替元の在庫）
// ================================================================================
let
    Prm  = #"設定値",
    Num  = (key as text, dflt as number) as number =>
               let v = Record.FieldOrDefault(Prm, key, null)
               in  if v = null or (v is text and Text.Trim(v) = "") then dflt else Number.From(v),
    StillMonths = Num("不動とみなす月数", 3),
    OverDays    = Num("在庫過多とみなす日数", 60),
    GapRate     = Num("ずれの注意割合", 0.3),
    GapSize     = Num("ずれの大きさの注意割合", 0.5),
    DLine       = Num("区分Dの患者数", 2),
    BaseDate    = Prm[発注日],

    Codes = Table.Buffer(#"M_理由コード"),
    Drugs = #"M_薬品",

    // 1つの薬について、当てはまるコードを {コード, 値, 線} の並びで返す
    Hits = (r as record) as list =>
        let
            stock = r[最終在庫] ?? 0,
            still = r[動きのない月数],
            days  = r[在庫日数],
            less  = r[その他の減少] ?? 0,
            adi   = r[平均間隔],
            typ   = r[需要の型],
            gr    = r[ずれの割合],
            gs    = r[ずれの大きさ],
            trans = r[経過措置期限],
            left  = if trans = null then null else Duration.Days(trans - BaseDate),
            out   = r[払出合計] ?? 0,
            rdays = r[直近の在庫日数],
            topG  = r[まとまり内で最多] = true
        in
            List.RemoveNulls({
                if stock > 0 and still <> null and still >= StillMonths then {"S01", still, StillMonths} else null,
                if days <> null and days >= OverDays then {"S02", days, OverDays} else null,
                if rdays <> null and rdays >= OverDays then {"S05", rdays, OverDays} else null,
                if r[不動品リスト] = true then {"S06", r[システム不動日数], null} else null,
                if less > 0 then {"S04", less, null} else null,
                if typ = "安定" then {"D01", adi, null}
                else if typ = "ばらつき大" then {"D02", adi, null}
                else if typ = "間欠" then {"D03", adi, null}
                else if typ = "まばら・大口" then {"D04", adi, null} else null,
                if (r[予定外の月数] ?? 0) >= 1 then {"D05", r[予定外の月数], 1} else null,
                if r[区分] = "D" and r[出どころ] <> "出納だけ" and r[切替先] = null then {"D06", r[患者の最少月], DLine} else null,
                if gr <> null and gr >= GapRate then {"P01", gr * 100, GapRate * 100} else null,
                if gr <> null and gr <= -GapRate then {"P02", -gr * 100, GapRate * 100} else null,
                if gs <> null and gs >= GapSize then {"P03", gs * 100, GapSize * 100} else null,
                if left <> null and left >= 0 then {"E01", left, null} else null,
                if left <> null and left < 0 then {"E02", -left, null} else null,
                if r[発注除外] = true and out > 0 then {"E03", out, null} else null,
                if trans <> null and topG then {"E04", null, null} else null,
                if r[切替先] <> null then {"E05", r[切替の使い切り日数], null} else null,
                if r[切替元] <> null then {"E06", r[切替元の在庫], null} else null,
                if r[出どころ] = "出納だけ" then {"Q01", null, null} else null,
                if r[出どころ] = "予定だけ" then {"Q02", null, null} else null
            }),

    Rows = List.Combine(Table.TransformRows(Drugs, (r) =>
               List.Transform(Hits(r), (h) => [薬品名 = r[薬品名], コード = h{0}, 値 = h{1}, 線 = h{2}, 在庫 = r[最終在庫] ?? 0]))),
    T    = Table.FromRecords(Rows, type table [薬品名 = text, コード = text, 値 = nullable number, 線 = nullable number, 在庫 = nullable number]),

    // 文言を作る：{値} {線} {在庫} を数字に置き換える
    F    = (n as nullable number) as text => if n = null then "" else Number.ToText(Number.Round(n, 0), "N0", "ja-JP"),
    J    = Table.ExpandTableColumn(Table.NestedJoin(T, {"コード"}, Codes, {"コード"}, "c", JoinKind.LeftOuter), "c",
               {"分類", "種類", "優先度", "文言"}, {"分類", "種類", "優先度", "文言の型"}),
    Msg  = Table.AddColumn(J, "文言", each
               Text.Replace(Text.Replace(Text.Replace([文言の型] ?? "", "{値}", F([値])), "{線}", F([線])), "{在庫}", F([在庫])),
               type text),
    Out  = Table.SelectColumns(Msg, {"薬品名", "コード", "分類", "種類", "優先度", "文言", "値", "線"}),
    Typed = Table.TransformColumnTypes(Out, {{"分類", type text}, {"種類", type text}, {"優先度", Int64.Type}}),
    Result = Table.Sort(Typed, {{"優先度", Order.Ascending}, {"コード", Order.Ascending}, {"薬品名", Order.Ascending}})
in
    Result
