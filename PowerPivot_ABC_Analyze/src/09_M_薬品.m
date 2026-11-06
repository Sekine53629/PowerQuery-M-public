// ================================================================================
// ⑨ クエリ名：M_薬品（データモデルに読み込む）
//    薬品ごとの1行。属性（略称・単位・棚番・経過措置・発注除外・グループ）と、在庫の必要性を判断する指標
//    T_日次・T_月次・T_判定理由 の [薬品名] をここにつなぐ
//
//    【予定の指標】T_日次（① の開始日～終了日＝発注日までの 1 年間）から
//      予定日数       ：その薬の予定があった日数
//      平均間隔       ：処方のあった日数 ÷ 予定日数（1 なら毎日、10 なら10日に1回）
//      ばらつき（CV²）：予定があった日の数量の（標準偏差 ÷ 平均）の2乗
//      需要の型       ：平均間隔 1.32、ばらつき 0.49 を境に4つに分ける（Syntetos–Boylan）
//                        安定／ばらつき大／間欠／まばら・大口
//      1日平均予定    ：予定数量の合計 ÷ 処方のあった日数
//    【出納の指標】T_月次（出納のある月）から
//      動いた月数     ：出納に載った月の数（出納の月数のうち）
//      1日平均払出    ：払出の合計 ÷ 出納の月の暦日数
//      最終在庫       ：最後に出納に載った月の在庫数量（そのあと動いていなければ今も同じ）
//      動きのない月数 ：最後に載った月から、一番新しい出納の月までの月数
//      在庫日数       ：最終在庫 ÷ 1日平均払出（今の在庫で何日もつか）
//      その他の減少   ：その他の増減のマイナス分の合計（返品・廃棄・移動など）
//      直近の払出合計 ：出納のある月のうち新しい「直近の月数」ヶ月（設定、初期値 3）の払出の合計
//      直近の1日平均払出：直近の払出合計 ÷ その月々の暦日数の合計
//      直近の在庫日数 ：最終在庫 ÷ 直近の1日平均払出（全期間の平均は、途中で需要が落ちた薬を隠すため）
//    【予定とのずれ】出納と予定の両方がある月から
//      ずれの割合     ：（払出 － 予定）の合計 ÷ 予定の合計
//      ずれの大きさ   ：月ごとの |払出 － 予定| の合計 ÷ 予定の合計（打ち消し合いを除いた大きさ）
//      予定外の月数   ：予定が0なのに払出があった月の数（臨時・新患・他科だけで出た月）
//    【そのほか】
//      経過措置期限   ：⑳ 経過措置該当（エリシアSのリスト）を優先し、なければ手入力の経過措置リスト
//      不動品リスト   ：㉑ 不動品リスト に載っている薬（出納の期間より前から動いていない薬もここで拾う）
//      システム不動日数／システム最終払出日／システム在庫／薬価／YJコード：㉑ を薬ごとにまとめたもの
//      供給状況       ：薬品グループの列（通常／限定／供給停止／経過措置／薬価削除予定。手入力）
//      まとまり内で最多：同じ「需要のまとまり」（屋号が2つ以上）の中で、直近の払出が一番多い薬（E04 に使う）
// ================================================================================
let
    Prm  = #"設定値",
    Num  = (key as text, dflt as number) as number =>
               let v = Record.FieldOrDefault(Prm, key, null)
               in  if v = null or (v is text and Text.Trim(v) = "") then dflt else Number.From(v),
    RecentN = Number.RoundDown(Num("直近の月数", 3)),

    Day  = Table.Buffer(#"T_日次"),
    Mon  = Table.Buffer(#"T_月次"),
    Sui  = #"出納",

    Clean = (v as any) as nullable text => if v = null then null else let t = Text.Trim(Text.From(v)) in if t = "" then null else t,
    Hand  = (name as text, cols as list) as table =>
                try Table.SelectColumns(Excel.CurrentWorkbook(){[Name = name]}[Content], cols, MissingField.UseNull)
                otherwise #table(cols, {}),

    // ---- 対象の薬：予定（T_日次）か出納（T_月次）に一度でも出てきた薬 ----
    FudoSrc = #"不動品リスト",
    KeikaSrc = #"経過措置該当",
    Names = List.Distinct(List.RemoveNulls(Day[薬品名] & Mon[薬品名] & FudoSrc[薬品名] & KeikaSrc[薬品名])),
    Base  = Table.FromColumns({Names}, type table [薬品名 = text]),

    // ---- 属性：略称・製剤（実績で一番新しい日付のもの） ----
    AbbrSrc = Table.SelectRows(Table.SelectColumns(#"日次データ", {"日付", "薬品名", "略称", "薬品区分"}),
                  each [略称] <> null or [薬品区分] <> null),
    Abbr    = Table.Group(AbbrSrc, {"薬品名"}, {
                  {"略称", each Clean(List.Last(List.RemoveNulls(Table.Sort(_, {{"日付", Order.Ascending}})[略称]))), type nullable text},
                  {"製剤_予定", each List.Last(List.RemoveNulls([薬品区分])), type nullable text}}),

    // ---- 属性：出納の最新の行（製剤・単位・棚番・コード） ----
    SuiLast = Table.Group(Sui, {"品名"}, {{"r", each Table.Max(_, "期間終了"), type record}}),
    SuiAttr = Table.ExpandRecordColumn(SuiLast, "r", {"製剤", "単位", "棚番", "コード", "品名の先頭空白"}),

    // ---- 属性：手入力テーブル（経過措置リスト・在庫情報・発注除外・薬品グループ） ----
    TransHand = Table.Distinct(Table.SelectRows(Table.TransformColumns(Hand("経過措置リスト", {"薬品名", "経過措置期限"}), {
                {"薬品名", Clean, type nullable text},
                {"経過措置期限", each try Date.From(_) otherwise null, type nullable date}}),
                each [薬品名] <> null), {"薬品名"}),
    // エリシアSの経過措置薬品リスト（⑳）を優先し、手入力は補助
    Trans = Table.Distinct(Table.Combine({Table.SelectColumns(KeikaSrc, {"薬品名", "経過措置期限"}), TransHand}), {"薬品名"}),
    // 不動品リスト（㉑）を薬ごとにまとめる：最終払出日は一番新しいもの、不動日数はその日からの日数
    Fudo = Table.Group(FudoSrc, {"薬品名"}, {
               {"システム最終払出日", each List.Max([最終払出日]), type nullable date},
               {"システム不動日数",   each List.Min([不動日数]), type nullable number},
               {"システム在庫",       each List.Sum([数量]), type nullable number},
               {"薬価",               each List.Max([薬価]), type nullable number},
               {"YJコード",           each List.First(List.RemoveNulls([YJコード]), null), type nullable text}}),
    FudoNames = List.Buffer(Fudo[薬品名]),
    Stock = Table.Distinct(Table.TransformColumns(Hand("在庫情報", {"薬品名", "列数"}), {
                {"薬品名", Clean, type nullable text},
                {"列数", each try Number.From(_) otherwise null, type nullable number}}), {"薬品名"}),
    Excl  = List.Buffer(List.RemoveNulls(List.Transform(Hand("発注除外", {"薬品名"})[薬品名], Clean))),
    Grp   = Table.Distinct(Table.TransformColumns(Hand("薬品グループ", {"薬品名", "成分規格", "需要のまとまり", "先発後発", "供給状況"}), {
                {"薬品名", Clean, type nullable text}, {"成分規格", Clean, type nullable text},
                {"需要のまとまり", Clean, type nullable text}, {"先発後発", Clean, type nullable text},
                {"供給状況", Clean, type nullable text}}), {"薬品名"}),

    // ---- 予定の指標 ----
    DayN  = List.Count(List.Distinct(Day[日付])),
    DayG  = Table.Group(Day, {"薬品名"}, {{"m", each
        let
            q    = List.Select([予定数量], (x) => x <> null and x > 0),
            n    = List.Count(q),
            tot  = List.Sum(q),
            mean = if n > 0 then tot / n else null,
            sd   = if n >= 2 then List.StandardDeviation(q) else null,
            cv2  = if sd <> null and mean > 0 then Number.Power(sd / mean, 2) else null,
            adi  = if n > 0 then DayN / n else null
        in
            [予定日数 = n, 予定数量合計 = tot, 予定患者数合計 = List.Sum([予定患者数]),
             最終予定日 = List.Max(Table.SelectRows(_, each [予定数量] <> null and [予定数量] > 0)[日付]),
             平均間隔 = adi, #"ばらつき（CV²）" = cv2,
             需要の型 = if adi = null then null
                        else if cv2 = null then (if adi < 1.32 then "安定" else "間欠")
                        else if adi < 1.32 then (if cv2 < 0.49 then "安定" else "ばらつき大")
                        else (if cv2 < 0.49 then "間欠" else "まばら・大口"),
             #"1日平均予定" = if DayN > 0 and tot <> null then tot / DayN else null], type record}}),

    // ---- 出納の指標 ----
    SuiMon   = Table.SelectRows(Mon, each [出納データあり]),
    SuiMonths = List.Buffer(List.Distinct(SuiMon[月])),
    MonN     = List.Count(SuiMonths),
    CalDays  = List.Sum(List.Transform(SuiMonths, (m) => Date.Day(Date.EndOfMonth(m)))),
    LatestM  = List.Max(SuiMonths),
    // 直近 N ヶ月（出納のある月の並びで新しい方から）
    RecentMonths = List.Buffer(List.LastN(List.Sort(SuiMonths), RecentN)),
    RecentDays   = List.Sum(List.Transform(RecentMonths, (m) => Date.Day(Date.EndOfMonth(m)))) ?? 0,
    MonG = Table.Group(SuiMon, {"薬品名"}, {{"m", each
        let
            moved   = Table.SelectRows(_, each [出納に載った]),
            lastRow = if Table.IsEmpty(moved) then null else Table.Max(moved, "月"),
            out     = List.Sum([払出数量]) ?? 0,
            perDay  = if CalDays > 0 then out / CalDays else null,
            lastM   = if lastRow = null then null else lastRow[月],
            stock   = if lastRow = null then null else lastRow[在庫数量],
            both    = Table.SelectRows(_, each [予定とのずれ] <> null),
            planSum = List.Sum(both[予定数量]) ?? 0,
            gapSum  = List.Sum(both[予定とのずれ]),
            absSum  = List.Sum(List.Transform(both[予定とのずれ], Number.Abs)),
            rec     = Table.SelectRows(_, each List.Contains(RecentMonths, [月])),
            recOut  = List.Sum(rec[払出数量]) ?? 0,
            recPd   = if RecentDays > 0 then recOut / RecentDays else null
        in
            [動いた月数 = Table.RowCount(moved), 払出合計 = out,
             払出のあった月数 = List.Count(List.Select([払出数量], (x) => x <> null and x > 0)),
             月平均払出 = if MonN > 0 then out / MonN else null,
             #"1日平均払出" = perDay,
             仕入合計 = List.Sum([仕入数量]), 入庫の概算合計 = List.Sum([入庫の概算]),
             その他の減少 = List.Sum(List.Transform([その他の増減], (x) => if x <> null and x < 0 then -x else 0)),
             最後に動いた月 = lastM, 最終在庫 = stock,
             動きのない月数 = if lastM = null then null
                              else (Date.Year(LatestM) - Date.Year(lastM)) * 12 + Date.Month(LatestM) - Date.Month(lastM),
             在庫日数 = if stock <> null and perDay <> null and perDay > 0 then stock / perDay else null,
             比較月数 = Table.RowCount(both),
             ずれの割合 = if planSum > 0 then gapSum / planSum else null,
             ずれの大きさ = if planSum > 0 then absSum / planSum else null,
             予定外の月数 = Table.RowCount(Table.SelectRows(both, each ([予定数量] ?? 0) = 0 and ([払出数量] ?? 0) > 0)),
             直近の払出合計 = recOut, #"直近の1日平均払出" = recPd,
             直近の在庫日数 = if stock <> null and recPd <> null and recPd > 0 then stock / recPd else null],
        type record}}),

    // ---- つなぐ ----
    J1 = Table.ExpandTableColumn(Table.NestedJoin(Base, {"薬品名"}, Abbr, {"薬品名"}, "x", JoinKind.LeftOuter), "x", {"略称", "製剤_予定"}),
    J2 = Table.ExpandTableColumn(Table.NestedJoin(J1, {"薬品名"}, SuiAttr, {"品名"}, "x", JoinKind.LeftOuter), "x",
             {"製剤", "単位", "棚番", "コード", "品名の先頭空白"}),
    J3 = Table.ExpandTableColumn(Table.NestedJoin(J2, {"薬品名"}, Trans, {"薬品名"}, "x", JoinKind.LeftOuter), "x",
             {"経過措置期限"}),
    J4 = Table.ExpandTableColumn(Table.NestedJoin(J3, {"薬品名"}, Stock, {"薬品名"}, "x", JoinKind.LeftOuter), "x", {"列数"}),
    J5 = Table.ExpandTableColumn(Table.NestedJoin(J4, {"薬品名"}, Grp, {"薬品名"}, "x", JoinKind.LeftOuter), "x",
             {"成分規格", "需要のまとまり", "先発後発", "供給状況"}),
    J6 = Table.ExpandTableColumn(Table.NestedJoin(J5, {"薬品名"}, DayG, {"薬品名"}, "x", JoinKind.LeftOuter), "x", {"m"}, {"予定"}),
    J7 = Table.ExpandTableColumn(Table.NestedJoin(J6, {"薬品名"}, MonG, {"薬品名"}, "x", JoinKind.LeftOuter), "x", {"m"}, {"出納"}),

    PlanCols = {"予定日数", "予定数量合計", "予定患者数合計", "最終予定日", "平均間隔", "ばらつき（CV²）", "需要の型", "1日平均予定"},
    SuiCols  = {"動いた月数", "払出合計", "払出のあった月数", "月平均払出", "1日平均払出", "仕入合計", "入庫の概算合計",
                "その他の減少", "最後に動いた月", "最終在庫", "動きのない月数", "在庫日数",
                "比較月数", "ずれの割合", "ずれの大きさ", "予定外の月数",
                "直近の払出合計", "直近の1日平均払出", "直近の在庫日数"},
    E1 = Table.ExpandRecordColumn(Table.TransformColumns(J7, {{"予定", each _ ?? Record.FromList(List.Repeat({null}, List.Count(PlanCols)), PlanCols)}}), "予定", PlanCols),
    E2 = Table.ExpandRecordColumn(Table.TransformColumns(E1, {{"出納", each _ ?? Record.FromList(List.Repeat({null}, List.Count(SuiCols)), SuiCols)}}), "出納", SuiCols),

    // 同じまとまり（屋号が2つ以上）の中で直近の払出が一番多い薬
    TopNames = List.Buffer(
        Table.SelectRows(
            Table.Group(Table.SelectRows(E2, each [需要のまとまり] <> null), {"需要のまとまり"}, {
                {"n",   each Table.RowCount(_), Int64.Type},
                {"top", each Table.First(Table.Sort(_, {{"直近の払出合計", Order.Descending}}))[薬品名], type text}}),
            each [n] >= 2)[top]),

    J8 = Table.ExpandTableColumn(Table.NestedJoin(E2, {"薬品名"}, Fudo, {"薬品名"}, "x", JoinKind.LeftOuter), "x",
             {"システム最終払出日", "システム不動日数", "システム在庫", "薬価", "YJコード"}),

    Fin = Table.AddColumn(J8, "f", each
            [製剤_表示 = [製剤] ?? [製剤_予定],
             経過措置 = List.Contains(Trans[薬品名], [薬品名]),
             発注除外 = List.Contains(Excl, [薬品名]),
             // 動いた月数は、予定だけの薬でも 0 で入るので「> 0」で見る
             出どころ = if ([予定日数] ?? 0) > 0 and ([動いた月数] ?? 0) > 0 then "予定・出納"
                        else if ([予定日数] ?? 0) > 0 then "予定だけ" else "出納だけ",
             まとまり内で最多 = List.Contains(TopNames, [薬品名]),
             不動品リスト = List.Contains(FudoNames, [薬品名])]),
    F2 = Table.ExpandRecordColumn(Fin, "f", {"製剤_表示", "経過措置", "発注除外", "出どころ", "まとまり内で最多", "不動品リスト"}),

    Out = Table.SelectColumns(F2, {
              "薬品名", "略称", "製剤_表示", "単位", "棚番", "コード", "品名の先頭空白",
              "経過措置", "経過措置期限", "発注除外", "成分規格", "需要のまとまり", "先発後発",
              "列数", "出どころ", "供給状況"} & PlanCols & SuiCols & {"まとまり内で最多",
              "不動品リスト", "システム不動日数", "システム最終払出日", "システム在庫", "薬価", "YJコード"}),
    Ren = Table.RenameColumns(Out, {{"製剤_表示", "製剤"}}),
    Typed = Table.TransformColumnTypes(Ren, {
              {"略称", type text}, {"製剤", type text}, {"単位", type text}, {"棚番", type text}, {"コード", type text},
              {"品名の先頭空白", Int64.Type}, {"経過措置", type logical}, {"経過措置期限", type date},
              {"発注除外", type logical}, {"成分規格", type text}, {"需要のまとまり", type text}, {"先発後発", type text},
              {"列数", type number}, {"出どころ", type text},
              {"予定日数", Int64.Type}, {"予定数量合計", type number}, {"予定患者数合計", type number},
              {"最終予定日", type date}, {"平均間隔", type number}, {"ばらつき（CV²）", type number},
              {"需要の型", type text}, {"1日平均予定", type number},
              {"動いた月数", Int64.Type}, {"払出合計", type number}, {"払出のあった月数", Int64.Type},
              {"月平均払出", type number}, {"1日平均払出", type number}, {"仕入合計", type number},
              {"入庫の概算合計", type number}, {"その他の減少", type number}, {"最後に動いた月", type date},
              {"最終在庫", type number}, {"動きのない月数", Int64.Type}, {"在庫日数", type number},
              {"比較月数", Int64.Type}, {"ずれの割合", Percentage.Type}, {"ずれの大きさ", Percentage.Type},
              {"予定外の月数", Int64.Type},
              {"供給状況", type text}, {"直近の払出合計", type number}, {"直近の1日平均払出", type number},
              {"直近の在庫日数", type number}, {"まとまり内で最多", type logical},
              {"不動品リスト", type logical}, {"システム不動日数", type number}, {"システム最終払出日", type date},
              {"システム在庫", type number}, {"薬価", type number}, {"YJコード", type text}}),
    Result = Table.Sort(Typed, {{"払出合計", Order.Descending}, {"薬品名", Order.Ascending}})
in
    Result
