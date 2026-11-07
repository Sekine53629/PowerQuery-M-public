// ================================================================================
// ㉕ クエリ名：薬品グループ案（シート「薬品グループ案」に表として読み込む）
//    薬品名を分解して、手入力テーブル「薬品グループ」に入れる 成分規格・需要のまとまり・先発後発 の案を出す（0.7.0）
//    ・後発品の名前は「成分名＋剤形＋規格「屋号」」で統一されているので、「屋号」を外せば 成分規格、
//      剤形（錠・OD錠・カプセルなど）も外せば 需要のまとまり（成分＋規格）になる
//        例：セルトラリン錠25mg「アメル」／セルトラリンOD錠25mg「トーワ」 → まとまり「セルトラリン 25mg」
//    ・内用液・細粒・顆粒・外用は単位が違うので、剤形を残して別のまとまりにする
//    ・先発品（屋号がない）は名前では結べない。「要確認」として残すので、成分名を見て手で同じまとまりに入れる
//    ・漢方・OTC・「徐放錠A」のような変則の名前も「要確認」
//    【載せる薬】（0.8.0）まとめる必要のある薬だけ。理由は列「載せた理由」
//      ・同じまとまり（案）の薬が 2 つ以上ある（需要がかぶる後発どうし）
//      ・経過措置（経過措置薬品リストか手入力の経過措置リスト）
//      ・出荷調整の疑い（⑬ T_まとまり月次 の最新月で G01 か G02）
//      ・供給状況が 限定／供給停止／薬価削除予定／経過措置（薬品グループの手入力）
//      ・切替先が入っている、または切替先に指定されている
//      ・上の薬と同じまとまり（案）の薬（切替先・寄せ先の候補）
//    【列の並び】（0.8.0）先頭 6 列は 薬品グループ と同じ名前・同じ順（薬品名／成分規格／需要のまとまり／先発後発／供給状況／切替先）。
//      この 6 列を範囲選択してコピーし、薬品グループ の表の末尾に貼れる（未登録 ○ の行だけ貼る）。
//      需要のまとまり・先発後発 は、登録済みがあればその値、なければ案
//    ・案はそのまま使わない。確認してから 薬品グループ に貼る（登録済みのまとまりと違えば「違い」に ○。空白の有無は違いとしない）
// ================================================================================
let
    Nz    = #"fx_正規化",
    Drugs = #"M_薬品",
    Clean = (v as any) as nullable text => if v = null then null else let t = Text.Trim(Text.From(v)) in if t = "" then null else t,
    F     = (n as nullable number) as text => if n = null then "" else Number.ToText(Number.Round(n, 0), "N0", "ja-JP"),

    // 剤形（長いものから順に照合する）
    Forms = {"口腔内崩壊錠", "フィルムコーティング錠", "チュアブル錠", "徐放錠", "腸溶錠", "糖衣錠", "OD錠", "錠",
             "カプセル", "ドライシロップ", "シロップ", "点眼液", "点鼻液", "吸入液", "内用液", "液",
             "細粒", "顆粒", "散", "テープ", "パップ", "ゲル", "軟膏", "クリーム", "ローション", "坐剤", "注"},
    TabletForms = {"口腔内崩壊錠", "フィルムコーティング錠", "チュアブル錠", "徐放錠", "腸溶錠", "糖衣錠", "OD錠", "錠", "カプセル"},
    Digits    = {"0" .. "9"},
    UnitChars = {"0" .. "9", ".", "m", "g", "μ", "µ", "L", "l", "%", "/", "単", "位"},

    // 末尾の [3g]・(医療用)・(旧) のようなかっこ書きを外す（繰り返し）
    StripTail = (t as text) as text =>
        let
            s      = Text.Trim(t),
            last   = if s = "" then "" else Text.End(s, 1),
            opener = if last = "]" then "[" else if last = ")" then "(" else if last = "）" then "（"
                     else if last = "】" then "【" else if last = "〕" then "〔" else null,
            p      = if opener = null then -1 else Text.PositionOf(s, opener, Occurrence.Last)
        in
            if p < 0 then s else @StripTail(Text.Start(s, p)),

    Parse = (name as text) as record =>
        let
            n0     = Nz(name),
            hasB   = Text.Contains(n0, "「") and Text.Contains(n0, "」"),
            brand  = if hasB then Text.Trim(Text.BetweenDelimiters(n0, "「", "」")) else null,
            core   = StripTail(if hasB then Text.BeforeDelimiter(n0, "「") else n0),
            chars  = Text.ToList(core),
            dpos   = List.PositionOfAny(chars, Digits),
            run    = if dpos < 0 then {} else List.FirstN(List.Skip(chars, dpos), each List.Contains(UnitChars, _)),
            str0   = Text.Replace(Text.TrimEnd(Text.Combine(run), {"/", "."}), "µ", "μ"),
            okStr  = str0 <> "" and List.ContainsAny(Text.ToList(str0), Digits) and str0 <> Text.Select(str0, Digits & {"."}),
            head   = Text.Trim(if dpos < 0 then core else Text.Start(core, dpos)),
            form   = List.First(List.Select(Forms, (f) => Text.EndsWith(head, f)), null),
            ingr   = Text.Remove(if form = null then head else Text.Start(head, Text.Length(head) - Text.Length(form)), {" ", "　"}),
            tablet = form <> null and List.Contains(TabletForms, form),
            grp    = if ingr = "" or not okStr or form = null then null
                     else if tablet then ingr & " " & str0 else ingr & form & " " & str0,
            spec   = if ingr = "" then null else ingr & (form ?? "") & (if okStr then str0 else ""),
            conf   = if brand <> null and form <> null and okStr then "後発（屋号あり）"
                     else if brand = null and form <> null and okStr then "要確認：屋号なし（先発か旧名称。成分名で結ぶ）"
                     else if form = null then "要確認：剤形が読めない"
                     else "要確認：規格が読めない"
        in
            [成分案 = if ingr = "" then null else ingr, 剤形 = form, 規格 = if okStr then str0 else null, 屋号 = brand,
             成分規格案 = spec, まとまり案 = grp, 先発後発案 = if brand <> null then "後発" else null, 確からしさ = conf],

    // 出荷調整の疑い（⑬ の最新月）
    Adj      = #"T_まとまり月次",
    AdjNames = List.Buffer(if Table.IsEmpty(Adj) then {}
                           else Table.SelectRows(Adj, each [月] = List.Max(Adj[月]) and ([G01] = true or [G02] = true))[薬品名]),

    Base = Table.SelectColumns(Drugs, {"薬品名", "略称", "成分規格", "需要のまとまり", "先発後発", "供給状況", "切替先", "切替元", "経過措置"}),
    P    = Table.ExpandRecordColumn(Table.AddColumn(Base, "p", each Parse([薬品名])), "p",
               {"成分案", "剤形", "規格", "屋号", "成分規格案", "まとまり案", "先発後発案", "確からしさ"}),
    Cnt  = Table.Group(Table.SelectRows(P, each [まとまり案] <> null), {"まとまり案"}, {{"同じまとまりの屋号数", each Table.RowCount(_), Int64.Type}}),
    J    = Table.ExpandTableColumn(Table.NestedJoin(P, {"まとまり案"}, Cnt, {"まとまり案"}, "c", JoinKind.LeftOuter), "c", {"同じまとまりの屋号数"}),

    // 載せる理由（薬ごと）
    R1   = Table.AddColumn(J, "理由1", each Text.Combine(List.RemoveNulls({
               if ([同じまとまりの屋号数] ?? 0) >= 2 then "同じ成分規格が " & F([同じまとまりの屋号数]) & " 品目" else null,
               if [経過措置] = true then "経過措置" else null,
               if List.Contains(AdjNames, [薬品名]) then "出荷調整の疑い" else null,
               if [供給状況] <> null and [供給状況] <> "通常" then Clean([供給状況]) else null,
               if [切替先] <> null or [切替元] <> null then "切替" else null}), "、"), type text),
    // 理由のある薬と同じまとまり（案）の薬も載せる（切替先・寄せ先の候補）
    HitGrp = List.Buffer(List.Distinct(Table.SelectRows(R1, each [理由1] <> "" and [まとまり案] <> null)[まとまり案])),
    R2   = Table.AddColumn(R1, "載せた理由", each
               if [理由1] <> "" then [理由1]
               else if [まとまり案] <> null and List.Contains(HitGrp, [まとまり案]) then "同じ成分規格の候補" else "", type text),
    Sel  = Table.SelectRows(R2, each [載せた理由] <> ""),

    // 薬品グループ と同じ 6 列（登録済みがあればその値、なければ案）
    Cols = Table.AddColumn(Sel, "f", each
               [成分規格_貼 = Clean([成分規格]) ?? [成分規格案],
                需要のまとまり_貼 = Clean([需要のまとまり]) ?? [まとまり案],
                先発後発_貼 = Clean([先発後発]) ?? [先発後発案],
                供給状況_貼 = Clean([供給状況]),
                切替先_貼 = Clean([切替先]),
                未登録 = if [需要のまとまり] = null then "○" else "",
                // 空白の有無だけの違いは違いとみなさない（「アリピプラゾール3mg」と「アリピプラゾール 3mg」）
                違い = if [需要のまとまり] <> null and [まとまり案] <> null
                          and Text.Remove(Nz([需要のまとまり]), {" ", "　"}) <> Text.Remove([まとまり案], {" ", "　"}) then "○" else "",
                並び = if [まとまり案] = null then 1 else 0]),
    E    = Table.ExpandRecordColumn(Table.RemoveColumns(Cols, {"成分規格", "需要のまとまり", "先発後発", "供給状況", "切替先"}), "f",
               {"成分規格_貼", "需要のまとまり_貼", "先発後発_貼", "供給状況_貼", "切替先_貼", "未登録", "違い", "並び"}),
    Ren  = Table.RenameColumns(E, {{"成分規格_貼", "成分規格"}, {"需要のまとまり_貼", "需要のまとまり"}, {"先発後発_貼", "先発後発"},
               {"供給状況_貼", "供給状況"}, {"切替先_貼", "切替先"},
               {"成分案", "成分（案）"}, {"成分規格案", "成分規格（案）"}, {"まとまり案", "需要のまとまり（案）"}, {"先発後発案", "先発後発（案）"}}),
    Out  = Table.SelectColumns(Ren, {"薬品名", "成分規格", "需要のまとまり", "先発後発", "供給状況", "切替先",
               "載せた理由", "確からしさ", "同じまとまりの屋号数", "略称", "成分（案）", "剤形", "規格", "屋号",
               "需要のまとまり（案）", "未登録", "違い", "並び"}),
    Typed = Table.TransformColumnTypes(Out, {
               {"薬品名", type text}, {"成分規格", type text}, {"需要のまとまり", type text}, {"先発後発", type text},
               {"供給状況", type text}, {"切替先", type text}, {"載せた理由", type text}, {"確からしさ", type text},
               {"同じまとまりの屋号数", Int64.Type}, {"略称", type text}, {"成分（案）", type text}, {"剤形", type text},
               {"規格", type text}, {"屋号", type text}, {"需要のまとまり（案）", type text},
               {"未登録", type text}, {"違い", type text}, {"並び", Int64.Type}}),
    Sorted = Table.Sort(Typed, {{"並び", Order.Ascending}, {"需要のまとまり（案）", Order.Ascending}, {"薬品名", Order.Ascending}}),
    Result = Table.RemoveColumns(Sorted, {"並び"})
in
    Result
