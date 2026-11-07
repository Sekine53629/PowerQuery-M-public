// ================================================================================
// ㉕ クエリ名：薬品グループ案（シート「薬品グループ案」に表として読み込む）
//    薬品名を分解して、手入力テーブル「薬品グループ」に入れる 成分規格・需要のまとまり・先発後発 の案を出す（0.7.0）
//    ・後発品の名前は「成分名＋剤形＋規格「屋号」」で統一されているので、「屋号」を外せば 成分規格、
//      剤形（錠・OD錠・カプセルなど）も外せば 需要のまとまり（成分＋規格）になる
//        例：セルトラリン錠25mg「アメル」／セルトラリンOD錠25mg「トーワ」 → まとまり「セルトラリン 25mg」
//    ・内用液・細粒・顆粒・外用は単位が違うので、剤形を残して別のまとまりにする
//    ・先発品（屋号がない）は名前では結べない。「要確認」として残すので、成分名を見て手で同じまとまりに入れる
//    ・漢方・OTC・「徐放錠A」のような変則の名前も「要確認」
//    ・案はそのまま使わない。確認してから 薬品グループ に貼る（登録済みのまとまりと違えば「違い」に ○。空白の有無は違いとしない）
//    ・同じまとまりの屋号数：案が同じ薬が何品目あるか（2 以上なら需要がかぶっている候補）
// ================================================================================
let
    Nz    = #"fx_正規化",
    Drugs = #"M_薬品",
    Clean = (v as any) as nullable text => if v = null then null else let t = Text.Trim(Text.From(v)) in if t = "" then null else t,

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

    Base = Table.SelectColumns(Drugs, {"薬品名", "略称", "需要のまとまり"}),
    P    = Table.ExpandRecordColumn(Table.AddColumn(Base, "p", each Parse([薬品名])), "p",
               {"成分案", "剤形", "規格", "屋号", "成分規格案", "まとまり案", "先発後発案", "確からしさ"}),
    Cnt  = Table.Group(Table.SelectRows(P, each [まとまり案] <> null), {"まとまり案"}, {{"同じまとまりの屋号数", each Table.RowCount(_), Int64.Type}}),
    J    = Table.ExpandTableColumn(Table.NestedJoin(P, {"まとまり案"}, Cnt, {"まとまり案"}, "c", JoinKind.LeftOuter), "c", {"同じまとまりの屋号数"}),
    Flag = Table.AddColumn(J, "f", each
               [未登録 = if [需要のまとまり] = null then "○" else "",
                // 空白の有無だけの違いは違いとみなさない（「アリピプラゾール3mg」と「アリピプラゾール 3mg」）
                違い   = if [需要のまとまり] <> null and [まとまり案] <> null
                            and Text.Remove(Nz([需要のまとまり]), {" ", "　"}) <> Text.Remove([まとまり案], {" ", "　"}) then "○" else "",
                並び   = if [まとまり案] = null then 1 else 0]),
    E    = Table.ExpandRecordColumn(Flag, "f", {"未登録", "違い", "並び"}),
    Ren  = Table.RenameColumns(E, {{"成分案", "成分（案）"}, {"成分規格案", "成分規格（案）"}, {"まとまり案", "需要のまとまり（案）"},
               {"先発後発案", "先発後発（案）"}, {"需要のまとまり", "登録済みのまとまり"}}),
    Out  = Table.SelectColumns(Ren, {"薬品名", "略称", "成分（案）", "剤形", "規格", "屋号", "成分規格（案）", "需要のまとまり（案）",
               "先発後発（案）", "確からしさ", "同じまとまりの屋号数", "登録済みのまとまり", "未登録", "違い", "並び"}),
    Typed = Table.TransformColumnTypes(Out, {
               {"薬品名", type text}, {"略称", type text}, {"成分（案）", type text}, {"剤形", type text}, {"規格", type text},
               {"屋号", type text}, {"成分規格（案）", type text}, {"需要のまとまり（案）", type text}, {"先発後発（案）", type text},
               {"確からしさ", type text}, {"同じまとまりの屋号数", Int64.Type}, {"登録済みのまとまり", type text},
               {"未登録", type text}, {"違い", type text}, {"並び", Int64.Type}}),
    Sorted = Table.Sort(Typed, {{"並び", Order.Ascending}, {"需要のまとまり（案）", Order.Ascending}, {"薬品名", Order.Ascending}}),
    Result = Table.RemoveColumns(Sorted, {"並び"})
in
    Result
