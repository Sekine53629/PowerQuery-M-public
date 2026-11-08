// ================================================================================
// ㉕ クエリ名：薬品グループ案（シート「薬品グループ案」に表として読み込む）
//    薬品名を分解して、手入力テーブル「薬品グループ」に入れる 成分規格・需要のまとまり・先発後発 の案を出す（0.7.0）
//    ・後発品の名前は「成分名＋剤形＋規格「屋号」」で統一されているので、「屋号」を外せば 成分規格、
//      剤形（錠・OD錠・カプセルなど）も外せば 需要のまとまり（成分＋規格）になる
//        例：セルトラリン錠25mg「アメル」／セルトラリンOD錠25mg「トーワ」 → まとまり「セルトラリン 25mg」
//    ・剤形の扱い（0.10.1。利用者の指示）：錠・OD錠・D錠・口腔内崩壊錠・ザイディス・カプセル・フィルム（口腔内フィルム）は同じまとまり。
//      徐放錠・CR錠・SRカプセルなどの徐放系は「成分徐放 規格」、腸溶錠は「成分腸溶 規格」で別のまとまり。
//      フィルムコーティング錠は普通の錠（ODフィルム・フィルム剤とは別物）
//      散・顆粒・細粒・末 は「成分散剤 規格」、内用液・シロップ・液 は「成分液剤 規格」で同じまとまり（変更調剤の イ・ウ）。
//      内服のまとまりでは成分名の末尾の塩・エステル（塩酸塩・マレイン酸塩・酒石酸塩・酪酸エステル・ナトリウム／Na など）を外して結ぶ
//      （後発品の名前では省かれることが多い。0.11.0）。外用は残す。成分（案）・成分規格 は表記どおり
//      ドライシロップと外用（点眼液・外用液・テープなど）は剤形を残して別のまとまり
//    ・規格の読み取り（0.10.1）：㎎・㎍・㎖ のような 1 文字の単位と大文字小文字をそろえる（25㎎・25MG → 25mg）。
//      数字が複数あるときは単位の付いたものを優先（「錠7.5 7.5mg」→ 7.5mg）、剤形が数字の後ろでも読む（「1%点眼液 5mL」→ 1%）。
//      mg/mL・単位・万単位・個 も単位として読む。規格は名前全体から読む（「錠2「サワイ」 2mg」「錠(25mg)」「糖衣錠(0.5) 0.5mg」）。
//      先頭の ★（OTC の印）は外す。配合錠・配合顆粒・眼軟膏・うがい液・サポ も剤形として読む（0.11.0）
//    ・先発品（屋号がない）は ㉖ 先発成分 の表（先発名 → 成分名）で成分名を引いて、後発と同じまとまりに結ぶ（0.11.0。先発後発（案）＝先発、
//      列 先発名 に引いた名前）。表にない先発品は「要確認」として残すので、成分名を見て手で入れる（手入力テーブル「先発成分」に足してもよい）
//    ・漢方・OTC のように剤形や規格のない名前は「要確認」（「徐放錠A200mg」のような形は 0.10.1 から読める）
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

    // 先発名 → 成分名：㉖ の表に、手入力テーブル「先発成分」（あれば）を優先して重ねる。長い先発名から順に名前の先頭と照合する
    HandBrand = try Table.SelectColumns(Excel.CurrentWorkbook(){[Name = "先発成分"]}[Content], {"先発名", "成分名"}, MissingField.UseNull)
                otherwise #table({"先発名", "成分名"}, {}),
    BrandTbl  = Table.Buffer(Table.Distinct(Table.SelectRows(
                    Table.TransformColumns(Table.Combine({HandBrand, Table.SelectColumns(#"先発成分", {"先発名", "成分名"})}),
                        {{"先発名", (v) => if v = null then null else Text.Trim(Nz(Text.From(v))), type text},
                         {"成分名", (v) => if v = null then null else Text.Trim(Text.From(v)), type text}}),
                    each [先発名] <> null and [先発名] <> "" and [成分名] <> null and [成分名] <> ""), {"先発名"})),
    BrandNames = List.Buffer(List.Sort(BrandTbl[先発名], (a, b) => Value.Compare(Text.Length(b), Text.Length(a)))),
    BrandOf  = (ingr as text) as nullable text => List.First(List.Select(BrandNames, (b) => Text.StartsWith(ingr, b)), null),
    IngrOf   = (b as text) as text => BrandTbl{[先発名 = b]}[成分名],

    // まとまりのキーでは、成分名の末尾の塩・エステルの表記を外す（後発品の名前では省かれることが多い。例：ゾルピデム酒石酸塩 → ゾルピデム）。
    // 内服（錠剤系・散剤系・液剤系・徐放系・腸溶系）だけ。外用はエステルで薬が変わる（ステロイドなど）ので残す。成分（案）・成分規格 は表記どおり
    SaltSuffix = {"酪酸エステルプロピオン酸エステル", "プロピオン酸エステル", "吉草酸エステル", "酪酸エステル", "酢酸エステル", "リン酸エステル", "エステル",
                  "臭化水素酸塩", "マレイン酸塩", "酒石酸塩", "クエン酸塩", "フマル酸塩", "メシル酸塩", "ベシル酸塩", "シュウ酸塩", "コハク酸塩", "トシル酸塩",
                  "パモ酸塩", "塩酸塩", "硫酸塩", "硝酸塩", "リン酸塩", "酢酸塩", "二カリウム", "ナトリウム", "カリウム", "カルシウム", "マグネシウム",
                  "塩化物", "臭化物", "水和物", "Na"},
    StripSalt = (t as text) as text =>
        let hit = List.First(List.Select(SaltSuffix, (sfx) => Text.EndsWith(t, sfx) and Text.Length(t) - Text.Length(sfx) >= 3), null)
        in  if hit = null then t else @StripSalt(Text.Start(t, Text.Length(t) - Text.Length(hit))),
    Clean = (v as any) as nullable text => if v = null then null else let t = Text.Trim(Text.From(v)) in if t = "" then null else t,
    F     = (n as nullable number) as text => if n = null then "" else Number.ToText(Number.Round(n, 0), "N0", "ja-JP"),

    // ---- 剤形（長いものから順に照合する。名前のどこにあってもよい）----
    //   変更調剤で「類似する別剤形」とされる範囲（保医発0305第12号）を同じまとまりにする：
    //     錠剤系（ア）：錠・OD錠・D錠・口腔内崩壊錠・ザイディス・チュアブル錠・糖衣錠・フィルムコーティング錠・カプセル・ソフトカプセル・丸、
    //       利用者の指示で フィルム（口腔内フィルム。ODフィルム）も → まとまり「成分 規格」
    //     散剤系（イ）：散・顆粒・細粒・末 → 「成分散剤 規格」　　液剤系（ウ）：内用液・シロップ・液 → 「成分液剤 規格」
    //     ドライシロップは調剤の仕方で イ にも ウ にもなるので、まとめず剤形を残す（手で寄せる）
    //   徐放系（別のまとまり「成分徐放 規格」）：徐放錠・徐放カプセル・CR錠・SR錠・SRカプセル・LA錠・ER錠・XR錠・R錠
    //   腸溶系（別のまとまり「成分腸溶 規格」）：腸溶錠・腸溶カプセル
    //   外用（点眼液・外用液・テープ・軟膏など）は変更調剤の対象外なので、剤形を残して別のまとまり
    Forms = {"口腔内崩壊錠", "フィルムコーティング錠", "フィルムコート錠", "チュアブル錠", "ザイディス錠", "ザイディス",
             "徐放カプセル", "徐放錠", "腸溶カプセル", "腸溶錠", "糖衣錠", "舌下錠", "バッカル錠", "配合OD錠", "配合錠", "配合カプセル",
             "ODフィルム", "OD錠", "D錠", "CR錠", "CRカプセル", "SR錠", "SRカプセル", "LA錠", "LAカプセル", "ER錠", "XR錠", "R錠", "錠",
             "ソフトカプセル", "カプセル", "フィルム", "ドライシロップ", "シロップ", "点眼液", "点鼻液", "吸入液", "うがい液", "外用液", "内用液", "液",
             "配合顆粒", "配合散", "細粒", "顆粒", "散", "末", "丸", "トローチ", "テープ", "パップ", "ゲル", "眼軟膏", "軟膏", "クリーム", "ローション",
             "坐剤", "サポ", "注"},
    TabletForms = {"口腔内崩壊錠", "フィルムコーティング錠", "フィルムコート錠", "チュアブル錠", "ザイディス錠", "ザイディス", "糖衣錠",
                   "配合OD錠", "配合錠", "配合カプセル", "ODフィルム", "OD錠", "D錠", "錠", "ソフトカプセル", "カプセル", "フィルム", "丸"},
    PowderForms = {"配合顆粒", "配合散", "散", "顆粒", "細粒", "末"},
    LiquidForms = {"内用液", "シロップ", "液"},
    SlowForms   = {"徐放カプセル", "徐放錠", "CR錠", "CRカプセル", "SR錠", "SRカプセル", "LA錠", "LAカプセル", "ER錠", "XR錠", "R錠"},
    EntericForms = {"腸溶カプセル", "腸溶錠"},
    Digits    = {"0" .. "9"},
    Letters   = {"a" .. "z"} & {"A" .. "Z"},
    NumChars  = Digits & {".", ","},
    // 数字のあとに続く単位の文字（mg・μg・mL・%・単位・個・万単位・mg/mL など）
    SufChars  = Letters & {"μ", "%", "/", "単", "位", "個", "万", "千"},
    UnitLetters = Letters & {"μ", "%", "単", "個"},
    // 単位の表記ゆれ：1 文字の単位（㎎ ㎍ ㎖ など）と µ を直してから読む。大文字小文字は読んだあとでそろえる
    UnitMap   = {{"㎎", "mg"}, {"㎍", "μg"}, {"㎖", "mL"}, {"㎕", "μL"}, {"㎏", "kg"}, {"㎜", "mm"}, {"㎝", "cm"}, {"µ", "μ"}},
    NormUnits = (t as text) as text => List.Accumulate(UnitMap, t, (acc, pr) => Text.Replace(acc, pr{0}, pr{1})),
    UnitFix   = {{"ml", "mL"}, {"mcg", "μg"}, {"ug", "μg"}, {"iu", "IU"}, {"meq", "mEq"}},
    CanonUnit = (u as text) as text => List.Accumulate(UnitFix, Text.Lower(u), (acc, pr) => Text.Replace(acc, pr{0}, pr{1})),

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

    // 先頭の印（★ などの OTC の印）を外す
    Marks = {"★", "☆", "※", "◎", "○", "●", "◆", "◇", "▲", "△", "▼", "▽", "■", "□", " ", "　"},
    // 名前の中の「数字＋単位」の並びを全部拾う（位置・数・単位）。前が数字のもの（7.5 の 5）は拾わない
    Runs = (core as text) as list =>
        let
            chars  = Text.ToList(core),
            starts = List.Select(List.Positions(chars), (i) =>
                         List.Contains(Digits, chars{i}) and (i = 0 or not List.Contains(NumChars, chars{i - 1}))),
            one    = (i as number) as record =>
                let
                    num  = List.FirstN(List.Skip(chars, i), (c) => List.Contains(NumChars, c)),
                    rest = List.Skip(chars, i + List.Count(num)),
                    suf  = List.FirstN(rest, (c) => List.Contains(SufChars, c)),
                    n    = Text.Remove(Text.TrimEnd(Text.Combine(num), {".", ","}), {","}),
                    u    = CanonUnit(Text.TrimEnd(Text.Combine(suf), {"/", "."}))
                in
                    [位置 = i, 数 = n, 単位 = u, 単位あり = List.ContainsAny(Text.ToList(u), UnitLetters)]
        in
            List.Transform(starts, one),

    Parse = (name as text) as record =>
        let
            n0     = Text.TrimStart(NormUnits(Nz(name)), Marks),
            hasB   = Text.Contains(n0, "「") and Text.Contains(n0, "」"),
            brand  = if hasB then Text.Trim(Text.BetweenDelimiters(n0, "「", "」")) else null,
            // 成分・剤形は「屋号」より前から読む。規格は名前全体から読む（「錠2「サワイ」 2mg」「錠(25mg)」のように後ろにあることがある）
            pre    = if hasB then Text.BeforeDelimiter(n0, "「") else n0,
            full   = if hasB then pre & " " & Text.AfterDelimiter(n0, "」") else n0,
            form   = List.First(List.Select(Forms, (f) => Text.Contains(pre, f)), null),
            fPos   = if form = null then -1 else Text.PositionOf(pre, form),
            runs   = Runs(full),
            // 規格：単位の付いた数字を優先。なければ剤形の直後の数字（「錠7.5」のような書き方）
            withU  = List.Select(runs, (r) => r[単位あり]),
            afterF = if form = null then {} else List.Select(runs, (r) => r[位置] = fPos + Text.Length(form)),
            pick   = List.First(withU, null) ?? List.First(afterF, null),
            str0   = if pick = null then null else pick[数] & pick[単位],
            okStr  = pick <> null,
            // 成分：剤形か規格の数字のどちらか早い方の手前まで（B12 のような成分の中の数字は切らない）。末尾のかっこ書きは外す
            dPos   = if pick = null then -1 else pick[位置],
            cut    = List.Min(List.RemoveNulls({if fPos >= 0 then fPos else null, if dPos >= 0 then dPos else null}), null),
            ingr0  = Text.Remove(StripTail(if cut = null then pre else Text.Start(full, cut)), {" ", "　"}),
            // 屋号のない薬は先発名の表で成分名を引く（先頭一致）。引けたら成分名で結ぶ
            brand0 = if brand = null and ingr0 <> "" then BrandOf(ingr0) else null,
            ingr   = if brand0 = null then ingr0 else IngrOf(brand0),
            key    = if ingr = "" then "" else StripSalt(ingr),
            grp    = if ingr = "" or not okStr or form = null then null
                     else if List.Contains(TabletForms, form) then key & " " & str0
                     else if List.Contains(SlowForms, form) then key & "徐放 " & str0
                     else if List.Contains(EntericForms, form) then key & "腸溶 " & str0
                     else if List.Contains(PowderForms, form) then key & "散剤 " & str0
                     else if List.Contains(LiquidForms, form) then key & "液剤 " & str0
                     else ingr & form & " " & str0,
            spec   = if ingr = "" then null else ingr & (form ?? "") & (if okStr then str0 else ""),
            conf   = if brand <> null and form <> null and okStr then "後発（屋号あり）"
                     else if brand0 <> null and form <> null and okStr then "先発（成分名で結んだ）"
                     else if brand = null and form <> null and okStr then "要確認：屋号なし（先発か旧名称。成分名で結ぶ）"
                     else if form = null then "要確認：剤形が読めない"
                     else "要確認：規格が読めない"
        in
            [成分案 = if ingr = "" then null else ingr, 剤形 = form, 規格 = if okStr then str0 else null, 屋号 = brand, 先発名 = brand0,
             成分規格案 = spec, まとまり案 = grp,
             先発後発案 = if brand <> null then "後発" else if brand0 <> null then "先発" else null, 確からしさ = conf],

    // 出荷調整の疑い（⑬ の最新月）
    Adj      = #"T_まとまり月次",
    AdjNames = List.Buffer(if Table.IsEmpty(Adj) then {}
                           else Table.SelectRows(Adj, each [月] = List.Max(Adj[月]) and ([G01] = true or [G02] = true))[薬品名]),

    Base = Table.SelectColumns(Drugs, {"薬品名", "略称", "成分規格", "需要のまとまり", "先発後発", "供給状況", "切替先", "切替元", "経過措置"}),
    P    = Table.ExpandRecordColumn(Table.AddColumn(Base, "p", each Parse([薬品名])), "p",
               {"成分案", "剤形", "規格", "屋号", "先発名", "成分規格案", "まとまり案", "先発後発案", "確からしさ"}),
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
               "載せた理由", "確からしさ", "同じまとまりの屋号数", "略称", "成分（案）", "剤形", "規格", "屋号", "先発名",
               "需要のまとまり（案）", "未登録", "違い", "並び"}),
    Typed = Table.TransformColumnTypes(Out, {
               {"薬品名", type text}, {"成分規格", type text}, {"需要のまとまり", type text}, {"先発後発", type text},
               {"供給状況", type text}, {"切替先", type text}, {"載せた理由", type text}, {"確からしさ", type text},
               {"同じまとまりの屋号数", Int64.Type}, {"略称", type text}, {"成分（案）", type text}, {"剤形", type text},
               {"規格", type text}, {"屋号", type text}, {"先発名", type text}, {"需要のまとまり（案）", type text},
               {"未登録", type text}, {"違い", type text}, {"並び", Int64.Type}}),
    Sorted = Table.Sort(Typed, {{"並び", Order.Ascending}, {"需要のまとまり（案）", Order.Ascending}, {"薬品名", Order.Ascending}}),
    Result = Table.RemoveColumns(Sorted, {"並び"})
in
    Result
