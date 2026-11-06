// ================================================================================
// ⑭ クエリ名：T_ロット（データモデルに読み込む）
//    ロット（仕入伝票の行）ごとの使用期限と数量。期限の早いロットから順に、直近の払出ペースで
//    「期限までに使い切れる量」と「返す量」を割り当てる
//    ・元：エリシアSの「YYYYMMDD_期限切れリスト」（設定「期限切迫フォルダパス」の一番新しいファイル。.xls／.xlsx／.csv／.txt。
//          列：品名／ロット番号／期限／期限切れ数量／単位）と、㉑ 不動品リスト のロット（期限日・数量）
//      同じ薬品名×ロットが両方にあれば期限切れリストを使う
//    ・品名の先頭の記号（「劇 」「向 」など）と先頭の空白を外して 薬品名 にする（⑤ 出納と同じ）
//    ・残り日数 ＝ 使用期限 － 基準日（設定の発注日）
//    ・期限までに使い切れる量 ＝ 直近の1日平均払出 × 残り日数 から、それより期限の早いロットの数量を引いた分
//      （使用期限が空欄のロットは全量使い切れるものとして扱う）。返す量 ＝ 数量 － 使い切れる量
//    ・フォルダパスが空欄でもリストがなくても壊れない（空の表）
// ================================================================================
let
    Prm   = #"設定値",
    Nz    = #"fx_正規化",
    Open  = #"fx_ファイルを開く",
    Norm  = (t as any) as text =>
        let s = if t = null then "" else Nz(Text.From(t))
        in  Text.Replace(Text.Replace(Text.Remove(s, {" ", "　", "#(lf)", "#(cr)", "#(tab)"}), "(", "（"), ")", "）"),
    CleanPath = (v as any) as nullable text =>
        if v = null or (v is text and Text.Trim(v) = "") then null
        else Text.TrimEnd(Text.Trim(Text.Remove(Text.From(v), {""""})), {"\", "/", " ", "　"}),
    // 数値：.xls と文字のファイルでは、数値が「表示されている文字」で来る（1,234／100錠／¥9,801／▲50 など）
    //   空白・桁区切り・円記号を除き、▲△ はマイナスにする。そのまま数値にできなければ、先頭の数字の部分だけを使う
    Num = (val as any) as nullable number =>
        if val = null then null
        else if val is number then val
        else
            let
                t0    = Text.Remove(Nz(Text.From(val)), {" ", "#(lf)", "#(cr)", "#(tab)", ",", "¥", "￥", "\"}),
                t1    = Text.Replace(Text.Replace(t0, "▲", "-"), "△", "-"),
                whole = if t1 = "" then null else try Number.FromText(t1, "en-US") otherwise null,
                head  = Text.Combine(List.FirstN(Text.ToList(t1), each List.Contains({"0" .. "9", ".", "-"}, _))),
                part  = if Text.Select(head, {"0" .. "9"}) = "" then null else try Number.FromText(head, "en-US") otherwise null
            in  if whole <> null then whole else part,
    Txt = (val as any) as nullable text =>
        if val = null then null else let s = Text.Trim(Text.From(val)) in if s = "" then null else s,
    // 日付：日付のセル・yyyy/mm/dd・yyyy-mm-dd・yyyy年m月d日・8桁の数字（文字のファイルでは文字で入っている）
    Dat = (val as any) as nullable date =>
        if val = null then null
        else if val is date then val
        else if val is datetime then Date.From(val)
        else
            let
                s0 = Nz(Text.From(val)),
                s  = Text.BeforeDelimiter(Text.Trim(Text.Replace(Text.Replace(Text.Replace(Text.Replace(Text.Replace(s0,
                         "年", "/"), "月", "/"), "日", ""), "-", "/"), ".", "/")), " "),
                p  = Text.Split(s, "/"),
                d  = if List.Count(p) = 3 and Text.Length(p{0}) = 4
                     then (try #date(Number.FromText(p{0}), Number.FromText(p{1}), Number.FromText(p{2})) otherwise null)
                     else if Text.Length(s) = 8 and Text.Select(s, {"0" .. "9"}) = s
                     then (try #date(Number.FromText(Text.Start(s, 4)), Number.FromText(Text.Middle(s, 4, 2)),
                                     Number.FromText(Text.End(s, 2))) otherwise null)
                     else null
            in  if d <> null then d else try Date.From(val) otherwise null,
    Marks = {"劇", "向", "毒", "麻", "覚", "習", "生"},
    SplitMark = (v as any) as record =>
        let s = if v = null then null else Nz(Text.From(v))
        in  if s = null or s = "" then [名 = null, 記号 = null]
            else if Text.Length(s) >= 2 and List.Contains(Marks, Text.Start(s, 1)) and Text.Middle(s, 1, 1) = " "
                 then [名 = Text.Trim(Text.Range(s, 2)), 記号 = Text.Start(s, 1)]
            else [名 = s, 記号 = null],
    NewestFile = (folder as nullable text, marker as text) as nullable record =>
        if folder = null then null
        else
            let
                fs  = Table.SelectRows(Folder.Files(folder), each
                          (Text.StartsWith(Text.Lower([Extension]), ".xls") or List.Contains({".csv", ".txt", ".tsv"}, Text.Lower([Extension]))) and not Text.StartsWith([Name], "~$")
                          and Text.Contains(Nz([Name]), marker)),
                srt = Table.Sort(fs, {{"Date modified", Order.Descending}, {"Name", Order.Descending}})
            in
                if Table.IsEmpty(srt) then null else srt{0},
    ReadTable = (f as record) as table =>
        let
            raw    = Open(f[Content], f[Name], "品名"){0}[Data],
            rows   = Table.ToRows(Table.FirstN(raw, 30)),
            hIdx   = List.PositionOf(List.Transform(rows, (r) => List.Contains(List.Transform(r, Norm), "品名")), true)
        in
            if hIdx < 0 then error Error.Record("読めないファイル",
                   "「" & f[Name] & "」に、見出し「品名」のある行がありません（先頭 30 行）")
            else Table.TransformColumnNames(
                     Table.PromoteHeaders(Table.Skip(raw, hIdx), [PromoteAllScalars = true]), Norm),

    Drugs    = Table.Buffer(Table.SelectColumns(#"M_薬品", {"薬品名", "直近の1日平均払出"})),
    BaseDate = Prm[発注日],
    LotType  = type table [薬品名 = text, ロット = nullable text, 使用期限 = nullable date, 数量 = nullable number,
                           単位 = nullable text, 出どころ = text, ファイル名 = nullable text],
    OutType  = type table [薬品名 = text, ロット = nullable text, 使用期限 = nullable date, 数量 = nullable number,
                           単位 = nullable text, 出どころ = text, 残り日数 = nullable number, それより早いロットの数量 = nullable number,
                           期限までに使い切れる量 = nullable number, 返す量 = nullable number, ファイル名 = nullable text],
    Empty    = #table(OutType, {}),

    // ---- 期限切れリスト ----
    Path  = CleanPath(Record.FieldOrDefault(Prm, "期限切迫フォルダパス", null)),
    F     = NewestFile(Path, "期限切れ"),
    Body  = if F = null then null else ReadTable(F),
    Kigen = if Body = null then #table(LotType, {}) else
        let
            names = Table.ColumnNames(Body),
            pick  = (c as text) as list =>
                        if List.Contains(names, c) then Table.Column(Body, c) else List.Repeat({null}, Table.RowCount(Body)),
            qty   = if List.Contains(names, "期限切れ数量") then pick("期限切れ数量") else pick("数量"),
            t     = Table.FromColumns({pick("品名"), pick("ロット番号"), pick("期限"), qty, pick("単位")},
                        {"品名", "ロット", "使用期限", "数量", "単位"}),
            split = Table.ExpandRecordColumn(Table.AddColumn(t, "m", each SplitMark([品名])), "m", {"名"}, {"薬品名"}),
            typed = Table.TransformColumns(split, {
                        {"ロット",   Txt, type nullable text}, {"使用期限", Dat, type nullable date},
                        {"数量",     Num, type nullable number}, {"単位", each if _ = null then null else Nz(_), type nullable text}}),
            withS = Table.AddColumn(Table.AddColumn(typed, "出どころ", (x) => "期限切れリスト", type text), "ファイル名", (x) => F[Name], type text)
        in
            Value.ReplaceType(Table.SelectColumns(withS, {"薬品名", "ロット", "使用期限", "数量", "単位", "出どころ", "ファイル名"}), LotType),

    // ---- 不動品リスト（㉑）のロット ----
    FudoSrc = #"不動品リスト",
    Fudo    = Value.ReplaceType(
                  Table.SelectColumns(
                      Table.AddColumn(Table.RenameColumns(FudoSrc, {{"期限日", "使用期限"}}), "出どころ", (x) => "不動品リスト", type text),
                      {"薬品名", "ロット", "使用期限", "数量", "単位", "出どころ", "ファイル名"}),
                  LotType),

    All   = Table.Combine({Kigen, Fudo}),
    Valid = Table.Distinct(
                Table.SelectRows(All, each [薬品名] <> null and [薬品名] <> "" and [数量] <> null and [数量] > 0),
                {"薬品名", "ロット"}),

    // ---- 直近の1日平均払出をつなぎ、期限の早いロットから割り当てる ----
    WithPd   = Table.ExpandTableColumn(
                   Table.NestedJoin(Valid, {"薬品名"}, Drugs, {"薬品名"}, "d", JoinKind.LeftOuter), "d", {"直近の1日平均払出"}),
    WithDays = Table.AddColumn(WithPd, "残り日数", each
                   if [使用期限] = null then null else Duration.Days([使用期限] - BaseDate), type nullable number),
    Alloc = (g as table) as table =>
        let
            s     = Table.Sort(g, {{"使用期限", Order.Ascending}, {"ロット", Order.Ascending}}),
            qs    = List.Buffer(s[数量]),
            idx   = Table.AddIndexColumn(s, "i", 0, 1),
            prior = Table.AddColumn(idx, "それより早いロットの数量", each List.Sum(List.FirstN(qs, [i])) ?? 0, type number),
            use   = Table.AddColumn(prior, "期限までに使い切れる量", each
                        if [使用期限] = null then [数量]
                        else
                            let cap = ([直近の1日平均払出] ?? 0) * List.Max({0, [残り日数] ?? 0})
                            in  List.Max({0, List.Min({[数量], cap - [それより早いロットの数量]})}),
                        type number),
            ret   = Table.AddColumn(use, "返す量", each [数量] - [期限までに使い切れる量], type number)
        in
            Table.RemoveColumns(ret, {"i"}),
    Grouped  = Table.Group(WithDays, {"薬品名"}, {{"t", each Alloc(_), type table}}),
    Combined = if Table.IsEmpty(Grouped) then Empty else Table.Combine(Grouped[t]),
    Out      = Table.SelectColumns(Combined, {"薬品名", "ロット", "使用期限", "数量", "単位", "出どころ", "残り日数",
                   "それより早いロットの数量", "期限までに使い切れる量", "返す量", "ファイル名"}),
    Result   = Table.Sort(Value.ReplaceType(Out, OutType), {{"薬品名", Order.Ascending}, {"使用期限", Order.Ascending}})
in
    Result
