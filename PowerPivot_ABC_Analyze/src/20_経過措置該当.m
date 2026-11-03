// ================================================================================
// ⑳ クエリ名：経過措置該当（接続のみ）
//    エリシアSの「経過措置薬品リスト.xls」（設定「経過措置フォルダパス」の中で一番新しいファイル）を読む
//    ・見出しは「品名」のある行（実物は 1 行目）。列：記号／製剤コード／製剤／コード／品名／メーカーコード／
//      メーカー／経過措置（期限）／在庫数量／単位。記号 × は在庫 0
//    ・品名の先頭の記号（「劇 」「向 」など 1 文字＋空白）と先頭の空白を外して 薬品名 にする（⑤ 出納と同じ）
//    ・フォルダパスが空欄、またはファイルがなければ空の表（⑨ は手入力の経過措置リストだけで動く）
//    ・⑨ M_薬品 では、このリストを手入力の経過措置リストより優先する
// ================================================================================
let
    Prm  = #"設定値",
    Nz   = #"fx_正規化",
    Norm = (t as any) as text =>
        let s = if t = null then "" else Nz(Text.From(t))
        in  Text.Replace(Text.Replace(Text.Remove(s, {" ", "　", "#(lf)", "#(cr)", "#(tab)"}), "(", "（"), ")", "）"),
    CleanPath = (v as any) as nullable text =>
        if v = null or (v is text and Text.Trim(v) = "") then null
        else Text.TrimEnd(Text.Trim(Text.Remove(Text.From(v), {""""})), {"\", "/", " ", "　"}),
    Num = (val as any) as nullable number =>
        if val = null then null
        else if val is number then val
        else let txt = Text.Remove(Nz(Text.From(val)), {" ", ","})
             in  if txt = "" then null else try Number.FromText(txt, "en-US") otherwise null,

    // 品名の先頭の記号（劇 ／向 など）と先頭の空白を外す
    Marks = {"劇", "向", "毒", "麻", "覚", "習", "生"},
    SplitMark = (v as any) as record =>
        let s = if v = null then null else Nz(Text.From(v))
        in  if s = null or s = "" then [名 = null, 記号 = null]
            else if Text.Length(s) >= 2 and List.Contains(Marks, Text.Start(s, 1)) and Text.Middle(s, 1, 1) = " "
                 then [名 = Text.Trim(Text.Range(s, 2)), 記号 = Text.Start(s, 1)]
            else [名 = s, 記号 = null],

    // フォルダの中で、名前に marker を含む一番新しいファイル
    NewestFile = (folder as nullable text, marker as text) as nullable record =>
        if folder = null then null
        else
            let
                fs  = Table.SelectRows(Folder.Files(folder), each
                          Text.StartsWith(Text.Lower([Extension]), ".xls") and not Text.StartsWith([Name], "~$")
                          and Text.Contains(Nz([Name]), marker)),
                srt = Table.Sort(fs, {{"Date modified", Order.Descending}, {"Name", Order.Descending}})
            in
                if Table.IsEmpty(srt) then null else srt{0},
    ReadTable = (f as record) as table =>
        let
            wb     = Excel.Workbook(f[Content], null, true),
            sheets = if Table.HasColumns(wb, "Kind") then Table.SelectRows(wb, each [Kind] = "Sheet")
                     else Table.SelectRows(wb, each
                              let n = Text.TrimEnd(Text.Remove(Text.From([Name]), {"'"}), "$")
                              in  not Text.Contains(n, "$") and not Text.Contains(n, "Print_") and not Text.Contains(n, "_xlnm")),
            raw    = sheets{0}[Data],
            rows   = Table.ToRows(Table.FirstN(raw, 30)),
            hIdx   = List.PositionOf(List.Transform(rows, (r) => List.Contains(List.Transform(r, Norm), "品名")), true)
        in
            Table.TransformColumnNames(
                Table.PromoteHeaders(Table.Skip(raw, List.Max({hIdx, 0})), [PromoteAllScalars = true]), Norm),

    OutType = type table [薬品名 = text, 記号 = nullable text, 経過措置期限 = nullable date, 在庫数量 = nullable number,
                          単位 = nullable text, コード = nullable text, ファイル名 = nullable text],
    Empty   = #table(OutType, {}),

    Path = CleanPath(Record.FieldOrDefault(Prm, "経過措置フォルダパス", null)),
    F    = NewestFile(Path, "経過措置"),
    Body = if F = null then null else ReadTable(F),
    Result = if Body = null then Empty else
        let
            names = Table.ColumnNames(Body),
            pick  = (c as text) as list =>
                        if List.Contains(names, c) then Table.Column(Body, c) else List.Repeat({null}, Table.RowCount(Body)),
            t     = Table.FromColumns({pick("品名"), pick("経過措置"), pick("在庫数量"), pick("単位"), pick("コード")},
                        {"品名", "経過措置", "在庫数量", "単位", "コード"}),
            split = Table.ExpandRecordColumn(Table.AddColumn(t, "m", each SplitMark([品名])), "m", {"名", "記号"}, {"薬品名", "記号"}),
            typed = Table.TransformColumns(split, {
                        {"経過措置", each try Date.From(_) otherwise null, type nullable date},
                        {"在庫数量", Num, type nullable number},
                        {"単位",     each if _ = null then null else Nz(_), type nullable text},
                        {"コード",   each if _ = null then null else Text.Trim(Text.From(_)), type nullable text}}),
            rows  = Table.SelectRows(typed, each [薬品名] <> null and [薬品名] <> ""),
            withF = Table.AddColumn(rows, "ファイル名", (x) => F[Name], type text),
            ren   = Table.RenameColumns(withF, {{"経過措置", "経過措置期限"}}),
            sel   = Table.SelectColumns(ren, {"薬品名", "記号", "経過措置期限", "在庫数量", "単位", "コード", "ファイル名"}),
            dist  = Table.Distinct(sel, {"薬品名"})
        in
            Value.ReplaceType(dist, OutType)
in
    Result
