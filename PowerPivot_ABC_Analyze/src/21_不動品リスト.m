// ================================================================================
// ㉑ クエリ名：不動品リスト（接続のみ）
//    エリシアSの「YYYYMM_不動品リスト.xls」（設定「不動品フォルダパス」の中で一番新しいファイル）を読む
//    ・仕入伝票の行（＝ロット）が 1 行。見出しは「品名」のある行（実物は 1 行目）
//    ・使う列：品名／ロット番号／期限日／最終払出日／不動日数／数量／単位／棚番／薬価／YJコード／JANコード／基準日／仕入先
//    ・品名の先頭の記号（「劇 」「向 」など）と先頭の空白を外して 薬品名 にする（⑤ 出納と同じ）
//    ・フォルダパスが空欄、またはファイルがなければ空の表
//    ・使い先：⑨ M_薬品（薬ごとにまとめて システム不動日数・最終払出日・薬価・YJコード）、⑭ T_ロット（ロットの期限）
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
    Txt = (val as any) as nullable text =>
        if val = null then null else let s = Text.Trim(Text.From(val)) in if s = "" then null else s,
    Dat = (val as any) as nullable date => try Date.From(val) otherwise null,

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

    OutType = type table [薬品名 = text, 記号 = nullable text, ロット = nullable text, 期限日 = nullable date, 最終払出日 = nullable date,
                          不動日数 = nullable number, 数量 = nullable number, 単位 = nullable text, 棚番 = nullable text,
                          薬価 = nullable number, YJコード = nullable text, JANコード = nullable text, 基準日 = nullable date,
                          仕入先 = nullable text, ファイル名 = nullable text],
    Empty   = #table(OutType, {}),

    Path = CleanPath(Record.FieldOrDefault(Prm, "不動品フォルダパス", null)),
    F    = NewestFile(Path, "不動品"),
    Body = if F = null then null else ReadTable(F),
    Result = if Body = null then Empty else
        let
            names = Table.ColumnNames(Body),
            pick  = (c as text) as list =>
                        if List.Contains(names, c) then Table.Column(Body, c) else List.Repeat({null}, Table.RowCount(Body)),
            cols  = {"品名", "ロット番号", "期限日", "最終払出日", "不動日数", "数量", "単位", "棚番", "薬価", "YJコード", "JANコード", "基準日", "仕入先"},
            t     = Table.FromColumns(List.Transform(cols, pick), cols),
            split = Table.ExpandRecordColumn(Table.AddColumn(t, "m", each SplitMark([品名])), "m", {"名", "記号"}, {"薬品名", "記号"}),
            typed = Table.TransformColumns(split, {
                        {"ロット番号", Txt, type nullable text},
                        {"期限日",     Dat, type nullable date},
                        {"最終払出日", Dat, type nullable date},
                        {"不動日数",   Num, type nullable number},
                        {"数量",       Num, type nullable number},
                        {"単位",       each if _ = null then null else Nz(_), type nullable text},
                        {"棚番",       Txt, type nullable text},
                        {"薬価",       Num, type nullable number},
                        {"YJコード",   Txt, type nullable text},
                        {"JANコード",  each if _ = null then null else (if _ is number then Number.ToText(_, "F0") else Text.Trim(Text.From(_))), type nullable text},
                        {"基準日",     Dat, type nullable date},
                        {"仕入先",     each if _ = null then null else Nz(_), type nullable text}}),
            rows  = Table.SelectRows(typed, each [薬品名] <> null and [薬品名] <> ""),
            withF = Table.AddColumn(rows, "ファイル名", (x) => F[Name], type text),
            ren   = Table.RenameColumns(withF, {{"ロット番号", "ロット"}}),
            sel   = Table.SelectColumns(ren, {"薬品名", "記号", "ロット", "期限日", "最終払出日", "不動日数", "数量", "単位", "棚番",
                        "薬価", "YJコード", "JANコード", "基準日", "仕入先", "ファイル名"})
        in
            Value.ReplaceType(sel, OutType)
in
    Result
