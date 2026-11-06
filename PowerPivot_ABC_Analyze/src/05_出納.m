// ================================================================================
// ⑤ クエリ名：出納（接続のみ）
//    出納フォルダの月次の出納（数量一覧 .xls）を全期間まとめて読む
//    ・期間（期間開始・期間終了）は、正確な方から順に次で決める
//        1. ファイルの中：見出し行（「品名」のある行）より上に書かれた日付2つ
//           （例「期間：2026/10/01～2026/10/31」。日付のセル・yyyy/mm/dd・yyyy-mm-dd・
//             yyyy年mm月dd日・8桁の数字のどれでも読む。出力日などが混ざっていても、
//             「前の日付 ≦ 後の日付」になる最初の組を期間とする）
//        2. ファイル名：8桁の数字2つ（yyyymmdd_yyyymmdd_数量一覧.xls）
//      期間が分からないファイルは読まない。同じ期間のファイルは更新日時が新しい1つだけ
//    ・列（固定）：製剤コード／製剤／コード／品名／繰越数量／仕入数量／払出数量／在庫数量／単位／棚番
//    ・出納に載るのは、その月に何かしら動いた薬だけ（動きのない薬は載らない）
//    ・その他の増減 ＝ 在庫数量 －（繰越数量 ＋ 仕入数量 － 払出数量）
//        マイナス：返品・廃棄・店舗間移動・譲渡など　プラス：移動の受け入れ・棚卸の調整など
//    ・品名の記号：品名の先頭の「劇 」「向 」など（劇薬・向精神薬の印。記号がない薬は先頭が空白 3 つ）。
//      照合のため 品名 からは記号と先頭の空白を外す。品名の先頭空白はその空白の数（記号がない薬で 3）
// ================================================================================
let
    Prm  = #"設定値",
    Nz   = #"fx_正規化",

    // ---- フォルダ（設定「出納フォルダパス」）とファイル名の目印（① で固定「数量一覧」） ----
    FolderPath = if Prm[出納フォルダパス] = null then error "設定テーブルの「出納フォルダパス」が空欄です"
                 else Prm[出納フォルダパス],
    Mark       = Prm[出納の目印],

    // 列名の照合：表記をそろえ、空白・改行を除き、かっこを全角にそろえる
    Norm = (t as any) as text =>
        let s = if t = null then "" else Nz(Text.From(t))
        in  Text.Replace(Text.Replace(Text.Remove(s, {" ", "　", "#(lf)", "#(cr)", "#(tab)"}), "(", "（"), ")", "）"),

    Cols = {"製剤コード", "製剤", "コード", "品名", "繰越数量", "仕入数量", "払出数量", "在庫数量", "単位", "棚番"},
    OutType = type table [期間開始 = date, 期間終了 = date, 日数 = Int64.Type,
                          製剤コード = nullable text, 製剤 = nullable text, コード = nullable text,
                          品名 = text, 品名の先頭空白 = Int64.Type, 品名の記号 = nullable text,
                          繰越数量 = nullable number, 仕入数量 = nullable number, 払出数量 = nullable number,
                          在庫数量 = nullable number, その他の増減 = nullable number,
                          単位 = nullable text, 棚番 = nullable text, ファイル名 = text],

    // ---- 日付の読み取り ----
    // 8桁の数字 → 日付（1990～2100年だけ）
    ToDate8 = (x as text) as nullable date =>
        let d = try #date(Number.FromText(Text.Start(x, 4)), Number.FromText(Text.Middle(x, 4, 2)), Number.FromText(Text.End(x, 2)))
                otherwise null
        in  if d = null or Date.Year(d) < 1990 or Date.Year(d) > 2100 then null else d,
    // 「2026/10/01」「2026-10-01」「2026.10.1」「2026年10月1日」→ 日付
    ToDateYmd = (t as text) as nullable date =>
        let
            p = Text.Split(t, "/"),
            d = if List.Count(p) = 3 then (try #date(Number.FromText(p{0}), Number.FromText(p{1}), Number.FromText(p{2})) otherwise null)
                else if List.Count(p) = 1 and Text.Length(t) = 8 then ToDate8(t)
                else null
        in  if d = null or Date.Year(d) < 1990 or Date.Year(d) > 2100 then null else d,
    // 1つのセルの値から、含まれる日付をすべて取り出す（文字の中に複数あってもよい）
    DatesIn = (v as any) as list =>
        if v = null then {}
        else if v is date then {v}
        else if v is datetime then {Date.From(v)}
        else if v is number then List.RemoveNulls({try ToDate8(Number.ToText(v, "D")) otherwise null})
        else
            let
                s     = Nz(Text.From(v)),
                s2    = Text.Replace(Text.Replace(Text.Replace(Text.Replace(Text.Replace(s,
                            "年", "/"), "月", "/"), "日", " "), "-", "/"), ".", "/"),
                chars = List.Transform(Text.ToList(s2), (c) => if List.Contains({"0" .. "9", "/"}, c) then c else " "),
                toks  = List.Select(Text.Split(Text.Combine(chars), " "), (x) => x <> "")
            in
                List.RemoveNulls(List.Transform(toks, ToDateYmd)),
    // 日付の並びから、「前 ≦ 後」になる最初の2つを期間として返す
    PairOf = (ds as list) as nullable record =>
        let
            n    = List.Count(ds),
            idx  = List.Select({0 .. n - 2}, (i) => ds{i} <= ds{i + 1}),
            i0   = if List.IsEmpty(idx) then null else List.First(idx)
        in  if n < 2 or i0 = null then null else [開始 = ds{i0}, 終了 = ds{i0 + 1}],

    // ---- 期間 1：ファイルの中（見出し行より上の日付） ----
    SheetsOf = (wb as table) as table =>
        if Table.HasColumns(wb, "Kind") then Table.SelectRows(wb, each [Kind] = "Sheet")
        else Table.SelectRows(wb, each
                 let n = Text.TrimEnd(Text.Remove(Text.From([Name]), {"'"}), "$")
                 in  not Text.Contains(n, "$") and not Text.Contains(n, "Print_") and not Text.Contains(n, "_xlnm")),
    HeaderIndex = (rows as list) as number =>
        List.PositionOf(List.Transform(rows, (r) => List.Contains(List.Transform(r, Norm), "品名")), true),
    PeriodInFile = (content as binary) as nullable record =>
        try
            let
                raw  = SheetsOf(Excel.Workbook(content, null, true)){0}[Data],
                rows = Table.ToRows(Table.FirstN(raw, 30)),
                hIdx = HeaderIndex(rows),
                top  = if hIdx <= 0 then {} else List.FirstN(rows, hIdx),
                ds   = List.Combine(List.Transform(List.Combine(top), DatesIn))
            in
                PairOf(ds)
        otherwise null,

    // ---- 期間 2：ファイル名の 8桁の数字2つ ----
    PeriodInName = (nm as text) as nullable record =>
        let
            parts = Text.SplitAny(Nz(nm), "_-. （）()"),
            ds    = List.Select(parts, (p) => Text.Length(p) = 8 and Text.Select(p, {"0" .. "9"}) = p),
            a     = if List.Count(ds) >= 2 then ToDate8(ds{0}) else null,
            b     = if List.Count(ds) >= 2 then ToDate8(ds{1}) else null
        in
            if a = null or b = null or b < a then null else [開始 = a, 終了 = b],

    Files  = Table.SelectRows(Folder.Files(FolderPath), each
                 Text.StartsWith(Text.Lower([Extension]), ".xls") and not Text.StartsWith([Name], "~$")
                 and Text.Contains(Nz([Name]), Mark)),
    WithP  = Table.AddColumn(Files, "期間", each PeriodInFile([Content]) ?? PeriodInName([Name])),
    Valid  = Table.SelectRows(WithP, each [期間] <> null),
    Keyed  = Table.AddColumn(Valid, "期間キー", each
                 Date.ToText([期間][開始], "yyyyMMdd") & "-" & Date.ToText([期間][終了], "yyyyMMdd"), type text),
    // 同じ期間のファイルは、更新日時が新しい1つだけ
    Latest = Table.Group(Keyed, {"期間キー"}, {{"f", each Table.First(Table.Sort(_, {{"Date modified", Order.Descending}}))}}),

    // ---- 1ファイルを読む ----
    ReadOne = (f as record) as table =>
        let
            raw    = SheetsOf(Excel.Workbook(f[Content], null, true)){0}[Data],
            hIdx   = HeaderIndex(Table.ToRows(raw)),
            body   = Table.TransformColumnNames(
                         Table.PromoteHeaders(Table.Skip(raw, List.Max({hIdx, 0})), [PromoteAllScalars = true]), Norm),
            names  = Table.ColumnNames(body),
            pick   = (c as text) as list =>
                         if List.Contains(names, c) then Table.Column(body, c) else List.Repeat({null}, Table.RowCount(body)),
            t      = Table.FromColumns(List.Transform(Cols, pick), Cols),
            p      = f[期間],
            fname  = f[Name]
        in
            Table.AddColumn(Table.AddColumn(Table.AddColumn(t,
                "期間開始", (x) => p[開始], type date),
                "期間終了", (x) => p[終了], type date),
                "ファイル名", (x) => fname, type text),

    Combined = if Table.IsEmpty(Latest) then #table(Cols & {"期間開始", "期間終了", "ファイル名"}, {})
               else Table.Combine(List.Transform(Latest[f], ReadOne)),

    // ---- 型と表記 ----
    Num = (val as any) as nullable number =>
        if val = null then null
        else if val is number then val
        else let txt = Text.Remove(Nz(Text.From(val)), {" ", ","})
             in  if txt = "" then null else try Number.FromText(txt, "en-US") otherwise null,
    Txt = (val as any) as nullable text =>
        if val = null then null else let s = Text.Trim(Text.From(val)) in if s = "" then null else s,
    Lead = (val as any) as number =>
        if val = null then 0
        else let s = Text.From(val) in Text.Length(s) - Text.Length(Text.TrimStart(s, {" ", "　"})),
    // 品名の先頭の記号（劇 ／向 など 1 文字＋空白）を外す
    Marks = {"劇", "向", "毒", "麻", "覚", "習", "生"},
    MarkOf = (val as any) as nullable text =>
        let s = if val = null then null else Nz(Text.From(val))
        in  if s <> null and Text.Length(s) >= 2 and List.Contains(Marks, Text.Start(s, 1)) and Text.Middle(s, 1, 1) = " "
            then Text.Start(s, 1) else null,
    NameOf = (val as any) as nullable text =>
        let s = if val = null then null else Nz(Text.From(val))
        in  if s = null then null
            else if MarkOf(s) <> null then (let t = Text.Trim(Text.Range(s, 2)) in if t = "" then null else t)
            else s,

    WithLead = Table.AddColumn(Table.AddColumn(Combined, "品名の先頭空白", each Lead([品名]), Int64.Type),
                   "品名の記号", each MarkOf([品名]), type nullable text),
    Typed = Table.TransformColumns(WithLead, {
                {"製剤コード", Txt, type nullable text},
                {"製剤",       each if _ = null then null else Nz(_), type nullable text},
                {"コード",     Txt, type nullable text},
                {"品名",       NameOf, type nullable text},
                {"繰越数量",   Num, type nullable number},
                {"仕入数量",   Num, type nullable number},
                {"払出数量",   Num, type nullable number},
                {"在庫数量",   Num, type nullable number},
                {"単位",       each if _ = null then null else Nz(_), type nullable text},
                {"棚番",       Txt, type nullable text}}),
    Rows = Table.SelectRows(Typed, each [品名] <> null and [品名] <> ""
                 and not List.Contains({"合計", "総合計", "小計", "品名"}, Norm([品名]))),

    // その他の増減 ＝ 在庫 －（繰越 ＋ 仕入 － 払出）。小数の誤差は 0.001 単位で丸める
    Diff = Table.AddColumn(Rows, "その他の増減", each
               if [在庫数量] = null then null
               else Number.Round([在庫数量] - (([繰越数量] ?? 0) + ([仕入数量] ?? 0) - ([払出数量] ?? 0)), 3),
               type nullable number),
    Days = Table.AddColumn(Diff, "日数", each Duration.Days([期間終了] - [期間開始]) + 1, Int64.Type),

    Ordered = Table.SelectColumns(Days, {"期間開始", "期間終了", "日数", "製剤コード", "製剤", "コード", "品名",
                  "品名の先頭空白", "品名の記号", "繰越数量", "仕入数量", "払出数量", "在庫数量", "その他の増減", "単位", "棚番", "ファイル名"}),
    Result  = Value.ReplaceType(Ordered, OutType)
in
    Result
