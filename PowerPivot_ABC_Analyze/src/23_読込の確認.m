// ================================================================================
// ㉓ クエリ名：読込の確認（シート「読込の確認」に表として読み込む）
//    設定のフォルダにあるファイルを 1 つずつ開いてみて、読めるか・読めないならなぜかを一覧にする
//    数字が出ない・少ないときは、まずこの表を見る
//    ・1 行目：実績を読む期間（① の開始日～終了日）
//    ・フォルダごと（実績／出納／期限切れリスト／経過措置薬品リスト／不動品リスト）に、ファイル 1 つで 1 行
//    ・列
//        種類      ：どのフォルダか
//        ファイル名
//        更新日時
//        形式      ：中身を見て決めた形式（Excel（.xls）／Excel（.xlsx）／文字（カンマ区切り・Shift_JIS）など）
//        シート数  ：Excel ブックのシートの数（文字のファイルは 1）
//        見出しの行：見出し（実績は「薬品名」、ほかは「品名」）が何行目にあったか
//        行数      ：見出しのある表（なければ 1 枚目）の行数
//        結果      ：「読める…」「読まない：…」（対象外。理由つき）「読めない：…」（開けない・見出しがない。理由つき）
//    ・3 つのリスト（期限切れ・経過措置・不動品）は、フォルダの中で一番新しい 1 つだけを読む
//    ・出納の期間は、ファイルの中の日付（見出しより上）が優先。ここではファイル名の期間だけを見ている
// ================================================================================
let
    Prm  = #"設定値",
    Nz   = #"fx_正規化",
    Open = #"fx_ファイルを開く",
    Norm = (t as any) as text =>
        let s = if t = null then "" else Nz(Text.From(t))
        in  Text.Replace(Text.Replace(Text.Remove(s, {" ", "　", "#(lf)", "#(cr)", "#(tab)"}), "(", "（"), ")", "）"),
    CleanPath = (v as any) as nullable text =>
        if v = null or (v is text and Text.Trim(v) = "") then null
        else Text.TrimEnd(Text.Trim(Text.Remove(Text.From(v), {""""})), {"\", "/", " ", "　"}),
    S0 = Date.From(Prm[開始日]),
    E0 = Date.From(Prm[終了日]),
    Ymd = (d as date) as text => Date.ToText(d, "yyyy/MM/dd"),
    ExtOk = (e as text) as logical =>
        Text.StartsWith(Text.Lower(e), ".xls") or List.Contains({".csv", ".txt", ".tsv"}, Text.Lower(e)),

    KindPlan = "実績（薬品使用量）",
    KindSui  = "出納（数量一覧）",
    Targets = {
        [種類 = KindPlan, フォルダ = CleanPath(Prm[フォルダパス]), 目印 = Prm[先頭], 先頭一致 = true,
         見出し = Norm(Prm[列名][薬品名]), 最新だけ = false],
        [種類 = KindSui, フォルダ = CleanPath(Prm[出納フォルダパス]), 目印 = Prm[出納の目印], 先頭一致 = false,
         見出し = "品名", 最新だけ = false],
        [種類 = "期限切れリスト", フォルダ = CleanPath(Record.FieldOrDefault(Prm, "期限切迫フォルダパス", null)),
         目印 = "期限切れ", 先頭一致 = false, 見出し = "品名", 最新だけ = true],
        [種類 = "経過措置薬品リスト", フォルダ = CleanPath(Record.FieldOrDefault(Prm, "経過措置フォルダパス", null)),
         目印 = "経過措置", 先頭一致 = false, 見出し = "品名", 最新だけ = true],
        [種類 = "不動品リスト", フォルダ = CleanPath(Record.FieldOrDefault(Prm, "不動品フォルダパス", null)),
         目印 = "不動品", 先頭一致 = false, 見出し = "品名", 最新だけ = true]},

    Row = (kind as text, nm as nullable text, modified as any, fmt as nullable text, nSheets as nullable number,
           hRow as nullable number, nRows as nullable number, result as text) as record =>
        [種類 = kind, ファイル名 = nm, 更新日時 = modified, 形式 = fmt, シート数 = nSheets,
         見出しの行 = hRow, 行数 = nRows, 結果 = result],

    // 8桁の数字 → 日付
    D8 = (x as text) as nullable date =>
        try #date(Number.FromText(Text.Start(x, 4)), Number.FromText(Text.Middle(x, 4, 2)), Number.FromText(Text.End(x, 2)))
        otherwise null,

    // 実績：ファイル名の年月が読む期間に入っているか、日付をどう決めるか
    PlanNote = (nm as text, sheets as table, isText as logical, hasDay as logical) as text =>
        let
            body  = Text.BeforeDelimiter(Text.AfterDelimiter(nm, Prm[先頭]), ".", {0, RelativePosition.FromEnd}),
            chars = Text.ToList(Nz(body)),
            from  = List.PositionOfAny(chars, {"0" .. "9"}),
            dg    = if from < 0 then ""
                    else Text.Combine(List.FirstN(List.Skip(chars, from), each List.Contains({"0" .. "9"}, _))),
            ym    = if Text.Length(dg) >= 6 then D8(Text.Start(dg, 6) & "01") else null,
            inP   = ym <> null and ym >= Date.AddMonths(Date.StartOfMonth(S0), -1) and ym <= E0,
            days  = List.Count(List.Select(sheets[シート名], each
                        _ <> null and Text.Contains(_, "月") and Text.Contains(_, "日") and not Text.StartsWith(Nz(_), "期間")))
        in
            if ym = null then "読まない：ファイル名に年月（yyyymm）がない"
            else if not inP then "読まない：ファイルの年月（" & Date.ToText(ym, "yyyy/MM") & "）が読む期間の外"
            else if isText and hasDay then "読める（「日付」の列で日ごとに分ける）"
            else if isText and Text.Length(dg) >= 8 then "読める（ファイル名の日付の 1 日分）"
            else if isText then "読まない：「日付」の列がなく、ファイル名に 8 桁の日付（yyyymmdd）もない"
            else if days = 0 then "読めない：「10月1日(水)」のような日ごとのシートがない"
            else "読める（日ごとのシート " & Text.From(days) & " 枚）",

    // 出納：ファイル名の期間（8桁の数字2つ）が読む期間にかかるか
    SuiNote = (nm as text) as text =>
        let
            parts = Text.SplitAny(Nz(nm), "_-. （）()"),
            ds    = List.Select(parts, (p) => Text.Length(p) = 8 and Text.Select(p, {"0" .. "9"}) = p),
            a     = if List.Count(ds) >= 2 then D8(ds{0}) else null,
            b     = if List.Count(ds) >= 2 then D8(ds{1}) else null
        in
            if a = null or b = null then "読める（ファイル名に期間がない。ファイルの中の日付で期間を決める。日付がなければ読まない）"
            else if b < S0 or a > E0 then "読まない：ファイル名の期間（" & Ymd(a) & "～" & Ymd(b) & "）が読む期間の外"
            else "読める（ファイル名の期間 " & Ymd(a) & "～" & Ymd(b) & "）",

    // ファイル 1 つを調べる
    Check = (t as record, f as record, newest as nullable text) as record =>
        let
            nm     = f[Name],
            nameOk = if t[先頭一致] then Text.StartsWith(nm, t[目印]) else Text.Contains(Nz(nm), t[目印]),
            extOk  = ExtOk(f[Extension]),
            old    = nameOk and extOk and newest <> null and nm <> newest,
            opened = if nameOk and extOk and not old then try Open(f[Content], nm) else null,
            failed = if opened = null then false else opened[HasError],
            sheets = if opened = null or failed then null else opened[Value],
            // .xls はシートが名前の順に並ぶので、見出しのある表を探して調べる（なければ 1 枚目）
            HeadOf = (tbl as table) as number =>
                List.PositionOf(List.Transform(Table.ToRows(Table.FirstN(tbl, 30)),
                    (r) => List.Contains(List.Transform(r, Norm), t[見出し])), true),
            withH  = if sheets = null then null
                     else Table.SelectRows(sheets, (s) => (try HeadOf(s[Data]) otherwise -1) >= 0),
            first  = if sheets = null or Table.IsEmpty(sheets) then null
                     else if Table.IsEmpty(withH) then sheets{0}[Data] else withH{0}[Data],
            top    = if first = null then {} else Table.ToRows(Table.FirstN(first, 30)),
            hIdx   = List.PositionOf(List.Transform(top, (r) => List.Contains(List.Transform(r, Norm), t[見出し])), true),
            isText = if first = null then false else Text.StartsWith(sheets{0}[形式], "文字"),
            hasDay = if hIdx < 0 then false else List.Contains(List.Transform(top{hIdx}, Norm), "日付"),
            result = if not nameOk then
                         "読まない：ファイル名" & (if t[先頭一致] then "が「" & t[目印] & "」で始まらない" else "に「" & t[目印] & "」がない")
                     else if not extOk then "読まない：拡張子が対象外（読めるのは .xls .xlsx .csv .txt）"
                     else if old then "読まない：もっと新しいファイルがある（一番新しい 1 つだけ読む）"
                     else if t[種類] = KindSui and Text.StartsWith(SuiNote(nm), "読まない") then SuiNote(nm)
                     else if failed then "読めない：" & opened[Error][Message]
                     else if first = null then "読めない：シートが見つからない"
                     else if hIdx < 0 then "読めない：見出し「" & t[見出し] & "」のある行が、先頭 30 行にない"
                     else if t[種類] = KindPlan then PlanNote(nm, sheets, isText, hasDay)
                     else if t[種類] = KindSui then SuiNote(nm)
                     else "読める"
        in
            Row(t[種類], nm, f[Date modified],
                if first = null then null else sheets{0}[形式],
                if sheets = null then null else Table.RowCount(sheets),
                if hIdx < 0 then null else hIdx + 1,
                if first = null then null else Table.RowCount(first),
                result),

    // フォルダ 1 つを調べる
    One = (t as record) as list =>
        if t[フォルダ] = null then {Row(t[種類], null, null, null, null, null, null, "設定のフォルダパスが空欄（読まない）")}
        else
            let
                tried = try Table.RowCount(Folder.Files(t[フォルダ])),
                files = Table.Sort(Table.SelectRows(Folder.Files(t[フォルダ]), each not Text.StartsWith([Name], "~$")),
                            {{"Name", Order.Ascending}}),
                cand  = Table.Sort(
                            Table.SelectRows(files, each
                                (if t[先頭一致] then Text.StartsWith([Name], t[目印]) else Text.Contains(Nz([Name]), t[目印]))
                                and ExtOk([Extension])),
                            {{"Date modified", Order.Descending}, {"Name", Order.Descending}}),
                newest = if t[最新だけ] and not Table.IsEmpty(cand) then cand{0}[Name] else null
            in
                if tried[HasError] then
                    {Row(t[種類], null, null, null, null, null, null,
                         "フォルダを開けない：" & t[フォルダ] & "（" & tried[Error][Message] & "）")}
                else if Table.IsEmpty(files) then
                    {Row(t[種類], null, null, null, null, null, null, "フォルダにファイルがない：" & t[フォルダ])}
                else if Table.IsEmpty(cand) then
                    {Row(t[種類], null, null, null, null, null, null,
                         "読めるファイルがない：名前に「" & t[目印] & "」が付く .xls .xlsx .csv .txt が、フォルダ " & t[フォルダ] & " にない")}
                    & Table.TransformRows(files, (f) => Check(t, f, newest))
                else Table.TransformRows(files, (f) => Check(t, f, newest)),

    Period = Row("読む期間", null, null, null, null, null, null,
                 Ymd(S0) & " ～ " & Ymd(E0) & "（発注日 " & Ymd(Date.From(Prm[発注日])) & "）"),
    All    = {Period} & List.Combine(List.Transform(Targets, One)),
    T      = Table.FromRecords(All, type table [種類 = text, ファイル名 = nullable text, 更新日時 = nullable datetime,
                 形式 = nullable text, シート数 = nullable number, 見出しの行 = nullable number, 行数 = nullable number,
                 結果 = text]),
    Result = Table.TransformColumnTypes(T, {{"シート数", Int64.Type}, {"見出しの行", Int64.Type}, {"行数", Int64.Type}})
in
    Result
