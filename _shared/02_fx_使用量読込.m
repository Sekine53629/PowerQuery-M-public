// ================================================================================
// ② クエリ名：fx_使用量読込
//    エリシアSの出力ファイル（.xls／.xlsx／.csv／.txt）を読む関数（実績・予測の両方で使う）
//      folderPath ：読むフォルダ
//      prefix     ：ファイル名の先頭（① で固定「薬品使用量」。分析データ・予測データどちらも読める）
//      startDate, endDate：読む期間
//      cols       ：① の [列名]
//      keepZero   ：true なら使用量0の行も残す（在庫だけある薬を見るとき。省略時 false）
//    出力列：日付／薬品区分／略称／薬品名／患者数／数量／在庫量／在庫量_発注残込み／ファイル名／シート名
//
//    ファイルは ㉒ fx_ファイルを開く で開く（.xls の癖への対策と、文字のファイルの読み方はそちら）
//      ・Excel ブック：シート名「10月1日(水)」とファイル名の年月から日付を決める
//      ・文字のファイル（.csv／.txt。シートがない）：次のどちらかで日付を決める
//          1. 見出しに「日付」の列があれば、その列の日付ごとに分ける（1 ファイルに何日分あってもよい）
//          2. なければ、ファイル名の 8 桁の日付（yyyymmdd）の 1 日分として読む（1 日 1 ファイル）
//        どちらでもないファイルは読まない（㉓ 読込の確認 に出る）
//      ・見出しは「薬品名」と書かれた行を探す。2行目の空欄（1行目と結合）は自動で読み飛ばす
//    ファイル名の日付：先頭の後ろで最初に出てくる数字の並び（yyyymm または yyyymmdd）
//      （「_」「-」「空白」などの区切りがあってもよい）
//    薬品名・製剤は fx_正規化 でそろえる。略称は表示用にそのまま残す
//    略称の列：見出し「略称」→「略称」を含む見出し→ B列 の順に探す
// ================================================================================
(folderPath as text, prefix as text, startDate as date, endDate as date, cols as record,
 optional keepZero as nullable logical) as table =>
let
    Nz   = #"fx_正規化",
    Open = #"fx_ファイルを開く",

    // 列名の照合：表記をそろえ、空白を除き、かっこを全角にそろえる
    Norm = (t as any) as text =>
        let s = if t = null then "" else Nz(Text.From(t))
        in  Text.Replace(Text.Replace(Text.Remove(s, {" ", "　", "#(lf)", "#(cr)", "#(tab)"}), "(", "（"), ")", "）"),

    ColDrug  = Norm(cols[薬品名]),
    ColQty   = Norm(cols[数量]),
    ColStock = Norm(cols[在庫量]),
    ColIncl  = Norm(Record.FieldOrDefault(cols, "在庫量込み", "")),
    ColCat   = Norm(cols[薬品区分]),
    ColPat   = Norm(cols[患者数]),
    ColAbbr  = Norm(Record.FieldOrDefault(cols, "略称", "")),
    ColDate  = Norm(Record.FieldOrDefault(cols, "日付", "日付")),
    KeepZero = keepZero = true,

    OutType = type table [日付 = date, 薬品区分 = nullable text, 略称 = nullable text, 薬品名 = text,
                          患者数 = nullable number, 数量 = nullable number, 在庫量 = nullable number,
                          在庫量_発注残込み = nullable number, ファイル名 = text, シート名 = text],
    Empty   = #table(OutType, {}),

    ソース = Folder.Files(folderPath),
    対象   = Table.SelectRows(ソース, each
                 Text.StartsWith([Name], prefix)
                 and not Text.StartsWith([Name], "~$")
                 and (Text.StartsWith(Text.Lower([Extension]), ".xls")
                      or List.Contains({".csv", ".txt", ".tsv"}, Text.Lower([Extension])))),

    // 先頭の後ろで最初に出てくる数字の並び（全角数字も可）
    Digits = (nm as text) as text =>
        let
            s     = Nz(Text.BeforeDelimiter(Text.AfterDelimiter(nm, prefix), ".", {0, RelativePosition.FromEnd})),
            chars = Text.ToList(s),
            from  = List.PositionOfAny(chars, {"0" .. "9"}),
            rest  = if from < 0 then {} else List.Skip(chars, from),
            run   = List.FirstN(rest, each List.Contains({"0" .. "9"}, _))
        in
            Text.Combine(run),
    年月追加 = Table.AddColumn(対象, "ファイル年月", each
                 let d = Digits([Name])
                 in  if Text.Length(d) >= 6
                     then try #date(Number.FromText(Text.Start(d, 4)), Number.FromText(Text.Middle(d, 4, 2)), 1) otherwise null
                     else null,
                 type nullable date),
    // 新しさの順番：ファイル名の日付（yyyymm は yyyymm00）→ 更新日時 → ファイル名
    キー追加 = Table.AddColumn(年月追加, "並びキー", each
                 let
                     d  = Digits([Name]),
                     d8 = if Text.Length(d) >= 8 then Text.Start(d, 8)
                          else if Text.Length(d) >= 6 then Text.Start(d, 6) & "00" else "00000000",
                     tm = try DateTime.ToText(DateTime.From([Date modified]), "yyyyMMddHHmmss") otherwise "00000000000000"
                 in  d8 & "|" & tm & "|" & [Name], type text),

    // 月をまたぐ週は前月のファイルに入ることがあるので、1か月前から見る
    期間内ファイル = Table.SelectRows(キー追加, each
                 [ファイル年月] <> null
                 and [ファイル年月] >= Date.AddMonths(Date.StartOfMonth(startDate), -1)
                 and [ファイル年月] <= endDate),

    // 日付の読み取り（文字のファイル用）：日付のセル・yyyy/mm/dd・yyyy-mm-dd・yyyy年m月d日・8桁の数字
    ToDate = (val as any) as nullable date =>
        if val = null then null
        else if val is date then val
        else if val is datetime then Date.From(val)
        else
            let
                s0 = Nz(Text.From(val)),
                s  = Text.BeforeDelimiter(Text.Trim(Text.Replace(Text.Replace(Text.Replace(Text.Replace(Text.Replace(s0,
                         "年", "/"), "月", "/"), "日", ""), "-", "/"), ".", "/")), " "),
                p  = Text.Split(s, "/")
            in
                if List.Count(p) = 3 and Text.Length(p{0}) = 4
                then (try #date(Number.FromText(p{0}), Number.FromText(p{1}), Number.FromText(p{2})) otherwise null)
                else if Text.Length(s) = 8 and Text.Select(s, {"0" .. "9"}) = s
                then (try #date(Number.FromText(Text.Start(s, 4)), Number.FromText(Text.Middle(s, 4, 2)),
                                Number.FromText(Text.End(s, 2))) otherwise null)
                else try Date.From(val) otherwise null,

    // 文字のファイル 1 つを「日付ごとの表」に分ける
    //   見出しに「日付」の列があれば、その列で日付ごとに分ける。なければファイル名の 8 桁の日付の 1 日分
    SplitByDate = (raw as table, nm as text) as list =>
        let
            rows = Table.ToRows(Table.FirstN(raw, 30)),
            hIdx = List.PositionOf(
                       List.Transform(rows, (r) => List.Contains(List.Transform(r, Norm), ColDrug)), true),
            hdr  = if hIdx < 0 then {} else List.Transform(rows{hIdx}, Norm),
            dIdx = if ColDate = "" then -1 else List.PositionOf(hdr, ColDate),
            dg   = Digits(nm),
            d8   = if Text.Length(dg) >= 8 then ToDate(Text.Start(dg, 8)) else null
        in
            if hIdx >= 0 and dIdx >= 0 then
                let
                    dcol = Table.ColumnNames(raw){dIdx},
                    body = Table.AddColumn(Table.Skip(raw, hIdx + 1), "日付直接",
                               (r) => ToDate(Record.Field(r, dcol)), type nullable date),
                    inP  = Table.SelectRows(body, (r) =>
                               r[日付直接] <> null and r[日付直接] >= startDate and r[日付直接] <= endDate),
                    top  = Table.FirstN(Table.Skip(raw, hIdx), 1),
                    grp  = Table.Group(inP, {"日付直接"}, {{"Data", (t) =>
                               Table.Combine({top, Table.RemoveColumns(t, {"日付直接"})}), type table}})
                in
                    Table.ToRecords(grp)
            else
                {[日付直接 = d8, Data = raw]},

    // 見出しの位置は後で探すので、ここでは見出しにしない
    シート取得 = Table.AddColumn(期間内ファイル, "Sheets", each Open([Content], [Name])),
    列選択     = Table.SelectColumns(シート取得, {"Name", "ファイル年月", "並びキー", "Sheets"}),
    展開0      = Table.SelectRows(
                     Table.ExpandTableColumn(列選択, "Sheets", {"シート名", "Data", "形式"}, {"シート名", "Data", "形式"}),
                     each [形式] <> null),
    // 文字のファイル（シートがない）は日付ごとに分ける。Excel のシートはそのまま（日付はシート名から）
    分割       = Table.AddColumn(展開0, "日別", (row) =>
                     if Text.StartsWith(row[形式], "文字") then SplitByDate(row[Data], row[Name])
                     else {[日付直接 = null, Data = row[Data]]}),
    展開       = Table.ExpandRecordColumn(
                     Table.ExpandListColumn(Table.RemoveColumns(分割, {"Data"}), "日別"),
                     "日別", {"日付直接", "Data"}),

    名前整形 = Table.AddColumn(展開, "シート名2", each
                 if [シート名] = null then null else Nz([シート名]), type nullable text),

    // シート名「10月1日(水)」とファイル名の年月から日付を作る
    //   月がファイルと同じ：ファイルの年月　／　月がファイルの翌月：翌月（12月→1月は翌年）
    //   「期間」で始まるシート（期間すべて）と、それ以外の月は読まない
    日付追加 = Table.AddColumn(名前整形, "日付", each if [日付直接] <> null then [日付直接] else try
                 let
                     s     = [シート名2],
                     mPart = Text.Trim(Text.BeforeDelimiter(s, "月")),
                     isDay = s <> null and not Text.StartsWith(s, "期間")
                             and Text.Contains(s, "月") and Text.Contains(s, "日")
                             and mPart <> "" and Text.Select(mPart, {"0" .. "9"}) = mPart,
                     fy    = Date.Year([ファイル年月]),
                     fm    = Date.Month([ファイル年月]),
                     m     = Number.From(mPart),
                     dd    = Number.From(Text.Select(Text.BetweenDelimiters(s, "月", "日"), {"0" .. "9"})),
                     same  = m = fm,
                     next  = m = (if fm = 12 then 1 else fm + 1),
                     y     = if same then fy else if next and fm = 12 then fy + 1 else fy
                 in
                     if isDay and (same or next) then #date(y, m, dd) else null
                 otherwise null, type nullable date),
    日別のみ = Table.SelectRows(日付追加, each
                 [日付] <> null and [日付] >= startDate and [日付] <= endDate and not ([Data] is null)),

    // 同じ日付のシートが2つ以上あるときは、一番新しいファイルの1枚だけ残す
    最新キー = Table.Group(日別のみ, {"日付"}, {{"採用キー", each List.Max([並びキー]), type text}}),
    採用     = List.Buffer(Table.TransformRows(最新キー, each Date.ToText([日付], "yyyyMMdd") & "#" & [採用キー])),
    日別最新 = Table.Distinct(
                 Table.SelectRows(日別のみ, each List.Contains(採用, Date.ToText([日付], "yyyyMMdd") & "#" & [並びキー])),
                 {"日付"}),

    // 各シート：見出し行を探し、必要な列だけを取り出す
    整形 = Table.AddColumn(日別最新, "表", each
        let
            d    = [日付],
            f    = [Name],
            s    = [シート名2],
            raw  = [Data],
            rows = Table.ToRows(raw),
            hIdx = List.PositionOf(
                       List.Transform(rows, (r) => List.Contains(List.Transform(r, Norm), ColDrug)), true),
            body = if hIdx < 0 then null
                   else Table.TransformColumnNames(
                            Table.PromoteHeaders(Table.Skip(raw, hIdx), [PromoteAllScalars = true]), Norm),
            pick = (tbl as table, col as text) as list =>
                       if col <> "" and List.Contains(Table.ColumnNames(tbl), col)
                       then Table.Column(tbl, col) else List.Repeat({null}, Table.RowCount(tbl)),
            // 略称：見出し「略称」→「略称」を含む見出し→ B列（ほかの列に使っていなければ）
            abbr = (tbl as table) as list =>
                       let
                           names  = Table.ColumnNames(tbl),
                           byName = if ColAbbr <> "" and List.Contains(names, ColAbbr) then ColAbbr
                                    else List.First(List.Select(names, (n) => Text.Contains(n, "略称")), null),
                           byPos  = if List.Count(names) >= 2
                                       and not List.Contains({ColCat, ColDrug, ColPat, ColQty, ColStock, ColIncl}, names{1})
                                    then names{1} else null,
                           col    = if byName <> null then byName else byPos
                       in
                           if col = null then List.Repeat({null}, Table.RowCount(tbl)) else Table.Column(tbl, col),
            ren  = if body = null then null
                   else Table.FromColumns(
                            {pick(body, ColCat), abbr(body), pick(body, ColDrug), pick(body, ColPat),
                             pick(body, ColQty), pick(body, ColStock), pick(body, ColIncl)},
                            {"薬品区分", "略称", "薬品名", "患者数", "数量", "在庫量", "在庫量_発注残込み"}),
            withMeta = if ren = null then null
                   else Table.AddColumn(Table.AddColumn(Table.AddColumn(ren,
                            "日付", (x) => d, type date),
                            "ファイル名", (x) => f, type text),
                            "シート名", (x) => s, type text)
        in
            if withMeta = null then Empty else withMeta),

    結合 = if Table.IsEmpty(整形) then Empty else Table.Combine(整形[表]),

    // （0.4.0 より前は Text.Remove に 2 文字以上の「mL」「cap」を渡していてエラーになり、文字で来た数値がすべて空になっていた）
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

    型変換 = Table.TransformColumns(結合, {
                 {"薬品区分", each if _ = null then null else Nz(_), type nullable text},
                 {"略称",     each if _ = null then null else Text.Trim(Text.From(_)), type nullable text},
                 {"薬品名",   each if _ = null then null else Nz(_), type nullable text},
                 {"患者数",   Num, type nullable number},
                 {"数量",     Num, type nullable number},
                 {"在庫量",   Num, type nullable number},
                 {"在庫量_発注残込み", Num, type nullable number}}),

    // 空の薬品名・合計行・見出しの繰り返しを除き、使用量（または患者数）がある行だけ残す
    有効行 = Table.SelectRows(型変換, each
                 [薬品名] <> null and [薬品名] <> ""
                 and not List.Contains({"合計", "総合計", "小計", ColDrug}, Norm([薬品名]))
                 and (KeepZero or ([数量] <> null and [数量] > 0) or ([患者数] <> null and [患者数] > 0))),

    結果 = Table.SelectColumns(有効行, {"日付", "薬品区分", "略称", "薬品名", "患者数", "数量", "在庫量",
                                         "在庫量_発注残込み", "ファイル名", "シート名"})
in
    結果
