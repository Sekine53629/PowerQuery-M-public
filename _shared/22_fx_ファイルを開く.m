// ================================================================================
// ㉒ クエリ名：fx_ファイルを開く（関数・⓪ の次に作る）
//    エリシアSの出力ファイル 1 つを開いて、中の表を返す関数（②⑤⑭⑳㉑㉓ が使う）
//      content  ：ファイルの中身（Folder.Files の [Content]）
//      fileName ：ファイル名（エラーの文に出す）
//      headerLabel：（省略可）見出しの文字（例「品名」）。渡すと、先頭 30 行にその見出しがある表だけを返す
//                   （1 つもなければ全部返す）。.xls はシートが名前の順に並ぶので、「1 枚目」を決めるために使う
//    出力列：シート名／Data（見出しにしていない、そのままの表）／形式
//
//    読める形式（拡張子ではなく、ファイルの中身で見分ける）
//      ・Excel ブック（.xlsx／.xls）… シートごとに 1 行
//      ・文字のファイル（.csv／.txt／.tsv）… 1 行だけ（シート名は空）
//          区切り　：カンマかタブを自動で見分ける（先頭 50 行で多い方）
//          文字コード：UTF-8（BOM あり・なし）／UTF-16／Shift_JIS を自動で見分ける
//          値はすべて文字で返す（空欄は null）。数値・日付への変換は呼び出す側で行う
//      ・拡張子が .xls でも中身が文字のファイルなら、文字のファイルとして読む
//    .xls（互換形式）の癖（Excel で確認、2026-11-07）と対策
//      ・シート一覧は Name と Data の 2 列だけ（Kind 列がない）。シートは名前の順に並ぶ。非表示のシートは出てこない
//      ・値はすべて「表示されている文字」で来る（数値は 1,234、日付は 2026/10/01 のような文字）。
//        数値・日付への変換は呼び出す側で行う。指数表示になった長い数字（JAN コードなど）は元に戻せない
//      ・シート名に付く「'」と末尾の「$」を外す（'10月1日(水)$' → 10月1日(水)）
//      ・名前の途中に「$」があるもの（印刷範囲などの名前）はシートとみなさない
//    開けないファイルは、ファイル名を付けたエラーにする（どのファイルかは ㉓ 読込の確認 にも出る）
// ================================================================================
(content as binary, fileName as text, optional headerLabel as nullable text) as table =>
let
    Nz    = #"fx_正規化",
    Norm  = (t as any) as text =>
        let s = if t = null then "" else Nz(Text.From(t))
        in  Text.Replace(Text.Replace(Text.Remove(s, {" ", "　", "#(lf)", "#(cr)", "#(tab)"}), "(", "（"), ")", "）"),
    Bytes = Binary.Buffer(content),
    Len   = Binary.Length(Bytes),
    Head  = Binary.ToList(Binary.Range(Bytes, 0, List.Min({8, Len}))),
    Has   = (sig as list) as logical =>
                List.Count(Head) >= List.Count(sig) and List.FirstN(Head, List.Count(sig)) = sig,
    IsOle = Has({208, 207, 17, 224}),      // D0 CF 11 E0：Excel 97-2003（.xls）
    IsZip = Has({80, 75}),                 // 50 4B（PK）：.xlsx
    OutCols = {"シート名", "Data", "形式"},

    // ---- Excel ブック ----
    FromExcel = () as table =>
        let
            wb     = Excel.Workbook(Bytes, null, true),
            sheets = if Table.HasColumns(wb, "Kind") then Table.SelectRows(wb, each [Kind] = "Sheet") else wb,
            named  = Table.AddColumn(sheets, "シート名", each
                         let
                             raw = Text.Trim(Text.Remove(Text.From([Name]), {"'"})),
                             cut = Text.TrimEnd(raw, "$")
                         in  if Text.Contains(cut, "$") or Text.Contains(cut, "Print_") or Text.Contains(cut, "_xlnm")
                             then null else cut, type nullable text),
            ok     = Table.SelectRows(named, each [シート名] <> null),
            kind   = if IsOle then "Excel（.xls）" else "Excel（.xlsx）",
            withK  = Table.AddColumn(ok, "形式", each kind, type text)
        in
            Table.SelectColumns(withK, OutCols),

    // ---- 文字のファイル ----
    FromText = () as table =>
        let
            // 文字コード：BOM があればそれに従う。なければ UTF-8 として読んでみて、読めない文字があれば Shift_JIS
            enc   = if Has({239, 187, 191}) then 65001
                    else if Has({255, 254}) then 1200
                    else if Has({254, 255}) then 1201
                    else if (try Text.Contains(Text.FromBinary(Bytes, 65001), "#(FFFD)") otherwise true) then 932
                    else 65001,
            txt0  = Text.FromBinary(Bytes, enc),
            txt   = if Text.StartsWith(txt0, "#(FEFF)") then Text.Range(txt0, 1) else txt0,
            lines = Lines.FromText(txt),
            // 区切り：先頭 50 行（空行を除く）で、タブとカンマの多い方
            head  = List.FirstN(List.Select(lines, each Text.Trim(_) <> ""), 50),
            count = (d as text) as number => List.Sum(List.Transform(head, each List.Count(Text.Split(_, d)) - 1)) ?? 0,
            tabs  = count("#(tab)"),
            delim = if tabs > 0 and tabs >= count(",") then "#(tab)" else ",",
            // 列の数：一番長い行に合わせる（表題の行が短くても、データの列が切れないように）
            n     = List.Max(List.Transform(lines, each List.Count(Text.Split(_, delim))), 1),
            tbl   = Csv.Document(Text.ToBinary(txt, 65001),
                        [Delimiter = delim, Columns = n, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
            clean = Table.TransformColumns(tbl, {}, each if _ = null or _ = "" then null else _),
            kind  = "文字（" & (if delim = "," then "カンマ区切り" else "タブ区切り") & "・"
                    & (if enc = 932 then "Shift_JIS" else if enc = 65001 then "UTF-8" else "UTF-16") & "）"
        in
            #table(OutCols, {{"", clean, kind}}),

    // 中身の先頭が「<」のファイル（HTML・XML を .xls の名前で保存したもの）は読めない
    First   = if Len = 0 then null else List.First(List.Select(Head, each _ <> 239 and _ <> 187 and _ <> 191), null),
    Tried   = try Table.Buffer(if IsOle or IsZip then FromExcel() else FromText()),
    // 見出しのある表だけに絞る（headerLabel を渡したとき）
    HasLabel = (tbl as table) as logical =>
        List.AnyTrue(List.Transform(Table.ToRows(Table.FirstN(tbl, 30)),
            (r) => List.Contains(List.Transform(r, Norm), Norm(headerLabel)))),
    Pick = (sheets as table) as table =>
        if headerLabel = null then sheets
        else
            let hit = Table.SelectRows(sheets, (s) => (try HasLabel(s[Data]) otherwise false))
            in  if Table.IsEmpty(hit) then sheets else hit,
    Result  = if Len = 0 then error Error.Record("読めないファイル", "「" & fileName & "」は空のファイルです")
              else if not (IsOle or IsZip) and First = 60 then error Error.Record("読めないファイル",
                       "「" & fileName & "」は中身が HTML か XML です。Excel で開いて .xlsx か .csv で保存し直してください")
              else if Tried[HasError] then error Error.Record("読めないファイル",
                       "「" & fileName & "」を開けません：" & Tried[Error][Message])
              else Pick(Tried[Value])
in
    Result
