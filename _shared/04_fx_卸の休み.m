// ================================================================================
// ④ クエリ名：fx_卸の休み（関数）
//    firstDay～lastDay の各日について、卸の営業日かどうかと休みの理由を返す
//    休み：土曜・日曜／国民の祝日・振替休日・国民の休日（計算で自動判定）／
//          年末年始（設定の「年末年始の休み始め」～「年末年始の休み終わり」）／
//          追加休業日テーブルに書いた日（お盆・卸の都合など）
//    ※ 春分の日・秋分の日は天文の近似式（1980～2099年で有効）
//    ⑧ M_日付 で使う（PowerQuery_ABC では ⑰ 納品カレンダー で使う。両方で同じコード）
// ================================================================================
(firstDay as date, lastDay as date) as table =>
let
    Prm       = #"設定値",
    FirstDay  = firstDay,
    LastDay   = lastDay,
    Years     = {Date.Year(FirstDay) .. Date.Year(LastDay)},

    // 「12/29」のような文字、または日付のセルから {月, 日} を取り出す
    MonthDay = (v as any, dflt as text) as list =>
        let
            x = if v = null or (v is text and Text.Trim(v) = "") then dflt else v,
            p = if x is date or x is datetime then null
                else Text.Split(Text.Replace(Text.Trim(Text.From(x)), "-", "/"), "/")
        in
            if p = null then {Date.Month(Date.From(x)), Date.Day(Date.From(x))}
            else {Number.From(p{List.Count(p) - 2}), Number.From(List.Last(p))},
    YeStart = MonthDay(Record.FieldOrDefault(Prm, "年末年始の休み始め", null), "12/29"),
    YeEnd   = MonthDay(Record.FieldOrDefault(Prm, "年末年始の休み終わり", null), "1/3"),

    // その年の国民の祝日（振替休日・国民の休日を除く）
    Holidays = (y as number) as list =>
        let
            NthMonday = (m as number, n as number) as date =>
                let f = #date(y, m, 1)
                in  Date.AddDays(f, Number.Mod(7 - Date.DayOfWeek(f, Day.Monday), 7) + 7 * (n - 1)),
            k    = y - 1980,
            Shun = Number.RoundDown(20.8431 + 0.242194 * k - Number.RoundDown(k / 4)),
            Shu  = Number.RoundDown(23.2488 + 0.242194 * k - Number.RoundDown(k / 4))
        in
            {
                [日付 = #date(y, 1, 1),   名前 = "元日"],
                [日付 = NthMonday(1, 2),  名前 = "成人の日"],
                [日付 = #date(y, 2, 11),  名前 = "建国記念の日"],
                [日付 = #date(y, 2, 23),  名前 = "天皇誕生日"],
                [日付 = #date(y, 3, Shun), 名前 = "春分の日"],
                [日付 = #date(y, 4, 29),  名前 = "昭和の日"],
                [日付 = #date(y, 5, 3),   名前 = "憲法記念日"],
                [日付 = #date(y, 5, 4),   名前 = "みどりの日"],
                [日付 = #date(y, 5, 5),   名前 = "こどもの日"],
                [日付 = NthMonday(7, 3),  名前 = "海の日"],
                [日付 = #date(y, 8, 11),  名前 = "山の日"],
                [日付 = NthMonday(9, 3),  名前 = "敬老の日"],
                [日付 = #date(y, 9, Shu), 名前 = "秋分の日"],
                [日付 = NthMonday(10, 2), 名前 = "スポーツの日"],
                [日付 = #date(y, 11, 3),  名前 = "文化の日"],
                [日付 = #date(y, 11, 23), 名前 = "勤労感謝の日"]
            },

    Base      = List.Combine(List.Transform(Years, Holidays)),
    BaseDates = List.Buffer(List.Transform(Base, each [日付])),
    // 振替休日：祝日が日曜なら、その後の最初の祝日でない日
    Subst     = List.Transform(
                    List.Select(Base, each Date.DayOfWeek([日付], Day.Monday) = 6),
                    (h) => [日付 = List.First(List.Select(
                                       List.Transform({1 .. 7}, (i) => Date.AddDays(h[日付], i)),
                                       (d) => not List.Contains(BaseDates, d))),
                            名前 = "振替休日"]),
    AllDates  = List.Dates(FirstDay, Duration.Days(LastDay - FirstDay) + 1, #duration(1, 0, 0, 0)),
    // 国民の休日：前日と翌日が祝日の日
    Kokumin   = List.Transform(
                    List.Select(AllDates, (d) => not List.Contains(BaseDates, d)
                                and List.Contains(BaseDates, Date.AddDays(d, -1))
                                and List.Contains(BaseDates, Date.AddDays(d, 1))),
                    (d) => [日付 = d, 名前 = "国民の休日"]),
    HolTbl    = Table.Buffer(Table.Distinct(Table.FromRecords(Base & Subst & Kokumin), {"日付"})),

    // 追加休業日テーブル（任意）：列「日付」「理由」
    ExtraRaw  = try Table.SelectColumns(Excel.CurrentWorkbook(){[Name = "追加休業日"]}[Content],
                        {"日付", "理由"}, MissingField.UseNull)
                otherwise #table({"日付", "理由"}, {}),
    Extra     = Table.Buffer(Table.SelectRows(Table.TransformColumns(ExtraRaw, {
                    {"日付", each try Date.From(_) otherwise null, type nullable date},
                    {"理由", each if _ = null or Text.Trim(Text.From(_)) = "" then "追加休業日" else Text.From(_), type text}}),
                    each [日付] <> null)),

    IsYearEnd = (d as date) as logical =>
        let
            md = Date.Month(d) * 100 + Date.Day(d),
            s  = YeStart{0} * 100 + YeStart{1},
            e  = YeEnd{0} * 100 + YeEnd{1}
        in
            if s <= e then md >= s and md <= e else md >= s or md <= e,

    Rows = List.Transform(AllDates, (d) =>
        let
            dow = Date.DayOfWeek(d, Day.Monday),
            hol = Table.SelectRows(HolTbl, each [日付] = d),
            ext = Table.SelectRows(Extra, each [日付] = d),
            why = if not Table.IsEmpty(ext) then ext{0}[理由]
                  else if not Table.IsEmpty(hol) then hol{0}[名前]
                  else if IsYearEnd(d) then "年末年始"
                  else if dow = 5 then "土曜"
                  else if dow = 6 then "日曜"
                  else null
        in
            [日付 = d, 曜日 = Text.Start(Date.DayOfWeekName(d, "ja-JP"), 1),
             卸の営業日 = why = null, 休みの理由 = why]),

    Result = Table.TransformColumnTypes(Table.FromRecords(Rows), {
                 {"日付", type date}, {"曜日", type text}, {"卸の営業日", type logical}, {"休みの理由", type text}})
in
    Result
