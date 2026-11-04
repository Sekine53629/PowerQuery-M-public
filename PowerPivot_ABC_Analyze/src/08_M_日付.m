// ================================================================================
// ⑧ クエリ名：M_日付（データモデルに読み込む）
//    日付の表。T_日次[日付] と T_月次[月] をここにつなぐ
//    ・範囲：実績（① の開始日～終了日）と出納の期間をすべて含む月の、1日～末日
//    ・処方のあった日：T_日次 にその日のデータがある（薬局が開いていた日の目安）
//    ・卸の営業日・休みの理由：④ fx_卸の休み と同じ判定
// ================================================================================
let
    Prm   = #"設定値",
    Days  = List.Buffer(List.Distinct(#"T_日次"[日付])),
    Mons  = #"T_月次"[月],
    Lo    = Date.StartOfMonth(List.Min(List.RemoveNulls({Date.From(Prm[開始日]), List.Min(Mons)}))),
    Hi    = Date.EndOfMonth(List.Max(List.RemoveNulls({Date.From(Prm[終了日]), List.Max(Mons)}))),
    Cal   = #"fx_卸の休み"(Lo, Hi),
    Add   = Table.AddColumn(Cal, "列", each
        let d = [日付], m = Date.Month(d), q = Date.QuarterOfYear(d)
        in  [年 = Date.Year(d), 月 = m,
             年月 = Date.ToText(d, "yyyy-MM"),
             月初 = Date.StartOfMonth(d),
             期 = {"1-3月", "4-6月", "7-9月", "10-12月"}{q - 1},
             曜日番号 = Date.DayOfWeek(d, Day.Monday) + 1,
             処方のあった日 = List.Contains(Days, d)]),
    Exp   = Table.ExpandRecordColumn(Add, "列", {"年", "月", "年月", "月初", "期", "曜日番号", "処方のあった日"}),
    Result = Table.TransformColumnTypes(
                 Table.SelectColumns(Exp, {"日付", "年", "月", "年月", "月初", "期", "曜日", "曜日番号",
                                           "処方のあった日", "卸の営業日", "休みの理由"}),
                 {{"年", Int64.Type}, {"月", Int64.Type}, {"年月", type text}, {"月初", type date},
                  {"期", type text}, {"曜日番号", Int64.Type}, {"処方のあった日", type logical}})
in
    Result
