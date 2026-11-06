// ================================================================================
// ① クエリ名：設定値
//    設定テーブル（計算に使う値とフォルダ）を読み込む
//    エリシアSの出力の列名・ファイル名は決まっているので、ここに固定で書く（設定テーブルでは変えない）
//    実績を読む期間（開始日～終了日）は、設定になければ「発注日までの 1 年間」にする（0.3.0）
// ================================================================================
let
    Source = Excel.CurrentWorkbook(){[Name = "設定"]}[Content],
    Typed  = Table.TransformColumnTypes(Source, {{"項目", type text}}),
    Rec    = Record.FromList(Typed[値], Typed[項目]),

    // 空欄なら初期値を返す
    Get = (key as text, dflt as any) as any =>
        let v = Record.FieldOrDefault(Rec, key, null)
        in  if v = null or (v is text and Text.Trim(v) = "") then dflt else v,

    // ---- 固定：薬品使用量分析データの列名（実績・予測とも同じ） ----
    Cols = [
        薬品名     = "薬品名",
        数量       = "予定使用量",
        在庫量     = "在庫数量（発注残を含まない）",
        在庫量込み = "在庫数量",
        薬品区分   = "製剤",
        患者数     = "患者数",
        略称       = "略称",
        日付       = "日付"       // 文字のファイル（.csv／.txt）で、1 つのファイルに何日分も入っているときの日付の列
    ],
    ColsHist = Cols,
    ColsFc   = Cols,
    // ---- 固定：ファイル名 ----
    //   薬品使用量分析データ：「薬品使用量」で始まる
    //   出納（月次）        ：「数量一覧」を含む（例：20261001_20261031_数量一覧.xls。.xlsx／.csv／.txt でもよい）
    Prefix    = "薬品使用量",
    SuiMark   = "数量一覧",

    // 発注日（空欄なら今日）
    OrderRaw   = Get("発注日", null),
    OrderDate  = if OrderRaw = null then Date.From(DateTime.LocalNow()) else Date.From(OrderRaw),
    // 予測期間（予測開始日が空欄なら発注日の翌日から）
    //   予測終了   ：区分別発注一覧（⑪）で使う「来週分」の終わり（予測日数）
    //   予測読込終了：予測ファイルを読む終わり（予測の読込日数。発注シートで連休の先まで見るため）
    FcStartRaw = Get("予測開始日", null),
    FcStart    = if FcStartRaw = null then Date.AddDays(OrderDate, 1) else Date.From(FcStartRaw),
    FcEnd      = Date.AddDays(FcStart, Number.From(Get("予測日数", 7)) - 1),
    FcLoadEnd  = List.Max({FcEnd, Date.AddDays(FcStart, Number.From(Get("予測の読込日数", 14)) - 1)}),
    // 実績を読む期間（設定に開始日・終了日があればその値。なければ次で決める）
    //   終了日 ＝ 予測開始日の前日（予測開始日が空欄なら発注日）
    //   開始日 ＝ 終了日の 1 年前の翌日が属する月の 1 日（出納が月単位なので、月の途中から読まない）
    EndRaw     = Get("終了日", null),
    HistEnd    = if EndRaw = null then Date.AddDays(FcStart, -1) else Date.From(EndRaw),
    StartRaw   = Get("開始日", null),
    HistStart  = if StartRaw = null then Date.StartOfMonth(Date.AddDays(Date.AddYears(HistEnd, -1), 1)) else Date.From(StartRaw),

    // フォルダパスの前後の「"」・空白・末尾の「\」を取り除く
    // （エクスプローラーの「パスのコピー」で付く「"」などに対応）
    CleanPath = (v as any) as nullable text =>
        if v = null then null
        else Text.TrimEnd(Text.Trim(Text.Remove(Text.From(v), {""""})), {"\", "/", " ", "　"}),

    Result = Rec & [列名 = Cols, 列名_実績 = ColsHist, 列名_予測 = ColsFc,
                    先頭 = Prefix, 出納の目印 = SuiMark,
                    発注日 = OrderDate, 予測開始 = FcStart, 予測終了 = FcEnd, 予測読込終了 = FcLoadEnd,
                    開始日 = HistStart, 終了日 = HistEnd,
                    フォルダパス     = CleanPath(Get("フォルダパス", null)),
                    予測フォルダパス = CleanPath(Get("予測フォルダパス", null)),
                    出納フォルダパス = CleanPath(Get("出納フォルダパス", null))]
in
    Result
