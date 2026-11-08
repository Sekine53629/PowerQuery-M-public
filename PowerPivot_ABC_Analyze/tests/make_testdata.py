"""PowerPivot_ABC のテストデータと期待値を作る（すべて架空のデータ）
使い方（tests フォルダで）：python make_testdata.py
  出力：tests/out/
    実績/薬品使用量分析データ2026MM.xlsx   … 薬品使用量分析データ（日ごとのシート）
    出納/2026MM01_2026MMDD_数量一覧.xlsx   … 出納の月次（その月に動いた薬だけ）
    システムリスト/経過措置薬品/経過措置薬品リスト.xlsx      … エリシアSの 3 リスト（⑳ ⑭ ㉑ が読む）
    システムリスト/期限切れ/YYYYMMDD_期限切れリスト.xlsx
    システムリスト/不動薬/YYYYMM_不動品リスト.xlsx
    手入力テーブル.xlsx                    … 設定ほか（bi/run.py --tables に渡せる。Excel に貼る値も同じ）
    期待値.xlsx                            … T_月次・M_薬品（主な指標）・T_判定理由の正解
    設定の値.txt                           … 分析用ブックの設定テーブルに入れる値
  実物は .xls なので、Excel での確認は必ず .xls でも行う（.xlsx だけでは旧形式の癖が出ない）
    .xls を作る       ：pwsh -File tools/make_xls.ps1（Excel を使う。tests/out/形式別/xls に出る）
                        LibreOffice（soffice）があれば --xls でも作れる（実績・出納を .xls に置き換える）
    .csv／.txt を作る ：python tests/make_formats.py（tests/out/形式別/csv・txt。--check で Python 版の結果を比べる）
  ※ このスクリプトは tests/out を作り直す（形式別 も消える）ので、あとで上の 2 つを実行し直す

  ※ 実データは使わない・置かない。ここで作るのは決まった乱数（SEED）の架空データだけ
"""
import os, sys, math, random, shutil, statistics, subprocess, datetime as dt
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font

SEED = 20261102
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
START, END = dt.date(2026, 4, 1), dt.date(2026, 9, 30)   # 実績・出納の期間（6か月）
BASE_DATE = dt.date(2026, 10, 1)                          # 設定「発注日」（経過措置の基準日）
TH = dict(不動とみなす月数=3, 在庫過多とみなす日数=60, ずれの注意割合=0.3, ずれの大きさの注意割合=0.5,
          区分Aの線=1000, 区分Bの線=300, 区分Dの患者数=2,          # 0.5.0：区分の線
          極小需要の月の払出=100, 傾向の線=0.2)                    # 0.10.0：傾向（少量の線と増減の線）
RECENT_N = 3                                              # 設定「直近の月数」の初期値（直近の払出・在庫日数・S05）
ABC_N = 3                                                 # 設定「ABCの月数」の初期値（区分の窓）
FC_DAYS = 7                                               # 設定「予測日数」の初期値（発注書 F の来週）
# 薬品グループ（手入力）：旧名称 → 新名称の切替の例（0.7.0）。アリピプラゾール錠3mg は予定では「（旧）」付き、出納では新しい名前
#   まとまり＝「アリピプラゾール 3mg」。（旧）が切替元、新しい名前が切替先。手入力テーブルの名前は正規化後の表記（半角のかっこ）
GROUP_ROWS = [  # 薬品名, 成分規格, 需要のまとまり, 先発後発, 供給状況, 切替先
    ("アリピプラゾール錠3mg(旧)", "アリピプラゾール錠3mg", "アリピプラゾール 3mg", "後発", None, "アリピプラゾール錠3mg"),
    ("アリピプラゾール錠3mg", "アリピプラゾール錠3mg", "アリピプラゾール 3mg", "後発", None, None),
    ("セルトラリン錠25mg「アメル」", "セルトラリン錠25mg", "セルトラリン 25mg", "後発", "通常", None),
    ("セルトラリン錠25mg「トーワ」", "セルトラリン錠25mg", "セルトラリン 25mg", "後発", None, None),
]
SUPPLY_OF = {r[0]: r[4] for r in GROUP_ROWS}             # 薬品名 → 供給状況（手入力）
GROUP_OF = {r[0]: r[2] for r in GROUP_ROWS}             # 薬品名（正規化後） → 需要のまとまり
SWITCH_TO = {r[0]: r[5] for r in GROUP_ROWS if r[5]}    # 切替元 → 切替先
# 卸の休み（④ と同じ：土日・祝日・年末年始）。2026 年 4～10 月の祝日を固定で列挙（振替休日・国民の休日を含む）
HOLIDAYS_2026 = {dt.date(2026, 4, 29), dt.date(2026, 5, 3), dt.date(2026, 5, 4), dt.date(2026, 5, 5), dt.date(2026, 5, 6),
                 dt.date(2026, 7, 20), dt.date(2026, 8, 11), dt.date(2026, 9, 21), dt.date(2026, 9, 22), dt.date(2026, 9, 23),
                 dt.date(2026, 10, 12)}
WD = "月火水木金土日"

# 薬：名前・略称・型・1回あたりの量・処方の確率・実際/予定の倍率・臨時の確率
#   kind: regular（定期）, sparse（まばら）, dead（最初の月だけ動く）, ending（途中で処方終了）,
#         extra（予定に出ない＝出納だけ）, planonly（予定だけ＝出納と名前がずれる）,
#         rising（処方の確率が 0.5p → 1.5p に増えていく。傾向「増加」の確認用）
DRUGS = [
    ("デエビゴ錠2.5mg",            "ﾃﾞｴﾋﾞｺﾞ2.5", "regular",  28, 0.95, 1.00, 0.10),
    ("セルトラリン錠25mg「アメル」", "ｾﾙﾄﾗﾘﾝ25",   "regular",  14, 0.80, 1.05, 0.15),
    ("セルトラリン錠25mg「トーワ」", "ｾﾙﾄﾗﾄﾜ25",   "rising",   14, 0.35, 1.00, 0.10),
    ("エスシタロプラム錠10mg",       "ｴｽｼﾀﾛ10",    "regular",  28, 0.60, 1.00, 0.05),
    ("タンドスピロン錠10mg「サワイ」", "ﾀﾝﾄﾞｽﾋﾟﾛﾝ10", "regular", 42, 0.30, 1.40, 0.20),
    ("ビペリデン錠1mg",              "ﾋﾞﾍﾟﾘﾃﾞﾝ1",  "sparse",   28, 0.08, 1.00, 0.05),
    ("リチウム炭酸塩錠200mg",        "ﾘﾁｳﾑ200",    "sparse",   84, 0.05, 1.00, 0.00),
    ("ハロペリドール錠0.75mg「アメル」", "ﾊﾛﾍﾟﾘ0.75", "dead",    28, 0.10, 1.00, 0.00),
    ("クロキサゾラム錠1㎎",          "ｸﾛｷｻｿﾞﾗﾑ1",  "dead",     14, 0.10, 1.00, 0.00),   # ㎎（1 文字の単位）の読み取りの確認用（0.10.1）
    ("ミルタザピン錠15mg",           "ﾐﾙﾀｻﾞﾋﾟﾝ15", "ending",   28, 0.50, 1.00, 0.00),
    ("ロラゼパム錠0.5mg",            "ﾛﾗｾﾞﾊﾟﾑ0.5", "extra",     7, 0.00, 1.00, 0.10),
    ("アリピプラゾール錠3mg",        "ｱﾘﾋﾟﾌﾟﾗ3",   "planonly", 28, 0.40, 1.00, 0.00),
    ("スルピリド錠５０ｍｇ",          "ｽﾙﾋﾟﾘﾄﾞ50",  "regular",  28, 0.40, 0.95, 0.05),
]
TRANSITIONAL = {"ハロペリドール錠0.75mg「アメル」": dt.date(2027, 3, 31)}   # 手入力の経過措置リスト
# エリシアSの経過措置薬品リストにだけある薬（期限切れ → E02）。⑨ はリストを手入力より優先する
LIST_TRANS = {"クロキサゾラム錠1㎎": dt.date(2026, 9, 30)}
ALL_TRANS = {**TRANSITIONAL, **LIST_TRANS}
# 不動品リストにだけ載る薬（出納にも予定にも一度も出ない → ⑨ で拾われ、S06・Q01 になる）
FUDO_ONLY = ("ペロスピロン錠4mg", "ﾍﾟﾛｽﾋﾟﾛﾝ4", 50)              # 名前・略称・在庫
MARKED = {"ハロペリドール錠0.75mg「アメル」": "劇", "リチウム炭酸塩錠200mg": "劇"}   # リストの品名の先頭に付く記号（ない薬は空白 3 つ）
LIST_DIR = os.path.join(OUT, "システムリスト")


def open_day(d):
    return d.weekday() < 6   # 月～土を営業日とする（祝日は無視した簡略版）


def month_iter():
    m = dt.date(START.year, START.month, 1)
    while m <= END:
        yield m
        m = dt.date(m.year + (m.month == 12), m.month % 12 + 1, 1)


def month_end(m):
    n = dt.date(m.year + (m.month == 12), m.month % 12 + 1, 1)
    return n - dt.timedelta(days=1)


def nz(s):
    """fx_正規化 と同じ考え方の簡易版（全角英数→半角）。名前の照合に使う"""
    return "".join(chr(ord(c) - 0xFEE0) if 0xFF01 <= ord(c) <= 0xFF5E else c for c in s).strip()


def simulate():
    rnd = random.Random(SEED)
    days = [START + dt.timedelta(i) for i in range((END - START).days + 1)]
    plan = defaultdict(dict)       # plan[date][name] = (数量, 患者数)
    actual = defaultdict(float)    # actual[(date, name)] = 払出
    for name, ab, kind, qty, p, mult, extra in DRUGS:
        for d in days:
            if not open_day(d):
                continue
            q = pts = 0
            if kind == "dead" and d >= dt.date(2026, 5, 1):
                pass
            elif kind == "ending" and d >= dt.date(2026, 7, 1):
                # 予定には残るが、実際は出ない（処方終了・残薬調整）
                if rnd.random() < p:
                    pts = rnd.randint(1, 2); plan[d][name] = (qty * pts, pts)
                continue
            elif kind == "rising":
                if rnd.random() < p * (0.5 + (d - START).days / max(1, (END - START).days)):
                    pts = rnd.randint(1, 3); q = qty * pts
            elif rnd.random() < p:
                pts = rnd.randint(1, 3) if kind in ("regular", "rising") else 1
                q = qty * pts
            if q > 0 and kind != "extra":
                plan[d][name] = (q, pts)
            a = q * mult if q else 0
            if rnd.random() < extra:
                a += qty * rnd.randint(1, 2)   # 臨時・新患・他科
            if a:
                actual[(d, name)] += round(a)
    return days, plan, actual


def ledger(days, actual):
    """在庫の動き：仕入・払出・その他の増減を日ごとに作り、月ごとにまとめる"""
    rnd = random.Random(SEED + 1)
    stock = {name: (300 if kind != "dead" else 400) for name, _, kind, *_ in DRUGS}
    pack = 100
    mon = defaultdict(lambda: defaultdict(float))   # mon[(month, name)][項目]
    for name, _, kind, *_ in DRUGS:
        for m in month_iter():
            mon[(m, name)]["繰越数量"] = None
    for d in days:
        m = dt.date(d.year, d.month, 1)
        for name, _, kind, *_ in DRUGS:
            r = mon[(m, name)]
            if r["繰越数量"] is None:
                r["繰越数量"] = stock[name]
            out = actual.get((d, name), 0)
            out = min(out, stock[name] + 1000)  # 欠品はさせない（在庫はマイナスにしない）
            if out:
                r["払出数量"] += out; stock[name] -= out
            if stock[name] < 150 and kind not in ("dead",) and open_day(d):
                buy = pack * math.ceil((300 - stock[name]) / pack)
                r["仕入数量"] += buy; stock[name] += buy
            if kind == "dead" and d == dt.date(2026, 4, 15):
                r["その他の増減"] -= 100; stock[name] -= 100     # 返品
            if name.startswith("デエビゴ") and d == dt.date(2026, 6, 20):
                r["その他の増減"] -= 28; stock[name] -= 28       # 期限切れ廃棄
            r["在庫数量"] = stock[name]
    return mon


def write_xlsx(path, sheets):
    wb = Workbook(); wb.remove(wb.active)
    for title, rows, merged_second in sheets:
        ws = wb.create_sheet(title)
        for r in rows:
            ws.append(r)
        # 実物に近づける：1000 以上の数値は桁区切りで表示（.xls では「1,234」という文字で読まれるので、その確認になる）
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, (int, float)) and not isinstance(c.value, bool) and abs(c.value) >= 1000:
                    c.number_format = "#,##0"
        if merged_second:   # 実物：1行目が見出し、2行目は空欄で1行目と縦に結合
            for c in range(1, len(rows[0]) + 1):
                ws.merge_cells(start_row=1, start_column=c, end_row=2, end_column=c)
        for c in ws[1]:
            c.font = Font(bold=True)
    wb.save(path)


HDR_KEIKA = ["記号", "製剤コード", "製剤", "コード", "品名", "メーカーコード", "メーカー", "経過措置", "在庫数量", "単位"]
HDR_KIGEN = ["記号", "製剤コード", "製剤", "コード", "品名", "メーカーコード", "メーカー", "棚番", "伝票日付", "No", "伝票区分コード", "伝票区分",
             "仕入先コード", "仕入先", "包装", "仕入数量", "単価", "期限", "期限切れ数量", "単位", "期限切れ金額", "ロット番号", "伝票備考",
             "連携ｺｰﾄﾞ", "連番", "品目備考"]
HDR_FUDO = ["記号", "製剤コード", "製剤", "コード", "品名", "メーカーコード", "メーカー", "棚番", "仕入伝票日付", "仕入区分コード", "仕入区分",
            "仕入先コード", "仕入先", "包装", "購入価", "仕入数量", "仕入金額", "期限日", "最終払出日", "払出区分", "払出区分コード", "払出先コード",
            "払出数量", "払出先", "単価", "数量", "単位", "金額", "薬価", "薬価金額", "連携コード", "連番", "品目備考", "ロット番号", "伝票備考",
            "基準日", "最終使用日", "不動日数", "YJコード", "出力日時", "JANコード", "レセプト電算コード"]


def fmt_num(x):
    return int(x) if float(x).is_integer() else x


def marked(n):
    """リストの品名：先頭に「劇 」などの記号、ない薬は空白 3 つ（実物と同じ。クエリが外して照合する）"""
    m = MARKED.get(n)
    return (m + " " if m else "   ") + n


def write_system_lists(mon, actual):
    """エリシアSの 3 リストを tests/out/システムリスト/<種類>/ に書く。ロットは在庫から決まった乱数で分ける
    戻り値：{薬品名: 不動日数 or None}（不動品リストに載せた薬）"""
    rnd = random.Random(SEED + 9)
    dirs = {k: os.path.join(LIST_DIR, k) for k in ("経過措置薬品", "期限切れ", "不動薬")}
    for d in dirs.values():
        os.makedirs(d, exist_ok=True)
    last_month = list(month_iter())[-1]
    stock = {n: mon[(last_month, n)]["在庫数量"] for n, *_ in DRUGS}
    stock[FUDO_ONLY[0]] = FUDO_ONLY[2]
    abbr = {n: ab for n, ab, *_ in DRUGS}; abbr[FUDO_ONLY[0]] = FUDO_ONLY[1]
    shelf = {n: f"{i + 1:03d}" for i, (n, *_) in enumerate(DRUGS)}; shelf[FUDO_ONLY[0]] = "099"
    last_out = {}
    for (d, n), q in actual.items():
        if q > 0 and (n not in last_out or d > last_out[n]):
            last_out[n] = d
    price = {n: round(rnd.uniform(5, 120), 1) for n in stock}
    yj = {n: f"{rnd.randint(1124, 1179)}0{rnd.randint(10, 99)}F{rnd.randint(1, 9)}{rnd.randint(100, 999)}" for n in stock}
    jan = {n: "49" + "".join(str(rnd.randint(0, 9)) for _ in range(11)) for n in stock}
    ws = [("ｱﾙﾌ", "アルフレッサ"), ("ｽｽﾞｹﾝ", "スズケン")]
    # ロット：在庫を 1～2 本に分ける。期限は基準日から 60～900 日
    lots = {}
    for n, st in stock.items():
        if not st or st <= 0:
            continue
        st = int(st)
        qs = [st] if st < 150 or rnd.random() < 0.4 else (lambda c: [c, st - c])(rnd.randint(1, st - 1))
        lots[n] = [dict(ロット=f"{rnd.choice('ABCDEFG')}{rnd.randint(1000, 9999)}", 数量=q,
                        期限=BASE_DATE + dt.timedelta(days=rnd.randint(60, 900)),
                        仕入日=BASE_DATE - dt.timedelta(days=rnd.randint(10, 400)), 仕入先=rnd.choice(ws)) for q in qs]

    def mk(n):
        return [1, "内服", abbr[n][:6], marked(n), "00000", None]

    # ---- 経過措置薬品リスト：手入力と同じ薬 ＋ 期限切れの薬 ----
    rows = [HDR_KEIKA]
    for n, lim in ALL_TRANS.items():
        st = stock.get(n) or 0
        rows.append(["×" if st <= 0 else None] + mk(n)[:5] + [None, lim.strftime("%Y/%m/%d"), fmt_num(st), "錠"])
    write_xlsx(os.path.join(dirs["経過措置薬品"], "経過措置薬品リスト.xlsx"), [("Sheet1", rows, False)])

    # ---- 期限切れリスト：基準日から 180 日以内に期限が来るロット ----
    rows = [HDR_KIGEN]; seq = 0
    for n, ls in lots.items():
        for l in ls:
            if (l["期限"] - BASE_DATE).days <= 180:
                seq += 1
                rows.append([None, 1, "内服", abbr[n][:6], marked(n), "00000", None, shelf[n], l["仕入日"].strftime("%Y/%m/%d"), 1, 11, "仕入",
                             l["仕入先"][0], l["仕入先"][1], None, fmt_num(l["数量"]), price[n], l["期限"].strftime("%Y/%m/%d"), fmt_num(l["数量"]),
                             "錠", round(l["数量"] * price[n]), l["ロット"], f"{l['仕入先'][1]}（{price[n]}/{l['ロット']}）",
                             f"5{l['仕入日'].year - 2018:02d}{l['仕入日'].month:02d}", seq, None])
    write_xlsx(os.path.join(dirs["期限切れ"], f"{END:%Y%m%d}_期限切れリスト.xlsx"), [("Sheet1", rows, False)])

    # ---- 不動品リスト：基準日の 90 日以上前から払出がない薬（一度も出ていない薬も）のロット ----
    rows = [HDR_FUDO]; seq = 0; fudo = {}
    for n, ls in lots.items():
        lo = last_out.get(n)
        idle = (BASE_DATE - lo).days if lo else None
        if lo is not None and idle < 90:
            continue
        fudo[n] = idle
        for l in ls:
            seq += 1
            rows.append([None, 1, "内服", abbr[n][:6], marked(n), "00000", None, shelf[n], l["仕入日"].strftime("%Y/%m/%d"), 11, "仕入",
                         l["仕入先"][0], l["仕入先"][1], "100錠", price[n], fmt_num(l["数量"]), round(l["数量"] * price[n]),
                         l["期限"].strftime("%Y/%m/%d"), lo.strftime("%Y/%m/%d") if lo else None, "払出" if lo else None, 21 if lo else None,
                         1 if lo else None, fmt_num(l["数量"]) if lo else None, "処方" if lo else None, price[n], fmt_num(l["数量"]), "錠",
                         round(l["数量"] * price[n]), price[n], round(l["数量"] * price[n]), f"5{l['仕入日'].year - 2018:02d}{l['仕入日'].month:02d}",
                         seq, None, l["ロット"], f"{l['仕入先'][1]}（{price[n]}/{l['ロット']}）", BASE_DATE.strftime("%Y/%m/%d"),
                         lo.strftime("%Y/%m/%d") if lo else None, idle, yj[n], BASE_DATE.strftime("%Y/%m/%d") + " 09:00:00", jan[n], 6 * 10**8 + seq])
    write_xlsx(os.path.join(dirs["不動薬"], f"{END:%Y%m}_不動品リスト.xlsx"), [("Sheet1", rows, False)])
    return fudo


def fc_stock(st):
    """予測データの在庫（発注残を含まない）：9 月末の在庫 − 250（0 未満は 0）"""
    return max(0, (st or 0) - 250)


def write_forecast(mon, info, plan_name, H, HDR):
    """予測：発注日の翌日から 14 日分（10/2～10/15）の薬品使用量予測データ（当月のファイル 202610）。
       別の乱数で作るので、1 年分の期待値は変わらない。10/1 のシートは作らない（③ が実績として読んでしまうため）。
       在庫数量（発注残を含まない）＝ 9 月末の在庫 − 250（0 未満は 0。不足する薬を作るため）、在庫数量（発注残込み）＝ その +50。
       戻り値：fc[日付][薬品名] = (数量, 患者数)"""
    rnd = random.Random(SEED + 2)
    last_m = list(month_iter())[-1]
    fc = {}
    for i in range(1, 15):
        d = BASE_DATE + dt.timedelta(days=i)
        if not open_day(d):
            continue
        fc[d] = {}
        for name, ab, kind, qty, p, mult, extra in DRUGS:
            if kind in ("extra", "dead"):
                continue
            if rnd.random() < p:
                pts = rnd.randint(1, 3) if kind in ("regular", "rising") else 1
                fc[d][name] = (qty * pts, pts)
    sheets = []
    for d in sorted(fc):
        rows = [HDR, [None] * len(HDR)]
        for n in sorted(fc[d]):
            q, p = fc[d][n]
            st = fc_stock(mon[(last_m, n)]["在庫数量"])
            rows.append(["内用", info[n][0], plan_name(n), p, q, st, st + 50, "錠", None, None, None])
        sheets.append((f"{d.month}月{d.day}日({WD[d.weekday()]})", rows, True))
    write_xlsx(os.path.join(H, "薬品使用量分析データ202610.xlsx"), sheets)
    return fc


def write_hand_tables(H, S):
    """手入力テーブル.xlsx（設定・経過措置リスト・発注除外・在庫情報・追加休業日・薬品グループ）。bi/run.py --tables に渡せる"""
    wb = Workbook(); wb.remove(wb.active)
    ws = wb.create_sheet("設定"); ws.append(["項目", "値"])
    for k, v in [("フォルダパス", H), ("出納フォルダパス", S), ("発注日", BASE_DATE.strftime("%Y/%m/%d")),
                 ("期限切迫フォルダパス", os.path.join(LIST_DIR, "期限切れ")), ("経過措置フォルダパス", os.path.join(LIST_DIR, "経過措置薬品")),
                 ("不動品フォルダパス", os.path.join(LIST_DIR, "不動薬"))]:
        ws.append([k, v])
    ws = wb.create_sheet("経過措置リスト"); ws.append(["薬品名", "経過措置期限"])
    for n, d in TRANSITIONAL.items():
        ws.append([n, d.strftime("%Y/%m/%d")])
    wb.create_sheet("発注除外").append(["薬品名", "理由", "登録日"])
    wb.create_sheet("在庫情報").append(["薬品名", "列数", "使用期限"])
    wb.create_sheet("追加休業日").append(["日付", "理由"])
    ws = wb.create_sheet("薬品グループ"); ws.append(["薬品名", "成分規格", "需要のまとまり", "先発後発", "供給状況", "切替先"])
    for row in GROUP_ROWS:
        ws.append(list(row))
    for w in wb:
        for c in w[1]:
            c.font = Font(bold=True)
    wb.save(os.path.join(OUT, "手入力テーブル.xlsx"))


def main():
    if os.path.isdir(OUT):
        shutil.rmtree(OUT)
    H, S = os.path.join(OUT, "実績"), os.path.join(OUT, "出納")
    os.makedirs(H); os.makedirs(S)
    days, plan, actual = simulate()
    mon = ledger(days, actual)
    info = {n: (ab, kind) for n, ab, kind, *_ in DRUGS}

    # ---- 実績：薬品使用量分析データ（予定）。planonly は名前をずらす ----
    HDR = ["製剤", "略称", "薬品名", "患者数", "予定使用量", "在庫数量（発注残を含まない）", "在庫数量",
           "単位", "過不足", "発注点", "備考"]
    def plan_name(n):
        return n + "（旧）" if info[n][1] == "planonly" else n
    stock_by_day = {}
    for m in month_iter():
        md = [d for d in days if d.month == m.month and open_day(d)]
        sheets = []
        tot = defaultdict(lambda: [0, 0])
        for d in md:
            for n, (q, p) in plan[d].items():
                tot[n][0] += p; tot[n][1] += q
        rows = [HDR, [None] * len(HDR)]
        for n in sorted(tot):
            rows.append(["内用", info[n][0], plan_name(n), tot[n][0], tot[n][1], None, None, "錠", None, None, None])
        sheets.append(("期間すべて", rows, True))
        for d in md:
            rows = [HDR, [None] * len(HDR)]
            for n in sorted(plan[d]):
                q, p = plan[d][n]
                st = mon[(m, n)]["在庫数量"]
                rows.append(["内用", info[n][0], plan_name(n), p, q, st, st, "錠", None, None, None])
            sheets.append((f"{d.month}月{d.day}日({WD[d.weekday()]})", rows, True))
        write_xlsx(os.path.join(H, f"薬品使用量分析データ{m:%Y%m}.xlsx"), sheets)

    # ---- 出納（実際）：その月に動いた薬だけ。品名の一部は先頭に空白、スルピリドは全角のまま ----
    SHDR = ["製剤コード", "製剤", "コード", "品名", "繰越数量", "仕入数量", "払出数量", "在庫数量", "単位", "棚番"]
    for m in month_iter():
        # 見出しより上の行：出力日と期間（⑤ はここから期間を読む。なければファイル名から）
        rows = [["数量一覧", None, f"出力日：{month_end(m) + dt.timedelta(days=1):%Y/%m/%d}"],
                [None, None, f"期間：{m:%Y/%m/%d}～{month_end(m):%Y/%m/%d}"],
                [],
                SHDR]
        for i, (n, ab, kind, *_) in enumerate(DRUGS):
            if kind == "planonly":
                pass
            r = mon[(m, n)]
            moved = r["仕入数量"] or r["払出数量"] or r["その他の増減"]
            if not moved:
                continue
            nm = ("   " if i % 2 else "") + n
            rows.append([1, "内用", ab[:6], nm, r["繰越数量"], r["仕入数量"], r["払出数量"], r["在庫数量"], "錠", f"{i + 1:03d}"])
        write_xlsx(os.path.join(S, f"{m:%Y%m%d}_{month_end(m):%Y%m%d}_数量一覧.xlsx"), [("Sheet1", rows, False)])

    fudo = write_system_lists(mon, actual)
    fc = write_forecast(mon, info, plan_name, H, HDR)
    write_hand_tables(H, S)
    expected(days, plan, mon, info, plan_name, fudo, fc)
    with open(os.path.join(OUT, "設定の値.txt"), "w", encoding="utf-8") as f:
        f.write("分析用ブックの設定テーブルに入れる値（ほかは空欄＝初期値）\n"
                f"フォルダパス\t{H}\n出納フォルダパス\t{S}\n"
                f"発注日\t{BASE_DATE:%Y/%m/%d}\n（開始日・終了日は不要：発注日までの 1 年間を読む。予測は 10/2～10/8 を同じ実績フォルダの 202610 から読む）\n"
                f"期限切迫フォルダパス\t{os.path.join(LIST_DIR, '期限切れ')}\n"
                f"経過措置フォルダパス\t{os.path.join(LIST_DIR, '経過措置薬品')}\n"
                f"不動品フォルダパス\t{os.path.join(LIST_DIR, '不動薬')}\n\n経過措置リスト：\n")
        for n, d in TRANSITIONAL.items():
            f.write(f"{n}\t{d:%Y/%m/%d}\n")
    if "--xls" in sys.argv:
        to_xls()
    print("できた：", OUT)


def expected(days, plan, mon, info, plan_name, fudo, fc):
    """M の定義どおりに期待値を計算する（T_月次・M_薬品の主な指標・T_判定理由・T_ABC月次・R_発注書）
    fudo：不動品リストに載せた薬 {薬品名: 不動日数}（S06。出納にも予定にもない薬はここから M_薬品 に入る）
    fc  ：予測 fc[日付][薬品名] = (数量, 患者数)（発注書 F の来週）"""
    # T_日次 相当
    tday = defaultdict(dict)
    for d, dd in plan.items():
        for n, (q, p) in dd.items():
            tday[nz(plan_name(n))][d] = (q, p)
    day_n = len({d for n in tday for d in tday[n]})
    months = list(month_iter())
    cal_days = sum(month_end(m).day for m in months)
    latest = months[-1]
    sui_names = {nz(n) for n, _, kind, *_ in DRUGS}
    names = sorted(set(tday) | {nz(n) for (m, n), r in mon.items()
                                if r["仕入数量"] or r["払出数量"] or r["その他の増減"]} | {nz(n) for n in fudo})
    raw = {nz(n): n for n, *_ in DRUGS}

    # ---- T_ABC月次（⑫）。tests/out は薬品グループが空なので、まとまり＝薬品名 ----
    def out_of(name, m):
        src = raw.get(name)
        r = mon.get((m, src)) if src else None
        moved = bool(r and (r["仕入数量"] or r["払出数量"] or r["その他の増減"]))
        return (r["払出数量"] if moved else 0), moved
    pat_m = defaultdict(float)                       # (月, 薬品名) → 予定の延べ患者数
    plan_m = defaultdict(float)                      # (月, 薬品名) → 予定数量
    for nme, dd in tday.items():
        for d, (q, p) in dd.items():
            pat_m[(dt.date(d.year, d.month, 1), nme)] += p
            plan_m[(dt.date(d.year, d.month, 1), nme)] += q
    exists = {}                                      # (月, 薬品名) → T_月次 に行があるか（動いたか予定があった）
    for nme in names:
        for m in months:
            o, moved = out_of(nme, m)
            exists[(m, nme)] = moved or plan_m.get((m, nme), 0) > 0
    key_of = lambda n2: GROUP_OF.get(n2, n2)             # まとまり（薬品グループにあればそれ、なければ薬品名）
    members = defaultdict(list)
    for n2 in names:
        members[key_of(n2)].append(n2)
    abc_rows, prev_cls = [], {}
    for i, m in enumerate(months):
        win = months[max(0, i + 1 - ABC_N):i + 1]
        seen = sorted({key_of(n2) for (mm, n2), ok in exists.items() if ok and mm <= m})
        vals = {k: sum(out_of(n2, w)[0] for n2 in members[k] for w in win) for k in seen}
        total = sum(vals.values()); cum = 0.0; cur = {}
        for k, v in sorted(vals.items(), key=lambda kv: (-kv[1], kv[0])):
            share = v / total if total else 0.0; cum += share
            avg = v / len(win)
            q = "A" if avg >= TH["区分Aの線"] else "B" if avg >= TH["区分Bの線"] else "C"
            pmin = min(sum(pat_m.get((w, n2), 0.0) for n2 in members[k]) for w in win)
            cls = "D" if pmin < TH["区分Dの患者数"] else q
            abc_rows.append(dict(月=m, 需要のまとまり=k, 払出=v, 構成比=share, 累積=(cum if total else 0.0), 月平均払出=avg,
                                 量の区分=q, 患者の最少月=pmin, 区分=cls, 前月の区分=prev_cls.get(k)))
            cur[k] = cls
        prev_cls = cur
    abc_latest = {r["需要のまとまり"]: r for r in abc_rows if r["月"] == months[-1]}
    sw_from = defaultdict(list)                          # 切替先 → [切替元]
    for a, b in SWITCH_TO.items():
        sw_from[b].append(a)

    # ---- 発注書 F の材料 ----
    #   来週：発注日の翌日から FC_DAYS 日。在庫は一番新しい日の 在庫数量（発注残込み）
    fc_by = defaultdict(dict)                        # 薬品名（正規化・表示名） → {日付: (数量, 患者数, 在庫込み)}
    for d, dd in fc.items():
        if BASE_DATE + dt.timedelta(days=1) <= d <= BASE_DATE + dt.timedelta(days=FC_DAYS):
            for n2, (q, p) in dd.items():
                fc_by[nz(plan_name(n2))][d] = (q, p, fc_stock(mon[(months[-1], n2)]["在庫数量"]) + 50)
    #   納品と納品の間（翌営業日納品）：日付 → 次の営業日
    def is_biz(d):
        return d.weekday() < 5 and d not in HOLIDAYS_2026
    def next_biz(d):
        x = d + dt.timedelta(days=1)
        while not is_biz(x):
            x += dt.timedelta(days=1)
        return x

    t_mon, m_drug, reasons, orders = [], [], [], []
    fmat = {}                                        # 薬品名 → F の材料
    for name in names:
        rows = []
        for m in months:
            src = raw.get(name)
            r = mon.get((m, src)) if src else None
            moved = bool(r and (r["仕入数量"] or r["払出数量"] or r["その他の増減"]))
            out = r["払出数量"] if moved else 0
            pl = sum(q for d, (q, p) in tday.get(name, {}).items() if d.year == m.year and d.month == m.month)
            gap = round(out - pl, 3)
            rows.append(dict(月=m, 薬品名=name, 出納に載った=moved, 払出数量=out, 予定数量=pl, 予定とのずれ=gap,
                             ずれの割合=(gap / pl if pl > 0 else None), 在庫数量=r["在庫数量"] if moved else None,
                             その他の増減=(r["その他の増減"] if moved else 0)))
        has_rows = any(x["出納に載った"] or x["予定数量"] for x in rows)     # T_月次 に行があるか（ないと出納の指標はすべて空欄）
        if not has_rows and name not in fudo:
            continue
        # T_月次 の行は「その月に出納に載った」か「その月に予定があった」薬だけ（⑦ と bi/model.py と同じ。動きも予定もない月は行なし）
        t_mon += [x for x in rows if x["出納に載った"] or x["予定数量"]]
        q = [v[0] for v in tday.get(name, {}).values() if v[0] > 0]
        n = len(q)
        adi = day_n / n if n else None
        cv2 = (statistics.stdev(q) / statistics.mean(q)) ** 2 if n >= 2 else None
        typ = None if adi is None else ("安定" if adi < 1.32 else "間欠") if cv2 is None else \
            (("安定" if cv2 < 0.49 else "ばらつき大") if adi < 1.32 else ("間欠" if cv2 < 0.49 else "まばら・大口"))
        moved = [x for x in rows if x["出納に載った"]]
        out_sum = sum(x["払出数量"] for x in rows)
        per_day = out_sum / cal_days
        # 直近 N ヶ月（設定「直近の月数」の初期値 3）：出納のある月の新しい方から
        recent = months[-RECENT_N:]
        rec_days = sum(month_end(m).day for m in recent)
        rec_out = sum(x["払出数量"] for x in rows if x["月"] in recent)
        prev_m = months[-2 * RECENT_N:-RECENT_N] if len(months) >= 2 * RECENT_N else []
        prev_out = sum(x["払出数量"] for x in rows if x["月"] in prev_m)
        pm = defaultdict(float)                      # 月 → 予定の延べ患者数
        for dd, (q0, p0) in tday.get(name, {}).items():
            pm[dt.date(dd.year, dd.month, 1)] += p0
        prev_pm = months[-2 * ABC_N:-ABC_N] if len(months) >= 2 * ABC_N else []
        pats3 = (sum(pm[m] for m in months[-ABC_N:]) / ABC_N, (sum(pm[m] for m in prev_pm) / len(prev_pm)) if prev_pm else None,
                 sum(pm[m] for m in months) / len(months))
        rec_pd = rec_out / rec_days if rec_days else None
        last = max(moved, key=lambda x: x["月"]) if moved else None
        still = (latest.year - last["月"].year) * 12 + latest.month - last["月"].month if last else None
        stock = last["在庫数量"] if last else None
        plan_sum = sum(x["予定数量"] for x in rows)
        gap_sum = sum(x["予定とのずれ"] for x in rows)
        abs_sum = sum(abs(x["予定とのずれ"]) for x in rows)
        extra_m = sum(1 for x in rows if x["予定数量"] == 0 and x["払出数量"] > 0)
        less = sum(-x["その他の増減"] for x in rows if x["その他の増減"] < 0)
        src = "予定・出納" if n and moved else "予定だけ" if n else "出納だけ"
        trans = ALL_TRANS.get(raw.get(name, name))
        ab = abc_latest.get(key_of(name))
        sw_to = SWITCH_TO.get(name)
        cls_d = "D" if sw_to else (ab["区分"] if ab else None)     # 切替元は D
        d = dict(薬品名=name, 出どころ=src, 区分=cls_d, 量の区分=(ab["量の区分"] if ab else None),
                 患者の最少月=(ab["患者の最少月"] if ab else None),
                 予定日数=(n if n else None), 平均間隔=adi, CV2=cv2, 需要の型=typ,
                 動いた月数=len(moved) if has_rows else None,
                 払出合計=out_sum if has_rows else None, 一日平均払出=per_day if has_rows else None,
                 最後に動いた月=last["月"] if last else None, 最終在庫=stock,
                 動きのない月数=still, 在庫日数=(stock / per_day if stock is not None and per_day > 0 else None),
                 その他の減少=less if has_rows else None, ずれの割合=(gap_sum / plan_sum if plan_sum > 0 else None),
                 ずれの大きさ=(abs_sum / plan_sum if plan_sum > 0 else None), 予定外の月数=extra_m if has_rows else None,
                 経過措置期限=trans, 直近の払出合計=rec_out if has_rows else None,
                 直近の在庫日数=(stock / rec_pd if stock is not None and rec_pd else None))
        m_drug.append(d)
        codes = []
        if (stock or 0) > 0 and still is not None and still >= TH["不動とみなす月数"]: codes.append("S01")
        if d["在庫日数"] is not None and d["在庫日数"] >= TH["在庫過多とみなす日数"]: codes.append("S02")
        if d["直近の在庫日数"] is not None and d["直近の在庫日数"] >= TH["在庫過多とみなす日数"]: codes.append("S05")
        if less > 0: codes.append("S04")
        if raw.get(name, name) in fudo: codes.append("S06")
        codes += {"安定": ["D01"], "ばらつき大": ["D02"], "間欠": ["D03"], "まばら・大口": ["D04"]}.get(typ, [])
        if extra_m >= 1: codes.append("D05")
        if cls_d == "D" and src != "出納だけ" and not sw_to: codes.append("D06")
        if sw_to: codes.append("E05")
        if name in sw_from: codes.append("E06")
        # ---- 発注書 F（㉔）の材料：まとまりの合算と発注先はループのあとで決める（0.9.0）----
        #   注意のうち、ロット・出荷調整によるものは期待値では見ない（行の有無もそれに依存する分は Excel 側だけ）
        fcd = fc_by.get(name, {})
        blocks = defaultdict(float)
        for dd, (q, p) in tday.get(name, {}).items():
            blocks[next_biz(dd)] += q
        fmat[name] = dict(plan7=sum(v[0] for v in fcd.values()), pats7=sum(v[1] for v in fcd.values()),
                          stock_f=(fcd[max(fcd)][2] if fcd else (stock if stock is not None else 0)),
                          gap_own=d["ずれの割合"], plan_sum=plan_sum, gap_sum=gap_sum, blocks=blocks, cls=cls_d, trans=trans, src=src,
                          sw_to=sw_to, is_sw_from=(name in sw_from), rec_out=rec_out,
                          prev_out=prev_out, out_sum=out_sum, avg_m=out_sum / len(months), pats3=pats3)
        gr, gs = d["ずれの割合"], d["ずれの大きさ"]
        if gr is not None and gr >= TH["ずれの注意割合"]: codes.append("P01")
        if gr is not None and gr <= -TH["ずれの注意割合"]: codes.append("P02")
        if gs is not None and gs >= TH["ずれの大きさの注意割合"]: codes.append("P03")
        if trans: codes.append("E01" if (trans - BASE_DATE).days >= 0 else "E02")
        if src == "出納だけ": codes.append("Q01")
        if src == "予定だけ": codes.append("Q02")
        reasons += [dict(薬品名=name, コード=c) for c in codes]
    # ---- 発注の収束（0.9.0）：まとまり（屋号 2 つ以上）の発注先（⑨ と同じ決まり）----
    members = defaultdict(list)
    for name in fmat:
        if name in GROUP_OF:
            members[GROUP_OF[name]].append(name)

    def pick(cands):
        return sorted(cands, key=lambda n: (-fmat[n]["rec_out"], n))[0] if cands else None
    target_of = {}
    for gname, mem in members.items():
        if len(mem) < 2:
            continue
        t = (pick([n for n in mem if fmat[n]["is_sw_from"]])
             or pick([n for n in mem if not fmat[n]["sw_to"] and SUPPLY_OF.get(n) in (None, "通常") and fmat[n]["trans"] is None])
             or pick([n for n in mem if not fmat[n]["sw_to"]]))
        if t:
            for n in mem:
                target_of[n] = t
    for d in m_drug:
        d["発注先"] = target_of.get(d["薬品名"])
    # ---- 傾向（0.10.0）：直近 N ヶ月 vs その前 N ヶ月 vs 年間（⑨ と同じ決まり）----
    prev_m_all = months[-2 * RECENT_N:-RECENT_N] if len(months) >= 2 * RECENT_N else []
    prev_days_all = sum(month_end(m).day for m in prev_m_all)
    rec_days_all = sum(month_end(m).day for m in months[-RECENT_N:])
    enough = len(prev_m_all) == RECENT_N

    def judge(r1, r2, rp, pre):
        if pre is not None:
            return pre
        if r1 is None or r2 is None:
            return "増加" if rp > 0 else "横ばい"
        if r1 >= 1 + TH["傾向の線"] and r2 >= 1 + TH["傾向の線"]:
            return "増加"
        if r1 <= 1 - TH["傾向の線"] and r2 <= 1 - TH["傾向の線"]:
            return "減少"
        return "横ばい"

    def trend_of(rec, prv, al, avg_m):
        rp = rec / rec_days_all; pp = prv / prev_days_all if prev_days_all else 0; ap = al / cal_days
        r1 = rp / pp if pp > 0 else None; r2 = rp / ap if ap > 0 else None
        pre = "月不足" if not enough else "少量" if avg_m < TH["極小需要の月の払出"] else None
        return r1, r2, judge(r1, r2, rp, pre)

    def pat_trend(rec, prv, al):
        r1 = rec / prv if (prv is not None and prv > 0) else None; r2 = rec / al if al > 0 else None
        pre = "月不足" if prv is None else "少量" if al < TH["区分Dの患者数"] else None
        return r1, r2, judge(r1, r2, rec, pre)
    trend_of_name = {}
    for name, f in fmat.items():
        mem = members[GROUP_OF[name]] if name in GROUP_OF else [name]
        own = trend_of(f["rec_out"], f["prev_out"], f["out_sum"], f["avg_m"])
        grp = trend_of(sum(fmat[n]["rec_out"] for n in mem), sum(fmat[n]["prev_out"] for n in mem),
                       sum(fmat[n]["out_sum"] for n in mem), sum(fmat[n]["avg_m"] for n in mem))
        trend_of_name[name] = (own, grp, pat_trend(*f["pats3"]))
    for d in m_drug:
        own, grp, pat = trend_of_name[d["薬品名"]]
        d["前期比"], d["年間比"], d["傾向"] = own
        d["まとまりの傾向"], d["患者の傾向"] = grp[2], pat[2]
        if own[2] == "増加": reasons.append(dict(薬品名=d["薬品名"], コード="D07"))
        if own[2] == "減少": reasons.append(dict(薬品名=d["薬品名"], コード="D08"))
    # ---- 発注書 F（㉔）：まとまりは予定・在庫を合算し、発注先の行に目安を出す。ほかの屋号は「寄せる」----
    for name, f in fmat.items():
        tgt = target_of.get(name)
        pooled = tgt is not None
        is_t = (not pooled) or tgt == name
        mem = members[GROUP_OF[name]] if pooled else [name]
        g_plan = sum(fmat[n]["plan7"] for n in mem); g_stock = sum(fmat[n]["stock_f"] for n in mem)
        if pooled:
            w = sum(fmat[n]["plan_sum"] for n in mem)
            gap_v = (sum(fmat[n]["gap_sum"] for n in mem) / w) if w > 0 else 0
        else:
            gap_v = f["gap_own"] or 0
        factor = max(0, 1 + gap_v)
        blocks = defaultdict(float)
        for n in mem:
            for k, v in fmat[n]["blocks"].items():
                blocks[k] += v
        max_q = math.ceil(round(max(blocks.values(), default=0) * factor, 6))
        cls, trans, src, sw_to = f["cls"], f["trans"], f["src"], f["sw_to"]
        no_order = bool(sw_to)                           # 発注除外・供給停止はテストデータにない。切替元は発注しない
        careful = cls == "D" or trans is not None or gap_v <= -TH["ずれの注意割合"] or (f["is_sw_from"] and not pooled)
        oq = 0 if (no_order or not is_t) else math.ceil(round(max(0, g_plan * factor - g_stock), 6))
        over = is_t and cls == "A" and max_q > 0 and g_stock + oq > max_q
        note_det = (gap_v <= -TH["ずれの注意割合"]) or cls == "D" or trans is not None or src == "出納だけ" or over or sw_to or f["is_sw_from"]
        verdict = ("寄せる" if (pooled and not is_t) else "発注しない" if no_order
                   else "控えめに" if (oq > 0 and careful) else "発注する" if oq > 0 else "不要")
        short = max(0, g_plan - g_stock)
        tr = (trend_of_name[name][1] if pooled else trend_of_name[name][0])[2]
        if g_plan > 0 or short > 0 or note_det or cls == "A":
            orders.append(dict(薬品名=name, 区分=cls, 来週の予定=f["plan7"], 来週の患者=f["pats7"], 在庫=f["stock_f"],
                               まとまりの予定=(g_plan if pooled else None), まとまりの在庫=(g_stock if pooled else None),
                               不足=(short if is_t else None), 傾向=(tr if tr in ("増加", "減少") else None),
                               発注の目安=oq, 必要な最大在庫=(max_q if is_t else None), 判定=verdict))
    wb = Workbook(); wb.remove(wb.active)
    for title, data in (("T_月次", t_mon), ("M_薬品", m_drug), ("T_判定理由", reasons), ("T_ABC月次", abc_rows), ("R_発注書", orders)):
        ws = wb.create_sheet(title)
        cols = list(data[0].keys())
        ws.append(cols)
        for r in data:
            ws.append([r[c] for c in cols])
    ws = wb.create_sheet("説明")
    for t in ["M の定義どおりに Python で計算した期待値。Excel の結果と突き合わせる",
              "・T_月次：その月に出納に載ったか予定があった薬の行だけ（動きも予定もない月は行なし。⑦ と同じ）",
              "・M_薬品：列名の「一日」「CV2」は、Excel では「1日」「ばらつき（CV²）」",
              "・T_ABC月次：まとまり＝薬品名。ただしアリピプラゾール錠3mg(旧)（予定）と アリピプラゾール錠3mg（出納）は薬品グループで同じまとまり「アリピプラゾール 3mg」（(旧) が切替元）。量の区分は月平均払出の線 1000／300、区分 D は窓の各月の予定患者数の最少が 2 未満",
              "・R_発注書：来週＝10/2～10/8 の予測（202610 のファイル）。期待値にある行だけ比べる（ロット・出荷調整の注意だけで載る行は Excel 側だけ）。"
              "まとまり（アリピプラゾール 3mg、セルトラリン 25mg）は予定・在庫を合算し、発注先（切替先 → 通常供給で直近の払出が最多）の行に目安、ほかは「寄せる」（0.9.0）",
              "・祝日は無視して月～土を営業日にしている（テストデータの作り方の都合）",
              "・アリピプラゾールは実績の名前を「（旧）」付きにして、出納とずらしてある（Q01・Q02 の確認用）",
              "・スルピリドは全角の「５０ｍｇ」。fx_正規化 で半角になり、出納と一致すれば正しい"]:
        ws.append([t])
    wb.save(os.path.join(OUT, "期待値.xlsx"))


def to_xls():
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        print("soffice が見つからないので .xls は作らない"); return
    for sub in ("実績", "出納"):
        d = os.path.join(OUT, sub)
        subprocess.run([exe, "--headless", "--convert-to", 'xls:MS Excel 97', "--outdir", d]
                       + [os.path.join(d, f) for f in os.listdir(d) if f.endswith(".xlsx")], check=True,
                       stdout=subprocess.DEVNULL)
        for f in os.listdir(d):
            if f.endswith(".xlsx"):
                os.remove(os.path.join(d, f))


if __name__ == "__main__":
    main()
