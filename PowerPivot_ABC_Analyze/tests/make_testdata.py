"""PowerPivot_ABC のテストデータと期待値を作る（すべて架空のデータ）
使い方（tests フォルダで）：python make_testdata.py
  出力：tests/out/
    実績/薬品使用量分析データ2026MM.xlsx   … 薬品使用量分析データ（日ごとのシート）
    出納/2026MM01_2026MMDD_数量一覧.xlsx   … 出納の月次（その月に動いた薬だけ）
    期待値.xlsx                            … T_月次・M_薬品（主な指標）・T_判定理由の正解
    設定の値.txt                           … 分析用ブックの設定テーブルに入れる値
  LibreOffice（soffice）があれば、.xls も作る（--xls）。実物は .xls なので、本番前は .xls でも試す

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
TH = dict(不動とみなす月数=3, 在庫過多とみなす日数=60, 一箱消化の注意日数=180,
          ずれの注意割合=0.3, ずれの大きさの注意割合=0.5, 標準の包装単位=100)
WD = "月火水木金土日"

# 薬：名前・略称・型・1回あたりの量・処方の確率・実際/予定の倍率・臨時の確率
#   kind: regular（定期）, sparse（まばら）, dead（最初の月だけ動く）, ending（途中で処方終了）,
#         extra（予定に出ない＝出納だけ）, planonly（予定だけ＝出納と名前がずれる）
DRUGS = [
    ("デエビゴ錠2.5mg",            "ﾃﾞｴﾋﾞｺﾞ2.5", "regular",  28, 0.95, 1.00, 0.10),
    ("セルトラリン錠25mg「アメル」", "ｾﾙﾄﾗﾘﾝ25",   "regular",  14, 0.80, 1.05, 0.15),
    ("エスシタロプラム錠10mg",       "ｴｽｼﾀﾛ10",    "regular",  28, 0.60, 1.00, 0.05),
    ("タンドスピロン錠10mg「サワイ」", "ﾀﾝﾄﾞｽﾋﾟﾛﾝ10", "regular", 42, 0.30, 1.40, 0.20),
    ("ビペリデン錠1mg",              "ﾋﾞﾍﾟﾘﾃﾞﾝ1",  "sparse",   28, 0.08, 1.00, 0.05),
    ("リチウム炭酸塩錠200mg",        "ﾘﾁｳﾑ200",    "sparse",   84, 0.05, 1.00, 0.00),
    ("ハロペリドール錠0.75mg「アメル」", "ﾊﾛﾍﾟﾘ0.75", "dead",    28, 0.10, 1.00, 0.00),
    ("クロキサゾラム錠1mg",          "ｸﾛｷｻｿﾞﾗﾑ1",  "dead",     14, 0.10, 1.00, 0.00),
    ("ミルタザピン錠15mg",           "ﾐﾙﾀｻﾞﾋﾟﾝ15", "ending",   28, 0.50, 1.00, 0.00),
    ("ロラゼパム錠0.5mg",            "ﾛﾗｾﾞﾊﾟﾑ0.5", "extra",     7, 0.00, 1.00, 0.10),
    ("アリピプラゾール錠3mg",        "ｱﾘﾋﾟﾌﾟﾗ3",   "planonly", 28, 0.40, 1.00, 0.00),
    ("スルピリド錠５０ｍｇ",          "ｽﾙﾋﾟﾘﾄﾞ50",  "regular",  28, 0.40, 0.95, 0.05),
]
TRANSITIONAL = {"ハロペリドール錠0.75mg「アメル」": dt.date(2027, 3, 31)}   # 経過措置リスト


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
            elif rnd.random() < p:
                pts = rnd.randint(1, 3) if kind == "regular" else 1
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
        if merged_second:   # 実物：1行目が見出し、2行目は空欄で1行目と縦に結合
            for c in range(1, len(rows[0]) + 1):
                ws.merge_cells(start_row=1, start_column=c, end_row=2, end_column=c)
        for c in ws[1]:
            c.font = Font(bold=True)
    wb.save(path)


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

    expected(days, plan, mon, info, plan_name)
    with open(os.path.join(OUT, "設定の値.txt"), "w", encoding="utf-8") as f:
        f.write("分析用ブックの設定テーブルに入れる値（ほかは空欄＝初期値）\n"
                f"フォルダパス\t{H}\n出納フォルダパス\t{S}\n開始日\t{START:%Y/%m/%d}\n終了日\t{END:%Y/%m/%d}\n"
                f"発注日\t{BASE_DATE:%Y/%m/%d}\n\n経過措置リスト：\n")
        for n, d in TRANSITIONAL.items():
            f.write(f"{n}\t{d:%Y/%m/%d}\n")
    if "--xls" in sys.argv:
        to_xls()
    print("できた：", OUT)


def expected(days, plan, mon, info, plan_name):
    """M の定義どおりに期待値を計算する（T_月次・M_薬品の主な指標・T_判定理由）"""
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
                                if r["仕入数量"] or r["払出数量"] or r["その他の増減"]})
    raw = {nz(n): n for n, *_ in DRUGS}
    t_mon, m_drug, reasons = [], [], []
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
        if not any(x["出納に載った"] for x in rows) and not any(x["予定数量"] for x in rows):
            continue
        t_mon += [x for x in rows if x["出納に載った"] or x["予定数量"] or name in sui_names]
        q = [v[0] for v in tday.get(name, {}).values() if v[0] > 0]
        n = len(q)
        adi = day_n / n if n else None
        cv2 = (statistics.stdev(q) / statistics.mean(q)) ** 2 if n >= 2 else None
        typ = None if adi is None else ("安定" if adi < 1.32 else "間欠") if cv2 is None else \
            (("安定" if cv2 < 0.49 else "ばらつき大") if adi < 1.32 else ("間欠" if cv2 < 0.49 else "まばら・大口"))
        moved = [x for x in rows if x["出納に載った"]]
        out_sum = sum(x["払出数量"] for x in rows)
        per_day = out_sum / cal_days
        last = max(moved, key=lambda x: x["月"]) if moved else None
        still = (latest.year - last["月"].year) * 12 + latest.month - last["月"].month if last else None
        stock = last["在庫数量"] if last else None
        plan_sum = sum(x["予定数量"] for x in rows)
        gap_sum = sum(x["予定とのずれ"] for x in rows)
        abs_sum = sum(abs(x["予定とのずれ"]) for x in rows)
        extra_m = sum(1 for x in rows if x["予定数量"] == 0 and x["払出数量"] > 0)
        less = sum(-x["その他の増減"] for x in rows if x["その他の増減"] < 0)
        src = "予定・出納" if n and moved else "予定だけ" if n else "出納だけ"
        trans = TRANSITIONAL.get(raw.get(name, ""))
        d = dict(薬品名=name, 出どころ=src, 予定日数=n, 平均間隔=adi, CV2=cv2, 需要の型=typ, 動いた月数=len(moved),
                 払出合計=out_sum, 一日平均払出=per_day, 最後に動いた月=last["月"] if last else None, 最終在庫=stock,
                 動きのない月数=still, 在庫日数=(stock / per_day if stock is not None and per_day > 0 else None),
                 一箱消化日数=(TH["標準の包装単位"] / per_day if per_day > 0 else None),
                 その他の減少=less, ずれの割合=(gap_sum / plan_sum if plan_sum > 0 else None),
                 ずれの大きさ=(abs_sum / plan_sum if plan_sum > 0 else None), 予定外の月数=extra_m, 経過措置期限=trans)
        m_drug.append(d)
        codes = []
        if (stock or 0) > 0 and still is not None and still >= TH["不動とみなす月数"]: codes.append("S01")
        if d["在庫日数"] is not None and d["在庫日数"] >= TH["在庫過多とみなす日数"]: codes.append("S02")
        if d["一箱消化日数"] is not None and d["一箱消化日数"] >= TH["一箱消化の注意日数"]: codes.append("S03")
        if less > 0: codes.append("S04")
        codes += {"安定": ["D01"], "ばらつき大": ["D02"], "間欠": ["D03"], "まばら・大口": ["D04"]}.get(typ, [])
        if extra_m >= 1: codes.append("D05")
        gr, gs = d["ずれの割合"], d["ずれの大きさ"]
        if gr is not None and gr >= TH["ずれの注意割合"]: codes.append("P01")
        if gr is not None and gr <= -TH["ずれの注意割合"]: codes.append("P02")
        if gs is not None and gs >= TH["ずれの大きさの注意割合"]: codes.append("P03")
        if trans: codes.append("E01" if (trans - BASE_DATE).days >= 0 else "E02")
        if src == "出納だけ": codes.append("Q01")
        if src == "予定だけ": codes.append("Q02")
        reasons += [dict(薬品名=name, コード=c) for c in codes]
    wb = Workbook(); wb.remove(wb.active)
    for title, data in (("T_月次", t_mon), ("M_薬品", m_drug), ("T_判定理由", reasons)):
        ws = wb.create_sheet(title)
        cols = list(data[0].keys())
        ws.append(cols)
        for r in data:
            ws.append([r[c] for c in cols])
    ws = wb.create_sheet("説明")
    for t in ["M の定義どおりに Python で計算した期待値。Excel の結果と突き合わせる",
              "・T_月次：Excel では出納のある月にすべての薬の行が出る（動かなかった月は払出0）。ここは主な行だけ",
              "・M_薬品：列名の「一日」「一箱」「CV2」は、Excel では「1日」「1箱」「ばらつき（CV²）」",
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
