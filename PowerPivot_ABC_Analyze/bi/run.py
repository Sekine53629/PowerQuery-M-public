"""BI ツール（Python 版）：データモデルを組み、月次レポート R1～R4 と統計 K0～K6 を作って Excel と HTML に出す
使い方（プロジェクトのルートで）：
  python bi/run.py --tables tests/out_clinic/export/手入力テーブル.xlsx [--out bi/out]
  出力：bi/out/モデル.xlsx（T_日次・T_月次・M_日付・M_薬品・M_理由コード・T_判定理由・各レポート）
        bi/out/ダッシュボード.html（グラフと表）
※ ここでの定義が「正解」。Power Pivot 側（M クエリ・ピボット）を同じ結果になるよう合わせていく
"""
import os, sys, math, argparse, datetime as dt, html
from collections import defaultdict
import pandas as pd
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from model import Model, Settings, month_start, month_end, months_between, fmt0, REASON_CODES


# ------------------------------------------------------------------------------
# ABC 区分（需要のまとまり単位。未登録の薬は薬品名単体）
# ------------------------------------------------------------------------------
def abc_window(t_month, m_drug, end_month, n_months, a_line, b_line):
    """end_month を含む直近 n_months の払出で ABC を付ける。戻り値：まとまり → (払出, 構成比, 累積, 区分)"""
    start = dt.date(end_month.year, end_month.month, 1)
    for _ in range(n_months - 1):
        start = month_start(start - dt.timedelta(days=1))
    t = t_month[(t_month["月"] >= start) & (t_month["月"] <= end_month) & t_month["出納データあり"]]
    cl = dict(zip(m_drug["薬品名"], m_drug["需要のまとまり"]))
    key = t["薬品名"].map(lambda n: cl.get(n) or n)
    s = t.groupby(key)["払出数量"].sum().sort_values(ascending=False)
    total = float(s.sum())
    res, cum = {}, 0.0
    for k, v in s.items():
        share = (v / total) if total else 0.0
        cum += share
        res[k] = (float(v), share, cum, "A" if cum <= a_line else "B" if cum <= b_line else "C")
    return res


def abc_monthly(model, settings):
    """T_ABC月次：月 × まとまり の区分（その月を終わりとする直近 n ヶ月）"""
    n, a, b = int(settings.num("ABCの月数")), settings.num("ABCのAの線"), settings.num("ABCのBの線")
    months = sorted(set(model.T_月次.loc[model.T_月次["出納データあり"], "月"]))
    rows = []
    for m in months:
        for k, (v, sh, cu, cls) in abc_window(model.T_月次, model.M_薬品, m, n, a, b).items():
            rows.append(dict(月=m, 需要のまとまり=k, 払出=v, 構成比=sh, 累積=cu, 区分=cls))
    return pd.DataFrame(rows, columns=["月", "需要のまとまり", "払出", "構成比", "累積", "区分"])


# ------------------------------------------------------------------------------
# 月次レポート R1～R4
# ------------------------------------------------------------------------------
def r1_abc(model, settings, abc_m):
    months = sorted(set(abc_m["月"]))
    if not months:
        return pd.DataFrame()
    cur, prev = months[-1], (months[-2] if len(months) >= 2 else None)
    c = abc_m[abc_m["月"] == cur].set_index("需要のまとまり")
    p = abc_m[abc_m["月"] == prev].set_index("需要のまとまり") if prev else None
    d = model.M_薬品
    rows = []
    for _, r in d.iterrows():
        k = r["需要のまとまり"] if isinstance(r["需要のまとまり"], str) else r["薬品名"]
        if k not in c.index:
            continue
        now = c.loc[k]
        before = p.loc[k, "区分"] if (p is not None and k in p.index) else None
        rows.append(dict(需要のまとまり=k, 薬品名=r["薬品名"], 略称=r["略称"], 先発後発=r["先発後発"],
                         直近の払出=now["払出"], 構成比=now["構成比"], 累積=now["累積"], 今月の区分=now["区分"], 前月の区分=before,
                         変更=("" if before in (None, now["区分"]) else f"{before}→{now['区分']}"),
                         **{"1箱消化日数": r["1箱消化日数"]}, 在庫日数=r["在庫日数"], 需要の型=r["需要の型"], 最終在庫=r["最終在庫"]))
    out = pd.DataFrame(rows)
    return out.sort_values(["累積", "需要のまとまり", "薬品名"]).reset_index(drop=True)


def r2_dead(model, settings):
    d, rs = model.M_薬品, model.T_判定理由
    horizon = settings.num("当面の予定日数")
    s_codes = rs[rs["コード"].isin(["S01", "S02", "S03", "S05", "S06"])]
    if s_codes.empty:
        return pd.DataFrame()
    msgs = s_codes.groupby("薬品名")["文言"].apply(lambda x: "／".join(x))
    rows = []
    for _, r in d[d["薬品名"].isin(msgs.index)].iterrows():
        per_day = r["直近の1日平均払出"] if not pd.isna(r["直近の1日平均払出"]) else 0.0
        stock = r["最終在庫"] if not pd.isna(r["最終在庫"]) else 0.0
        pack = r["包装単位"] if not pd.isna(r["包装単位"]) and r["包装単位"] > 0 else settings.num("標準の包装単位")
        surplus = max(0.0, stock - per_day * horizon)
        rows.append(dict(薬品名=r["薬品名"], 略称=r["略称"], 棚番=r["棚番"], 最終在庫=stock, 最後に動いた月=r["最後に動いた月"],
                         動きのない月数=r["動きのない月数"], 在庫日数=r["在庫日数"], 直近の在庫日数=r["直近の在庫日数"],
                         **{"1箱消化日数": r["1箱消化日数"], "直近の1日平均払出": per_day},
                         余った量=surplus, 返品できる箱数=math.floor(surplus / pack) if pack else None,
                         需要のまとまり=r["需要のまとまり"], 警告=msgs[r["薬品名"]]))
    return pd.DataFrame(rows).sort_values(["動きのない月数", "在庫日数"], ascending=[False, False], na_position="last").reset_index(drop=True)


def r3_expiry(model, settings):
    d, rs = model.M_薬品, model.T_判定理由
    base = settings.base_date
    near = settings.num("期限切迫の日数")
    e_msgs = rs[rs["コード"].isin(["E01", "E02"])].groupby("薬品名")["文言"].apply(lambda x: "／".join(x))
    rows = []
    for _, r in d[d["経過措置期限"].notna()].iterrows():
        left = (r["経過措置期限"] - base).days
        per_day = r["1日平均払出"] if not pd.isna(r["1日平均払出"]) else 0.0
        stock = r["最終在庫"] if not pd.isna(r["最終在庫"]) else 0.0
        usable = per_day * max(0, left)
        rows.append(dict(薬品名=r["薬品名"], 略称=r["略称"], 経過措置期限=r["経過措置期限"], 残り日数=left, 期限切迫=(left <= near),
                         最終在庫=stock, **{"1日平均払出": per_day}, 期限までに使い切れる量=usable, 余る量=max(0.0, stock - usable),
                         警告=e_msgs.get(r["薬品名"], "")))
    return pd.DataFrame(rows).sort_values("残り日数").reset_index(drop=True)


def r4_shortage(model, settings):
    """出荷調整の可能性：同じまとまりの中で屋号の構成比が大きく入れ替わった月（G01）、
       仕入がないまま払出が続いて在庫が尽きかけた月（G02）"""
    t, d = model.T_月次, model.M_薬品
    pt, tol = settings.num("出荷調整の構成比の変化"), settings.num("出荷調整の合計の許容")
    cl = dict(zip(d["薬品名"], d["需要のまとまり"]))
    tt = t[t["出納データあり"]].copy()
    tt["まとまり"] = tt["薬品名"].map(lambda n: cl.get(n))
    tt = tt[tt["まとまり"].notna()]
    rows = []
    for k, g in tt.groupby("まとまり"):
        if g["薬品名"].nunique() < 2:
            continue
        piv = g.pivot_table(index="月", columns="薬品名", values="払出数量", aggfunc="sum").fillna(0).sort_index()
        tot = piv.sum(axis=1)
        share = piv.div(tot.replace(0, np.nan), axis=0).fillna(0)
        months = list(piv.index)
        for i in range(1, len(months)):
            m0, m1 = months[i - 1], months[i]
            if tot[m0] <= 0:
                continue
            tot_change = (tot[m1] - tot[m0]) / tot[m0]
            for n in piv.columns:
                delta = share.loc[m1, n] - share.loc[m0, n]
                if abs(delta) >= pt and abs(tot_change) <= tol:
                    rows.append(dict(需要のまとまり=k, 薬品名=n, 月=m1, コード="G01",
                                     前月の構成比=share.loc[m0, n], 当月の構成比=share.loc[m1, n], 変化=delta,
                                     まとまりの払出_前月=tot[m0], まとまりの払出_当月=tot[m1],
                                     理由=f"構成比が {round(delta*100):+d} ポイント動きました（まとまり全体は {round(tot_change*100):+d}%）"))
    # G02：仕入 0 が 2 ヶ月以上続き、払出があり、在庫が月の払出を下回った（月に 1 箱未満の薬は除く：箱単位の発注なら仕入 0 の月は普通）
    pack_of = dict(zip(d["薬品名"], d["包装単位"].fillna(settings.num("標準の包装単位"))))
    for n, g in tt.sort_values("月").groupby("薬品名"):
        g = g.reset_index(drop=True)
        for i in range(1, len(g)):
            a, b = g.loc[i - 1], g.loc[i]
            if (a["仕入数量"] or 0) == 0 and (b["仕入数量"] or 0) == 0 and (b["払出数量"] or 0) >= pack_of.get(n, 100) and (a["払出数量"] or 0) > 0 \
                    and b["在庫数量"] is not None and not pd.isna(b["在庫数量"]) and b["在庫数量"] < b["払出数量"]:
                rows.append(dict(需要のまとまり=cl.get(n), 薬品名=n, 月=b["月"], コード="G02", 前月の構成比=None, 当月の構成比=None, 変化=None,
                                 まとまりの払出_前月=None, まとまりの払出_当月=None,
                                 理由=f"仕入がないまま払出が続き、在庫 {fmt0(b['在庫数量'])}（月の払出 {fmt0(b['払出数量'])}）"))
    cols = ["需要のまとまり", "薬品名", "月", "コード", "前月の構成比", "当月の構成比", "変化", "まとまりの払出_前月", "まとまりの払出_当月", "理由"]
    out = pd.DataFrame(rows, columns=cols)
    return out.sort_values(["月", "需要のまとまり", "薬品名"], ascending=[False, True, True]).reset_index(drop=True)


# ------------------------------------------------------------------------------
# 統計 K0～K6
# ------------------------------------------------------------------------------
STOCK_BINS = [(0, 30, "～30日"), (30, 60, "30～60日"), (60, 90, "60～90日"), (90, 180, "90～180日"), (180, 365, "180～365日"), (365, 1e18, "365日超")]


def stock_bin(days):
    if days is None or (isinstance(days, float) and math.isnan(days)):
        return "払出なし"
    for lo, hi, lab in STOCK_BINS:
        if lo <= days < hi:
            return lab
    return "365日超"


def k0_summary(model, settings, abc_m):
    t = model.T_月次[model.T_月次["出納データあり"]].copy()
    still_m = int(settings.num("不動とみなす月数"))
    months = sorted(set(t["月"]))
    last_seen, last_stock = {}, {}
    cls_prev = None
    rows = []
    for m in months:
        g = t[t["月"] == m]
        for _, r in g[g["出納に載った"]].iterrows():
            last_seen[r["薬品名"]] = m
            if not pd.isna(r["在庫数量"]):
                last_stock[r["薬品名"]] = r["在庫数量"]
        dead = [n for n, lm in last_seen.items() if months_between(lm, m) >= still_m and last_stock.get(n, 0) > 0]
        a = abc_m[abc_m["月"] == m].set_index("需要のまとまり")["区分"]
        changed = int(sum(1 for k in a.index if cls_prev is not None and k in cls_prev.index and cls_prev[k] != a[k]))
        cls_prev = a
        rows.append(dict(月=m, 払出合計=float(g["払出数量"].sum()), 仕入合計=float(g["仕入数量"].sum()),
                         ロス量=float((-g["その他の増減"].fillna(0)).clip(lower=0).sum()),
                         在庫数量合計=float(sum(last_stock.values())), 動いた薬品数=int(g["出納に載った"].sum()),
                         不動薬の数=len(dead), 不動薬の在庫=float(sum(last_stock[n] for n in dead)),
                         予定患者数合計=float(g["予定患者数"].fillna(0).sum()), ABC変更数=changed,
                         A=int((a == "A").sum()), B=int((a == "B").sum()), C=int((a == "C").sum())))
    return pd.DataFrame(rows)


def k1_stock_bins(model):
    d = model.M_薬品.copy()
    d["在庫日数の帯"] = d["直近の在庫日数"].map(stock_bin)       # 直近 N ヶ月の払出で見た在庫日数
    order = ["払出なし"] + [b[2] for b in STOCK_BINS]
    g = d.groupby("在庫日数の帯").agg(薬品数=("薬品名", "count"), 在庫数量=("最終在庫", "sum")).reindex(order).fillna(0).reset_index()
    return g


def k2_quadrant(model, settings):
    d = model.M_薬品.copy()
    box_d = settings.num("1箱消化の注意日数")
    d["回転"] = d["直近の1箱消化日数"].map(lambda x: "払出なし" if pd.isna(x) else ("回らない" if x >= box_d else "回る"))
    d["型"] = d["需要の型"].fillna("予定なし")
    pv = d.pivot_table(index="型", columns="回転", values="薬品名", aggfunc="count").fillna(0).astype(int)
    pv2 = d.pivot_table(index="型", columns="回転", values="最終在庫", aggfunc="sum").fillna(0)
    return pv, pv2


def k4_loss(model):
    t = model.T_月次[model.T_月次["出納データあり"]].copy()
    t["ロス"] = (-t["その他の増減"].fillna(0)).clip(lower=0)
    t["受け入れ"] = t["その他の増減"].fillna(0).clip(lower=0)
    bym = t.groupby("月").agg(ロス=("ロス", "sum"), 受け入れ=("受け入れ", "sum")).reset_index()
    top = t.groupby("薬品名")["ロス"].sum().sort_values(ascending=False).head(15).reset_index()
    return bym, top


def k6_share(model):
    t, d = model.T_月次, model.M_薬品
    cl = dict(zip(d["薬品名"], d["需要のまとまり"]))
    tt = t[t["出納データあり"]].copy()
    tt["まとまり"] = tt["薬品名"].map(lambda n: cl.get(n))
    tt = tt[tt["まとまり"].notna()]
    multi = tt.groupby("まとまり")["薬品名"].nunique()
    tt = tt[tt["まとまり"].isin(multi[multi >= 2].index)]
    g = tt.groupby(["まとまり", "月", "薬品名"])["払出数量"].sum().reset_index()
    tot = g.groupby(["まとまり", "月"])["払出数量"].transform("sum")
    g["構成比"] = (g["払出数量"] / tot.replace(0, np.nan)).fillna(0)
    return g


# ------------------------------------------------------------------------------
# 出力
# ------------------------------------------------------------------------------
def write_excel(path, sheets):
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for name, df in sheets:
            df2 = df.copy()
            if isinstance(df2, pd.Series):
                df2 = df2.to_frame()
            df2.to_excel(xw, sheet_name=name[:31], index=not isinstance(df2.index, pd.RangeIndex))
            ws = xw.sheets[name[:31]]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                w = max((len(str(c.value)) if c.value is not None else 0) for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(8, int(w * 1.4) + 2), 60)


def df_html(df, n=None, pct=(), int_cols=()):
    if df is None or len(df) == 0:
        return "<p class='empty'>該当なし</p>"
    d = df.head(n) if n else df
    out = ["<table><thead><tr>"] + [f"<th>{html.escape(str(c))}</th>" for c in d.columns] + ["</tr></thead><tbody>"]
    for _, r in d.iterrows():
        out.append("<tr>")
        for c in d.columns:
            v = r[c]
            if v is None or (isinstance(v, float) and math.isnan(v)):
                s = ""
            elif c in pct:
                s = f"{v*100:.1f}%"
            elif isinstance(v, (dt.date, dt.datetime)):
                s = v.strftime("%Y/%m") if (c in ("月", "最後に動いた月")) else v.strftime("%Y/%m/%d")
            elif isinstance(v, bool) or isinstance(v, np.bool_):
                s = "○" if v else ""
            elif isinstance(v, (int, np.integer)) or c in int_cols:
                s = f"{int(round(v)):,}"
            elif isinstance(v, (float, np.floating)):
                s = f"{v:,.1f}" if abs(v) < 100 else f"{v:,.0f}"
            else:
                s = html.escape(str(v))
            cls = " class='num'" if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool) else ""
            out.append(f"<td{cls}>{s}</td>")
        out.append("</tr>")
    out.append("</tbody></table>")
    if n and len(df) > n:
        out.append(f"<p class='note'>先頭 {n} 行のみ表示（全 {len(df)} 行は モデル.xlsx）</p>")
    return "".join(out)


def build_dashboard(path, model, settings, R, K):
    import plotly.graph_objects as go
    import plotly.io as pio
    k0, k1, (k2n, k2s), (k4m, k4top), k6 = K["K0"], K["K1"], K["K2"], K["K4"], K["K6"]
    figs = []
    # K0
    f = go.Figure()
    f.add_bar(x=k0["月"], y=k0["払出合計"], name="払出")
    f.add_bar(x=k0["月"], y=k0["仕入合計"], name="仕入")
    f.add_scatter(x=k0["月"], y=k0["ロス量"], name="ロス（返品・廃棄・移動）", yaxis="y2", mode="lines+markers")
    f.update_layout(title="月ごとの払出・仕入・ロス", barmode="group", yaxis2=dict(overlaying="y", side="right", title="ロス"), height=380)
    figs.append(("K0 月次の動き", f))
    f = go.Figure()
    f.add_scatter(x=k0["月"], y=k0["在庫数量合計"], name="在庫数量の合計（最後に載った在庫の合計）", mode="lines+markers")
    f.add_scatter(x=k0["月"], y=k0["不動薬の在庫"], name="不動薬の在庫", mode="lines+markers")
    f.add_bar(x=k0["月"], y=k0["不動薬の数"], name="不動薬の数", yaxis="y2", opacity=0.4)
    f.update_layout(title="在庫と不動薬の推移", yaxis2=dict(overlaying="y", side="right", title="薬品数"), height=380)
    figs.append(("K3 不動・過剰の推移", f))
    # K1
    f = go.Figure()
    f.add_bar(x=k1["在庫日数の帯"], y=k1["薬品数"], name="薬品数")
    f.add_bar(x=k1["在庫日数の帯"], y=k1["在庫数量"], name="在庫数量", yaxis="y2", opacity=0.5)
    f.update_layout(title="在庫日数の分布（直近の払出で、今の在庫が何日もつか）", yaxis2=dict(overlaying="y", side="right"), barmode="group", height=360)
    figs.append(("K1 在庫日数の分布", f))
    # K4
    f = go.Figure()
    f.add_bar(x=k4m["月"], y=k4m["ロス"], name="ロス（減少）")
    f.add_bar(x=k4m["月"], y=k4m["受け入れ"], name="受け入れ（増加）")
    f.update_layout(title="その他の増減の推移", barmode="group", height=340)
    figs.append(("K4 ロス集計", f))
    # K6：切替のあったまとまりを優先（構成比の変化が大きい順に上位 4 つ）
    var = k6.groupby(["まとまり", "薬品名"])["構成比"].agg(lambda x: x.max() - x.min()).groupby("まとまり").max().sort_values(ascending=False)
    for k in list(var.index[:4]):
        g = k6[k6["まとまり"] == k]
        f = go.Figure()
        for n, gg in g.groupby("薬品名"):
            f.add_bar(x=gg["月"], y=gg["構成比"], name=n)
        f.update_layout(title=f"屋号の構成比：{k}", barmode="stack", yaxis=dict(tickformat=".0%"), height=320)
        figs.append((f"K6 {k}", f))
    parts = []
    for i, (title, f) in enumerate(figs):
        parts.append(pio.to_html(f, include_plotlyjs=("cdn" if i == 0 else False), full_html=False, config={"displaylogo": False}))
    pct = ("構成比", "累積", "前月の構成比", "当月の構成比", "変化")
    css = """
    body{font-family:'Yu Gothic UI','Meiryo',sans-serif;margin:24px;color:#222;background:#fafafa}
    h1{font-size:22px} h2{font-size:17px;border-left:6px solid #3b6ea5;padding-left:8px;margin-top:36px}
    table{border-collapse:collapse;font-size:12px;background:#fff} th,td{border:1px solid #ccc;padding:3px 6px;white-space:nowrap}
    th{background:#e9eef5} td.num{text-align:right} .note,.empty{color:#666;font-size:12px}
    .kpi{display:flex;gap:16px;flex-wrap:wrap} .kpi div{background:#fff;border:1px solid #ddd;padding:10px 16px;min-width:140px}
    .kpi b{display:block;font-size:20px} .wrap{overflow-x:auto}
    """
    last = k0.iloc[-1] if len(k0) else None
    kpis = ""
    if last is not None:
        rs = model.T_判定理由
        kpis = "<div class='kpi'>" + "".join(
            f"<div>{lab}<b>{val}</b></div>" for lab, val in [
                ("対象月", last["月"].strftime("%Y/%m")), ("在庫数量の合計", f"{last['在庫数量合計']:,.0f}"),
                ("不動薬の数", f"{int(last['不動薬の数'])}"), ("不動薬の在庫", f"{last['不動薬の在庫']:,.0f}"),
                ("今月のロス", f"{last['ロス量']:,.0f}"), ("ABC 変更", f"{int(last['ABC変更数'])}"),
                ("警告（優先度1）", f"{int((rs['優先度'] == 1).sum())}"), ("注意（優先度2）", f"{int((rs['優先度'] == 2).sum())}"),
                ("薬品数", f"{len(model.M_薬品)}")]) + "</div>"
    body = [f"<h1>在庫分析ダッシュボード（Python 版・架空データ）</h1>",
            f"<p class='note'>期間 {settings.start:%Y/%m/%d}～{settings.end:%Y/%m/%d}　基準日 {settings.base_date:%Y/%m/%d}　作成 {dt.datetime.now():%Y/%m/%d %H:%M}</p>",
            kpis, "<h2>K0 月次サマリー</h2>", "<div class='wrap'>" + df_html(k0, int_cols=("払出合計", "仕入合計", "ロス量", "在庫数量合計", "不動薬の在庫", "予定患者数合計")) + "</div>"]
    body += [f"<h2>{t}</h2>{p}" for (t, _), p in zip(figs, parts)]
    body += ["<h2>K2 4 象限（需要の型 × 回転。回転は直近の 1 箱消化日数）：薬品数</h2>", df_html(k2n.reset_index()),
             "<h2>K2 4 象限：最終在庫の合計</h2>", df_html(k2s.reset_index(), int_cols=tuple(k2s.columns)),
             "<h2>R1 ABC 区分表（直近 %d ヶ月の払出、まとまり単位）</h2>" % int(settings.num("ABCの月数")), "<div class='wrap'>" + df_html(R["R1"], n=60, pct=pct) + "</div>",
             "<h2>R2 不動薬の処理候補</h2>", "<div class='wrap'>" + df_html(R["R2"], n=60) + "</div>",
             "<h2>R3 経過措置・期限切迫</h2>", "<div class='wrap'>" + df_html(R["R3"]) + "</div>",
             "<h2>R4 出荷調整の可能性</h2>", "<div class='wrap'>" + df_html(R["R4"], n=60, pct=pct) + "</div>",
             "<h2>T_判定理由（警告・注意のみ、先頭 80 行）</h2>", "<div class='wrap'>" + df_html(model.T_判定理由[model.T_判定理由["優先度"] <= 2], n=80) + "</div>",
             "<h2>K4 ロスの多い薬</h2>", df_html(k4top, int_cols=("ロス",))]
    doc = f"<!DOCTYPE html><html lang='ja'><head><meta charset='utf-8'><title>在庫分析ダッシュボード</title><style>{css}</style></head><body>{''.join(body)}</body></html>"
    with open(path, "w", encoding="utf-8") as f:
        f.write(doc)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tables", required=True, help="手入力テーブル.xlsx（設定ほか）")
    ap.add_argument("--out", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "out"))
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    settings = Settings(a.tables)
    model = Model(settings)
    abc_m = abc_monthly(model, settings)
    R = dict(R1=r1_abc(model, settings, abc_m), R2=r2_dead(model, settings), R3=r3_expiry(model, settings), R4=r4_shortage(model, settings))
    K = dict(K0=k0_summary(model, settings, abc_m), K1=k1_stock_bins(model), K2=k2_quadrant(model, settings), K4=k4_loss(model), K6=k6_share(model))
    sheets = [("T_日次", model.T_日次), ("T_月次", model.T_月次), ("M_日付", model.M_日付), ("M_薬品", model.M_薬品),
              ("M_理由コード", model.M_理由コード), ("T_判定理由", model.T_判定理由), ("T_ロット", model.T_ロット), ("T_ABC月次", abc_m),
              ("R1_ABC区分表", R["R1"]), ("R2_不動薬の処理候補", R["R2"]), ("R3_経過措置・期限切迫", R["R3"]), ("R4_出荷調整の可能性", R["R4"]),
              ("K0_月次サマリー", K["K0"]), ("K1_在庫日数の分布", K["K1"]), ("K2_4象限_薬品数", K["K2"][0]), ("K2_4象限_在庫", K["K2"][1]),
              ("K4_ロス_月別", K["K4"][0]), ("K4_ロス_薬品別", K["K4"][1]), ("K6_屋号の構成比", K["K6"]), ("出納", model.出納)]
    xp = os.path.join(a.out, "モデル.xlsx")
    write_excel(xp, sheets)
    hp = os.path.join(a.out, "ダッシュボード.html")
    build_dashboard(hp, model, settings, R, K)
    print(f"日次データ {len(model.日次データ)} 行　出納 {len(model.出納)} 行")
    print(f"T_日次 {len(model.T_日次)}　T_月次 {len(model.T_月次)}　M_日付 {len(model.M_日付)}　M_薬品 {len(model.M_薬品)}　T_判定理由 {len(model.T_判定理由)}")
    print(f"システムリスト：経過措置 {len(model.経過措置該当)}　期限切れ {len(model.期限切れリスト)}　不動品 {len(model.不動品リスト)}　→ T_ロット {len(model.T_ロット)}（返す量 {model.T_ロット['返す量'].sum() if len(model.T_ロット) else 0:,.0f}）")
    print("理由コード別：", model.T_判定理由["コード"].value_counts().sort_index().to_dict())
    print("出どころ：", model.M_薬品["出どころ"].value_counts().to_dict(), "　需要の型：", model.M_薬品["需要の型"].value_counts(dropna=False).to_dict())
    print(f"R1 {len(R['R1'])} 行　R2 {len(R['R2'])} 行　R3 {len(R['R3'])} 行　R4 {len(R['R4'])} 行")
    print("出力：", xp, "／", hp)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    main()
