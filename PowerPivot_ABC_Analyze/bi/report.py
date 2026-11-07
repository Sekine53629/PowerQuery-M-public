"""監視レポート（A4 横・5 ページ）：毎月「何をすべきか」が一目で分かる形に絞ったレポート
  A 発注点の切り替え提案　B 不動在庫と転送・返品推奨　C 経過措置・期限切迫　D 出荷調整の管理票　E 極小需要薬のビュー
使い方（プロジェクトのルートで）：
  python bi/report.py --tables tests/out_clinic/export/手入力テーブル.xlsx [--out bi/out] [--pdf]
  出力：bi/out/監視レポート.html（--pdf で 監視レポート.pdf も。Edge のヘッドレス印刷を使う）
"""
import os, sys, math, html, argparse, subprocess, datetime as dt
from collections import defaultdict
import pandas as pd
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from model import Model, Settings, month_start, month_end, months_between, fmt0, DEFAULTS
from run import abc_monthly, r3_expiry, r4_shortage, k6_share, r5_order

DEFAULTS.update(発注点の日数=7, 発注量の日数=21, 転送候補の在庫日数=120, 予測できるずれの線=0.5, 発注点見直しの変化=0.2)
EDGE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"


# ------------------------------------------------------------------------------
# 小さな SVG 部品（印刷で崩れないよう JS を使わない）
# ------------------------------------------------------------------------------
def esc(s):
    return html.escape("" if s is None else str(s))


def spark(values, w=130, h=26, color="#3b6ea5", mark_last=3):
    v = [0 if (x is None or (isinstance(x, float) and math.isnan(x))) else float(x) for x in values]
    if not v:
        return ""
    mx = max(v) or 1.0
    n = len(v)
    bw = w / n
    rects = []
    for i, x in enumerate(v):
        bh = (x / mx) * (h - 2)
        c = "#e0772b" if i >= n - mark_last else color
        rects.append(f"<rect x='{i*bw+1:.1f}' y='{h-bh:.1f}' width='{max(bw-2,1):.1f}' height='{bh:.1f}' fill='{c}'/>")
    return f"<svg width='{w}' height='{h}' class='spark'>{''.join(rects)}</svg>"


def hbar(value, vmax, w=110, h=12, color="#3b6ea5", label=None):
    v = 0 if value is None or (isinstance(value, float) and math.isnan(value)) else float(value)
    bw = (v / vmax * w) if vmax else 0
    t = esc(label if label is not None else fmt0(v))
    return (f"<svg width='{w+60}' height='{h+2}' class='hbar'><rect x='0' y='1' width='{w}' height='{h}' fill='#eee'/>"
            f"<rect x='0' y='1' width='{bw:.1f}' height='{h}' fill='{color}'/><text x='{w+4}' y='{h-1}' font-size='10'>{t}</text></svg>")


def stacked(months, series, w=360, h=70, colors=("#2a9d8f", "#e76f51", "#3b6ea5", "#999", "#c9b037")):
    """月ごとの構成比の積み上げ棒。series: {名前: [構成比...]}"""
    n = len(months)
    if n == 0:
        return ""
    bw = (w - 10) / n
    out = []
    for i in range(n):
        y = h - 14
        for j, (name, vals) in enumerate(series.items()):
            v = vals[i] or 0
            bh = v * (h - 16)
            y -= bh
            out.append(f"<rect x='{i*bw+5:.1f}' y='{y:.1f}' width='{max(bw-2,1):.1f}' height='{bh:.1f}' fill='{colors[j % len(colors)]}'/>")
        if i % 3 == 0 or i == n - 1:
            out.append(f"<text x='{i*bw+5:.1f}' y='{h-3}' font-size='8' fill='#555'>{months[i].strftime('%y/%m')}</text>")
    legend = "".join(f"<span class='lg'><i style='background:{colors[j % len(colors)]}'></i>{esc(k)}</span>" for j, k in enumerate(series))
    return f"<svg width='{w}' height='{h}'>{''.join(out)}</svg><div class='legend'>{legend}</div>"


def gauge(days, limit, w=120):
    """残り日数の帯：期限までの日数と、在庫が何日分かを並べる"""
    d = 0 if days is None else max(0, min(days, limit))
    return (f"<svg width='{w+4}' height='12'><rect x='0' y='1' width='{w}' height='10' fill='#eee'/>"
            f"<rect x='0' y='1' width='{d/limit*w:.1f}' height='10' fill='{'#c0392b' if days is not None and days < 90 else '#e0772b' if days is not None and days < 180 else '#2a9d8f'}'/></svg>")


# ------------------------------------------------------------------------------
# 共通の計算
# ------------------------------------------------------------------------------
def num_or(v, d=0.0):
    return d if v is None or (isinstance(v, float) and math.isnan(v)) else float(v)


def monthly_series(model, key_of):
    """月 × キー の払出（出納のある月）。戻り値：months, {key: [払出...]}"""
    t = model.T_月次[model.T_月次["出納データあり"]]
    months = sorted(set(t["月"]))
    s = defaultdict(lambda: [0.0] * len(months))
    idx = {m: i for i, m in enumerate(months)}
    for _, r in t.iterrows():
        k = key_of(r["薬品名"])
        if k is None:
            continue
        s[k][idx[r["月"]]] += num_or(r["払出数量"])
    return months, s


def holding_policy(r, settings):
    """方針書の 4 区分：回転（直近の1日平均払出 × 30 ≧ 極小需要の月の払出）× 予測できるか（ずれの大きさ）"""
    per_day = num_or(r.get("直近の1日平均払出"))
    turns = per_day * 30 >= settings.num("極小需要の月の払出")
    gs = r.get("ずれの大きさ")
    predictable = (gs is not None and not pd.isna(gs) and gs < settings.num("予測できるずれの線"))
    if per_day <= 0:
        return "持たない（払出なし）"
    if turns and predictable:
        return "自動発注で持つ"
    if turns:
        return "最低在庫を持つ"
    if predictable:
        return "持たずに取り寄せ"
    return "持たない"


# ------------------------------------------------------------------------------
# A 発注点の切り替え提案
# ------------------------------------------------------------------------------
def page_a(model, settings, abc_m):
    d = model.M_薬品
    cl = dict(zip(d["薬品名"], d["需要のまとまり"]))
    key_of = lambda n: cl.get(n) or n
    months, series = monthly_series(model, key_of)
    n = int(settings.num("ABCの月数"))
    cur_m = months[-1]
    as_map = lambda m: {r["需要のまとまり"]: (r["払出"], r["構成比"], r["累積"], r["区分"]) for _, r in abc_m[abc_m["月"] == m].iterrows()}
    cur = as_map(cur_m)
    prev_m = months[-2] if len(months) >= 2 else None
    prev = as_map(prev_m) if prev_m else {}
    rop_days, oq_days = settings.num("発注点の日数"), settings.num("発注量の日数")
    chg_line = settings.num("発注点見直しの変化")
    rows = []
    for k, (out, share, cum, cls) in cur.items():
        s = series.get(k, [])
        recent = sum(s[-n:]); before = sum(s[-2 * n:-n]) if len(s) >= 2 * n else None
        change = ((recent - before) / before) if before else None
        pcls = prev.get(k, (None, None, None, None))[3]
        drugs = d[(d["需要のまとまり"] == k) | ((d["需要のまとまり"].isna()) & (d["薬品名"] == k))]
        if drugs.empty:
            continue
        main = drugs.sort_values("直近の払出合計", ascending=False).iloc[0]
        per_day = num_or(main["直近の1日平均払出"])
        rop = math.ceil(per_day * rop_days)
        oq = math.ceil(per_day * oq_days)
        policy = holding_policy(main, settings)
        reasons = []
        if pcls and pcls != cls:
            reasons.append(f"区分 {pcls}→{cls}")
        if change is not None and abs(change) >= chg_line:
            reasons.append(f"直近{n}ヶ月の払出が前の{n}ヶ月より {change*100:+.0f}%")
        if (change is None or abs(change) < chg_line) and (not pcls or pcls == cls):
            continue
        rows.append(dict(key=k, 屋号=len(drugs), 代表=main["薬品名"], 区分=f"{pcls or '－'}→{cls}", 変更=(pcls != cls),
                         直近=recent, 前=before, 変化=change, 構成比=share, 持ち方=policy, 発注点=rop, 発注量=oq,
                         理由="、".join(reasons), spark=s))
    rows.sort(key=lambda r: (not r["変更"], -(abs(r["変化"]) if r["変化"] is not None else 0)))
    counts = {c: sum(1 for v in cur.values() if v[3] == c) for c in "ABCD"}
    changed = sum(1 for r in rows if r["変更"])
    body = [f"<div class='kpi'><div>対象月<b>{cur_m:%Y/%m}</b></div><div>区分 A／B／C／D<b>{counts['A']}／{counts['B']}／{counts['C']}／{counts['D']}</b></div>"
            f"<div>区分が変わったまとまり<b>{changed}</b></div><div>払出が ±{chg_line*100:.0f}% 以上動いた<b>{sum(1 for r in rows if r['変化'] is not None and abs(r['変化'])>=chg_line)}</b></div></div>",
            "<p class='how'>見方：区分が変わった、または直近の払出が大きく動いたまとまりだけを載せています。発注点の目安 ＝ 直近の1日平均払出 × "
            f"{rop_days:.0f} 日、発注量の目安 ＝ 直近の1日平均払出 × {oq_days:.0f} 日。持ち方は 回転（月の払出が極小需要の線以上）× 予測できるか（ずれの大きさ）で決めています。</p>",
            "<table><thead><tr><th>需要のまとまり</th><th>屋号</th><th>代表の薬（払出の多い屋号）</th><th>区分</th><th>直近%dヶ月</th><th>前の%dヶ月</th><th>変化</th><th>構成比</th>"
            "<th>12ヶ月の払出（右3本＝直近）</th><th>持ち方</th><th>発注点の目安</th><th>発注量の目安</th><th>理由</th></tr></thead><tbody>" % (n, n)]
    for r in rows[:40]:
        cls_html = f"<b class='chg'>{esc(r['区分'])}</b>" if r["変更"] else esc(r["区分"])
        ch = "" if r["変化"] is None else f"{r['変化']*100:+.0f}%"
        body.append(f"<tr><td>{esc(r['key'])}</td><td class='num'>{r['屋号']}</td><td>{esc(r['代表'])}</td><td>{cls_html}</td>"
                    f"<td class='num'>{fmt0(r['直近'])}</td><td class='num'>{fmt0(r['前'])}</td><td class='num'>{ch}</td><td class='num'>{r['構成比']*100:.1f}%</td>"
                    f"<td>{spark(r['spark'])}</td><td>{esc(r['持ち方'])}</td><td class='num'>{fmt0(r['発注点'])}</td><td class='num'>{fmt0(r['発注量'])}</td><td>{esc(r['理由'])}</td></tr>")
    body.append("</tbody></table>")
    if len(rows) > 40:
        body.append(f"<p class='note'>ほか {len(rows)-40} 件はモデル.xlsx の R1 を参照</p>")
    return "A 発注点の切り替え提案", "区分や払出が動いたまとまりの発注点を見直す", "".join(body)


# ------------------------------------------------------------------------------
# B 不動在庫と転送・返品推奨
# ------------------------------------------------------------------------------
def page_b(model, settings):
    d = model.M_薬品
    still_m, over_d = settings.num("不動とみなす月数"), settings.num("在庫過多とみなす日数")
    horizon, transfer_d = settings.num("当面の予定日数"), settings.num("転送候補の在庫日数")
    cl = dict(zip(d["薬品名"], d["需要のまとまり"]))
    # まとまりの中で今も動いている屋号（転送先・切替先の目安）
    alive = defaultdict(list)
    for _, r in d.iterrows():
        if num_or(r["直近の払出合計"]) > 0 and isinstance(r["需要のまとまり"], str):
            alive[r["需要のまとまり"]].append(r["薬品名"])
    rows = []
    for _, r in d.iterrows():
        stock = num_or(r["最終在庫"])
        if stock <= 0:
            continue
        still = r["動きのない月数"]; still = None if pd.isna(still) else int(still)
        rdays = r["直近の在庫日数"]; rdays = None if pd.isna(rdays) else float(rdays)
        sys_idle = r["システム不動日数"] if "システム不動日数" in r and not pd.isna(r["システム不動日数"]) else None
        dead = (still is not None and still >= still_m) or (r.get("不動品リスト") is True)
        over = (rdays is not None and rdays >= over_d) or (num_or(r["直近の払出合計"]) == 0)
        if not (dead or over):
            continue
        per_day = num_or(r["直近の1日平均払出"])
        surplus = max(0.0, stock - per_day * horizon)
        others = [x for x in alive.get(r["需要のまとまり"], []) if x != r["薬品名"]]
        if dead:
            act = "返品（できなければ廃棄か他店へ転送）"
        elif others and (rdays is None or rdays >= transfer_d):
            act = f"他店へ転送（同じ成分は {others[0]} が動いている）"
        elif rdays is not None and rdays >= transfer_d:
            act = "他店へ転送か返品"
        else:
            act = "様子見（発注を止める）"
        rows.append(dict(薬品名=r["薬品名"], 棚番=r["棚番"], まとまり=r["需要のまとまり"], 在庫=stock, 最後=r["最後に動いた月"], 不動=still, システム不動=sys_idle,
                         在庫日数=rdays, 余った=surplus, 処置=act, dead=dead))
    rows.sort(key=lambda x: (not x["dead"], -(x["不動"] or 0), -(x["在庫日数"] or 1e9)))
    tot_stock = sum(r["在庫"] for r in rows)
    vmax = max([r["在庫"] for r in rows] + [1])
    body = [f"<div class='kpi'><div>不動（{still_m:.0f}ヶ月以上動きなし）<b>{sum(1 for r in rows if r['dead'])} 品目</b></div><div>在庫過多（直近の在庫 {over_d:.0f} 日分以上）<b>{sum(1 for r in rows if not r['dead'])} 品目</b></div>"
            f"<div>対象の在庫数量<b>{fmt0(tot_stock)}</b></div></div>",
            f"<p class='how'>見方：余った量 ＝ 在庫 − 直近の1日平均払出 × 当面の予定 {horizon:.0f} 日。"
            f"同じ成分の別の屋号が動いていれば「転送」、動きがなければ「返品」を勧めます。処置は目安で、最終判断は現場です。</p>",
            "<table><thead><tr><th>薬品名</th><th>棚番</th><th>需要のまとまり</th><th>在庫</th><th>最後に動いた月</th><th>動きなし</th><th>不動品リスト</th><th>直近の在庫日数</th><th>余った量</th><th>処置の案</th></tr></thead><tbody>"]
    for r in rows[:45]:
        body.append(f"<tr class='{'dead' if r['dead'] else ''}'><td>{esc(r['薬品名'])}</td><td>{esc(r['棚番'])}</td><td>{esc(r['まとまり'])}</td>"
                    f"<td>{hbar(r['在庫'], vmax)}</td><td>{r['最後'].strftime('%Y/%m') if r['最後'] is not None and not pd.isna(r['最後']) else ''}</td>"
                    f"<td class='num'>{'' if r['不動'] is None else str(r['不動'])+'ヶ月'}</td><td class='num'>{'' if r['システム不動'] is None else fmt0(r['システム不動'])+'日'}</td><td class='num'>{'' if r['在庫日数'] is None else fmt0(r['在庫日数'])+'日'}</td>"
                    f"<td class='num'>{fmt0(r['余った'])}</td><td>{esc(r['処置'])}</td></tr>")
    body.append("</tbody></table>")
    if not rows:
        body.append("<p class='empty'>該当なし</p>")
    return "B 不動在庫と転送・返品推奨", "動かない在庫を返品・転送して資金と廃棄を減らす", "".join(body)


# ------------------------------------------------------------------------------
# C 経過措置・期限切迫
# ------------------------------------------------------------------------------
def page_c(model, settings):
    r3 = r3_expiry(model, settings)
    near = settings.num("期限切迫の日数")
    d = model.M_薬品
    cl = dict(zip(d["薬品名"], d["需要のまとまり"]))

    def alternative(name):
        k = cl.get(name)
        if not isinstance(k, str):
            return None
        g = d[(d["需要のまとまり"] == k) & (d["薬品名"] != name) & d["経過措置期限"].isna()]
        if g.empty:
            return None
        g = g.sort_values(["先発後発", "直近の払出合計"], ascending=[False, False])   # 後発を優先
        return g.iloc[0]["薬品名"]
    body = [f"<p class='how'>見方：経過措置の期限（経過措置リスト）までに、直近の払出ペースで在庫を使い切れるかを見ます。使い切れない分は早めに返品・転送します。"
            f"残り {near:.0f} 日以内を「切迫」としています。<br>注：薬のロットの使用期限は出納・予測データに含まれないため、ここでは見られません。"
            "在庫情報テーブルに「使用期限」列を足せば、同じ表に載せられます。</p>"]
    if r3.empty:
        body.append("<p class='empty'>経過措置リストに登録がありません</p>")
    else:
        body.append(f"<div class='kpi'><div>経過措置品<b>{len(r3)} 品目</b></div><div>期限超過<b>{int((r3['残り日数']<0).sum())}</b></div>"
                    f"<div>切迫（{near:.0f}日以内）<b>{int(((r3['残り日数']>=0)&(r3['残り日数']<=near)).sum())}</b></div><div>余る量の合計<b>{fmt0(r3['余る量'].sum())}</b></div></div>")
        body.append("<table><thead><tr><th>薬品名</th><th>経過措置期限</th><th>残り日数</th><th>残り（1年を満幅）</th><th>在庫</th><th>1日平均払出</th><th>期限までに使い切れる量</th><th>余る量</th><th>処置の案（代替の屋号つき）</th></tr></thead><tbody>")
        for _, r in r3.iterrows():
            left = int(r["残り日数"])
            alt = alternative(r["薬品名"])
            if left < 0:
                act = "期限切れ。返品か廃棄、処方の切替を確認"
            elif r["余る量"] > 0:
                act = f"余る {fmt0(r['余る量'])} を返品・転送。発注は停止"
            else:
                act = "使い切れる見込み。発注量を残り日数分に抑える"
            if alt:
                act += f"／代替は「{alt}」。在庫がなくなるまで管理を継続"
            body.append(f"<tr class='{'dead' if left < 0 else 'warn' if left <= near else ''}'><td>{esc(r['薬品名'])}</td><td>{r['経過措置期限']:%Y/%m/%d}</td>"
                        f"<td class='num'>{left}</td><td>{gauge(left, 365)}</td><td class='num'>{fmt0(r['最終在庫'])}</td><td class='num'>{r['1日平均払出']:.1f}</td>"
                        f"<td class='num'>{fmt0(r['期限までに使い切れる量'])}</td><td class='num'>{fmt0(r['余る量'])}</td><td>{esc(act)}</td></tr>")
        body.append("</tbody></table>")
    return "C 経過措置・期限切迫品", "期限までに使い切れない在庫を早めに動かす", "".join(body)


# ------------------------------------------------------------------------------
# D 出荷調整の管理票
# ------------------------------------------------------------------------------
def page_d(model, settings):
    d = model.M_薬品
    r4 = r4_shortage(model, settings)
    k6 = k6_share(model)
    months = sorted(set(k6["月"]))
    recent_months = set(months[-int(settings.num("直近の月数")):])
    flagged = set(r4.loc[r4["月"].isin(recent_months) | (r4["コード"] == "G01"), "需要のまとまり"].dropna())
    flag_rows = {(r["需要のまとまり"], r["薬品名"]): r for _, r in r4.iterrows()}
    groups = []
    base_date = settings.base_date
    for k, g in d[d["需要のまとまり"].notna()].groupby("需要のまとまり"):
        if len(g) < 2:
            continue
        # 経過措置品は推奨・次点から外し、在庫がある限り「管理継続」として載せ続ける
        trans = g[g["経過措置期限"].notna() & ((g["最終在庫"].fillna(0) > 0) | (g["直近の払出合計"].fillna(0) > 0))]
        rest = g[~g["薬品名"].isin(trans["薬品名"])]
        gen = rest[rest["先発後発"] == "後発"].sort_values("直近の払出合計", ascending=False)
        brand = rest[rest["先発後発"] != "後発"].sort_values("直近の払出合計", ascending=False)
        if len(trans):
            flagged.add(k)
        groups.append((k, g, gen, brand, trans))
    # 表示順：疑いのあるまとまり → 直近払出の多い順
    groups.sort(key=lambda x: (x[0] not in flagged, -num_or(x[1]["直近の払出合計"].sum())))
    body = ["<p class='how'>見方：同じ成分規格のまとまりごとに、推奨ジェネリック（直近の払出が最も多い後発）、次点、先発の在庫と動きを並べます。"
            "「疑い」は、構成比の急な入れ替わり（G01）か、仕入がないまま在庫が尽きかけた月（G02）があるまとまりです。推奨の在庫が薄いときは次点への切替、"
            "切替後に残った旧推奨の在庫は転送・返品を勧めます。</p>"]
    body.append("<div class='kpi'>" + f"<div>屋号が 2 つ以上のまとまり<b>{len(groups)}</b></div><div>疑いあり<b>{len(flagged)}</b></div>"
                f"<div>G01 構成比の入れ替わり<b>{int((r4['コード']=='G01').sum())}</b></div><div>G02 入荷停止の疑い<b>{int((r4['コード']=='G02').sum())}</b></div></div>")
    body.append("<table><thead><tr><th>需要のまとまり</th><th>役割</th><th>薬品名</th><th>直近%dヶ月の払出</th><th>構成比</th><th>在庫</th><th>直近の在庫日数</th><th>先月の仕入</th><th>状態</th><th>12ヶ月の構成比</th><th>切替・転送の案</th></tr></thead><tbody>" % int(settings.num("直近の月数")))
    last_m = months[-1] if months else None
    t = model.T_月次
    shown = 0
    for k, g, gen, brand, trans in groups:
        if shown >= 8 and k not in flagged:
            break
        shown += 1
        roles = []
        if len(gen):
            roles.append(("推奨ジェネリック", gen.iloc[0]))
        if len(gen) >= 2:
            roles.append(("次点ジェネリック", gen.iloc[1]))
        for _, b in brand.iterrows():
            roles.append(("先発品", b))
        for _, x in gen.iloc[2:].iterrows():
            roles.append(("その他の後発", x))
        for _, x in trans.iterrows():
            roles.append(("経過措置（管理継続）", x))
        top_name = g.sort_values("直近の払出合計", ascending=False).iloc[0]["薬品名"]
        tot = num_or(g["直近の払出合計"].sum())
        gk = k6[k6["まとまり"] == k]
        series = {}
        for _, (role, r) in enumerate(roles):
            s = gk[gk["薬品名"] == r["薬品名"]].set_index("月")["構成比"]
            series[r["薬品名"]] = [float(s.get(m, 0.0)) for m in months]
        chart = stacked(months, series, w=330, h=64)
        # 切替・転送の案
        rec, nxt = (roles[0][1] if roles and roles[0][0] == "推奨ジェネリック" else None), (roles[1][1] if len(roles) > 1 and roles[1][0] == "次点ジェネリック" else None)
        advice = []
        if rec is not None:
            rd = rec["直近の在庫日数"]
            fr = flag_rows.get((k, rec["薬品名"]))
            if fr is not None and fr["コード"] == "G02":
                advice.append(f"推奨「{rec['薬品名']}」の入荷が止まっている疑い → {('次点「'+nxt['薬品名']+'」へ切替') if nxt is not None else '先発へ切替'}")
            elif rd is not None and not pd.isna(rd) and rd < settings.num("発注点の日数"):
                advice.append(f"推奨の在庫が {rd:.0f} 日分と薄い → 次点を併用")
        for role, r in roles:
            fr = flag_rows.get((k, r["薬品名"]))
            if fr is not None and fr["コード"] == "G01" and num_or(fr["変化"]) < 0 and num_or(r["最終在庫"]) > 0:
                advice.append(f"「{r['薬品名']}」は {fr['月']:%Y/%m} に構成比 {fr['前月の構成比']*100:.0f}%→{fr['当月の構成比']*100:.0f}%。残った在庫 {fmt0(r['最終在庫'])} は転送・返品を検討")
        for _, x in trans.iterrows():
            left = (x["経過措置期限"] - base_date).days
            stock = num_or(x["最終在庫"]); per_day = num_or(x["直近の1日平均払出"])
            usable = per_day * max(0, left)
            new_rec = rec["薬品名"] if rec is not None else (brand.iloc[0]["薬品名"] if len(brand) else None)
            if x["薬品名"] == top_name:
                advice.append(f"【警告】経過措置「{x['薬品名']}」がまだ一番出ています（期限 {x['経過措置期限']:%Y/%m/%d}、残り {left} 日）。"
                              + (f"推奨を「{new_rec}」に置き換え、処方の切替を医師に依頼" if new_rec else "代替の屋号がありません。登録を"))
            elif stock > 0:
                advice.append(f"経過措置「{x['薬品名']}」の在庫 {fmt0(stock)} は "
                              + (f"期限まで使い切れます（{fmt0(usable)}）" if usable >= stock else f"{fmt0(stock - usable)} が余る見込み。返品・転送を") + "。管理を継続")
            else:
                advice.append(f"経過措置「{x['薬品名']}」は在庫 0。期限 {x['経過措置期限']:%Y/%m/%d} まで管理を継続")
        if not advice:
            advice.append("切替の必要なし")
        first = True
        for role, r in roles:
            fr = flag_rows.get((k, r["薬品名"]))
            state = "" if fr is None else ("構成比の入れ替わり（%s）" % fr["月"].strftime("%Y/%m") if fr["コード"] == "G01" else "入荷停止の疑い（%s）" % fr["月"].strftime("%Y/%m"))
            if role == "経過措置（管理継続）":
                state = f"経過措置 期限 {r['経過措置期限']:%Y/%m/%d}（残り {(r['経過措置期限'] - base_date).days} 日）"
            lm = t[(t["薬品名"] == r["薬品名"]) & (t["月"] == last_m)]
            buy = fmt0(lm["仕入数量"].iloc[0]) if len(lm) else ""
            sh = (num_or(r["直近の払出合計"]) / tot) if tot else 0
            rd = r["直近の在庫日数"]
            body.append(f"<tr class='{'grp' if first else ''} {'warn' if (fr is not None or role.startswith('経過措置')) else ''}'>"
                        + (f"<td rowspan='{len(roles)}' class='top'>{esc(k)}{'<br><span class=tag>経過措置</span>' if len(trans) else ('<br><span class=tag>疑い</span>' if k in flagged else '')}</td>" if first else "")
                        + f"<td>{esc(role)}</td><td>{esc(r['薬品名'])}</td><td class='num'>{fmt0(r['直近の払出合計'])}</td><td class='num'>{sh*100:.0f}%</td>"
                        f"<td class='num'>{fmt0(r['最終在庫'])}</td><td class='num'>{'' if rd is None or pd.isna(rd) else fmt0(rd)+'日'}</td><td class='num'>{buy}</td><td>{esc(state)}</td>"
                        + (f"<td rowspan='{len(roles)}' class='top'>{chart}</td><td rowspan='{len(roles)}' class='top advice'>{'<br>'.join(esc(a) for a in advice)}</td>" if first else "")
                        + "</tr>")
            first = False
    body.append("</tbody></table>")
    body.append(f"<p class='note'>疑いのあるまとまりを先頭に、ほかは直近の払出が多い順に {shown} 件まで。全件はモデル.xlsx の K6 を参照</p>")
    return "D 出荷調整の管理票", "推奨ジェネリックの入荷状況を見て、切替と在庫の転送を決める", "".join(body)


# ------------------------------------------------------------------------------
# E 極小需要薬のビュー
# ------------------------------------------------------------------------------
def page_e(model, settings):
    d = model.M_薬品
    tiny_out = settings.num("極小需要の月の払出")
    rows = []
    for _, r in d.iterrows():
        recent = num_or(r["直近の払出合計"])
        if r["出どころ"] == "出納だけ" and recent == 0 and num_or(r["最終在庫"]) == 0:
            continue
        if recent > tiny_out * int(settings.num("直近の月数")):      # 直近 N ヶ月の払出が 極小需要の月の払出 × N 以下
            continue
        typ = r["需要の型"] if isinstance(r["需要の型"], str) else "予定なし"
        stock = num_or(r["最終在庫"]); rd = r["直近の在庫日数"]
        excl = bool(r["発注除外"])
        if excl:
            act = "発注除外（取り寄せ）として運用中"
        elif recent == 0 and stock > 0:
            act = "払出なし。在庫を返品・転送し、発注除外に登録"
        elif recent == 0:
            act = "払出も在庫もなし。発注除外に登録"
        elif stock <= 0:
            act = "在庫なし。患者ごとに取り寄せ（発注除外に登録）"
        elif rd is not None and not pd.isna(rd) and rd >= settings.num("在庫過多とみなす日数"):
            act = "最小限だけ持つ。余りは転送"
        else:
            act = "最小限の在庫で運用"
        rows.append(dict(薬品名=r["薬品名"], 型=typ, 間隔=r["平均間隔"], 患者=num_or(r["予定患者数合計"]), 直近=recent, 在庫=stock, 在庫日数=rd,
                         最後=r["最後に動いた月"], 動いた月=r["動いた月数"], 除外=excl, 処置=act))
    rows.sort(key=lambda x: (x["除外"], -(x["在庫日数"] if x["在庫日数"] is not None and not pd.isna(x["在庫日数"]) else 1e9), -x["在庫"]))
    body = [f"<div class='kpi'><div>極小需要（直近 {int(settings.num('直近の月数'))} ヶ月の払出が月 {fmt0(tiny_out)} 以下）<b>{len(rows)} 品目</b></div>"
            f"<div>うち在庫あり<b>{sum(1 for r in rows if r['在庫']>0)}</b></div><div>その在庫数量<b>{fmt0(sum(r['在庫'] for r in rows))}</b></div>"
            f"<div>発注除外に登録済み<b>{sum(1 for r in rows if r['除外'])}</b></div></div>",
            "<p class='how'>見方：ほとんど出ない薬は、最小限だけ持つか、持たずに患者ごとに取り寄せるかを決めます。在庫が直近の払出に対して多すぎるものは転送・返品の候補です。"
            "平均間隔は「処方のあった日数 ÷ その薬の予定があった日数」で、大きいほどまれです。</p>",
            "<table><thead><tr><th>薬品名</th><th>需要の型</th><th>平均間隔（日）</th><th>年間の予定患者数</th><th>直近の払出</th><th>在庫</th><th>直近の在庫日数</th><th>最後に動いた月</th><th>動いた月数</th><th>発注除外</th><th>処置の案</th></tr></thead><tbody>"]
    for r in rows[:60]:
        rd = r["在庫日数"]
        body.append(f"<tr class='{'warn' if (rd is not None and not pd.isna(rd) and rd >= settings.num('在庫過多とみなす日数')) or (r['直近']==0 and r['在庫']>0) else ''}'>"
                    f"<td>{esc(r['薬品名'])}</td><td>{esc(r['型'])}</td><td class='num'>{'' if r['間隔'] is None or pd.isna(r['間隔']) else f'{r['間隔']:.1f}'}</td>"
                    f"<td class='num'>{fmt0(r['患者'])}</td><td class='num'>{fmt0(r['直近'])}</td><td class='num'>{fmt0(r['在庫'])}</td>"
                    f"<td class='num'>{'' if rd is None or pd.isna(rd) else fmt0(rd)+'日'}</td><td>{r['最後'].strftime('%Y/%m') if r['最後'] is not None and not pd.isna(r['最後']) else ''}</td>"
                    f"<td class='num'>{'' if r['動いた月'] is None or pd.isna(r['動いた月']) else int(r['動いた月'])}</td><td>{'○' if r['除外'] else ''}</td><td>{esc(r['処置'])}</td></tr>")
    body.append("</tbody></table>")
    if len(rows) > 60:
        body.append(f"<p class='note'>ほか {len(rows)-60} 件はモデル.xlsx の M_薬品 を参照</p>")
    return "E 要管理の極小需要薬", "持つか取り寄せるかを決め、余分な在庫を減らす", "".join(body)


# ------------------------------------------------------------------------------
# 組み立て
# ------------------------------------------------------------------------------
CSS = """
@page { size: A4 landscape; margin: 10mm 10mm 12mm 10mm; }
body { font-family: 'Yu Gothic UI', 'Meiryo', sans-serif; color: #222; margin: 0; font-size: 10.5px; }
.page { page-break-after: always; padding: 6px 8px; }
.page:last-child { page-break-after: auto; }
.hdr { display: flex; justify-content: space-between; align-items: baseline; border-bottom: 3px solid #3b6ea5; margin-bottom: 6px; }
.hdr h1 { font-size: 17px; margin: 0; } .hdr .goal { color: #3b6ea5; font-weight: bold; font-size: 12px; }
.hdr .meta { color: #666; font-size: 10px; }
.kpi { display: flex; gap: 8px; margin: 6px 0; } .kpi div { border: 1px solid #cfd8e3; background: #f4f7fb; padding: 4px 10px; min-width: 120px; font-size: 10px; }
.kpi b { display: block; font-size: 16px; }
.how { color: #444; font-size: 10px; margin: 4px 0 6px 0; line-height: 1.4; }
table { border-collapse: collapse; width: 100%; } th, td { border: 1px solid #c9c9c9; padding: 2px 5px; vertical-align: middle; }
th { background: #e9eef5; font-size: 10px; white-space: nowrap; } td { font-size: 10.5px; } td.num { text-align: right; white-space: nowrap; }
td.top { vertical-align: top; } td.advice { max-width: 220px; white-space: normal; }
tr.dead td { background: #fdecea; } tr.warn td { background: #fff4e5; } tr.grp td { border-top: 2px solid #888; }
b.chg { color: #c0392b; } .tag { background: #c0392b; color: #fff; padding: 0 4px; font-size: 9px; border-radius: 3px; }
.note, .empty { color: #666; font-size: 10px; } .legend { font-size: 9px; color: #444; } .lg { margin-right: 8px; white-space: nowrap; }
.lg i { display: inline-block; width: 9px; height: 9px; margin-right: 3px; vertical-align: middle; }
svg.spark, svg.hbar { vertical-align: middle; }
@media screen { body { background: #ddd; } .page { background: #fff; width: 277mm; min-height: 190mm; margin: 10px auto; box-shadow: 0 0 6px #999; } }
"""


def page_f(model, settings):
    """F 発注書（来週の分）：㉔ R_発注書 と同じ表"""
    df = r5_order(model, settings)
    n_days = int(settings.num("予測日数"))
    cnt = {k: int((df["判定"] == k).sum()) for k in ("発注する", "控えめに", "発注しない", "不要")} if len(df) else {}
    body = [f"<div class='kpi'><div>来週<b>{settings.fc_start:%m/%d}～{settings.fc_end:%m/%d}（{n_days} 日）</b></div>"
            f"<div>発注する<b>{cnt.get('発注する', 0)}</b></div><div>控えめに<b>{cnt.get('控えめに', 0)}</b></div>"
            f"<div>発注しない<b>{cnt.get('発注しない', 0)}</b></div><div>載せた薬<b>{len(df)}</b></div></div>",
            "<p class='how'>見方：来週の予定（予測データ）と在庫（発注残込み）の差が 発注の目安（錠）。1 年の実績とのずれで補正しています。"
            "「発注しない」は不足していても発注しない薬（発注除外・供給停止）。「控えめに」は目安より少なめに。"
            "必要な最大在庫は、納品が止まる最長の期間（連休）に出る量で、区分 A の薬の在庫の上限の目安です。</p>",
            "<table><thead><tr><th>区分</th><th>略称</th><th>薬品名</th><th>来週の予定</th><th>患者</th><th>在庫</th><th>不足</th><th>ずれ</th>"
            "<th>見込み</th><th>発注の目安</th><th>必要な最大在庫</th><th>最大の期間</th><th>判定</th><th>注意</th></tr></thead><tbody>"]
    for _, r in df.head(70).iterrows():
        cls = "warn" if r["判定"] == "発注しない" or "上限を超える" in (r["注意"] or "") else ""
        body.append(f"<tr class='{cls}'><td>{esc(r['区分'] if isinstance(r['区分'], str) else '')}</td><td>{esc(r['略称'] if isinstance(r['略称'], str) else '')}</td>"
                    f"<td>{esc(r['薬品名'])}</td><td class='num'>{fmt0(r['来週の予定'])}</td><td class='num'>{fmt0(r['来週の患者'])}</td>"
                    f"<td class='num'>{fmt0(r['在庫'])}</td><td class='num'>{fmt0(r['不足'])}</td><td class='num'>{r['ずれ']*100:+.0f}%</td>"
                    f"<td class='num'>{fmt0(r['見込み'])}</td><td class='num'><b>{fmt0(r['発注の目安'])}</b></td><td class='num'>{fmt0(r['必要な最大在庫'])}</td>"
                    f"<td>{esc(r['最大の期間'])}</td><td>{esc(r['判定'])}</td><td>{esc(r['注意'])}</td></tr>")
    body.append("</tbody></table>")
    if len(df) > 70:
        body.append(f"<p class='note'>ほか {len(df)-70} 件はモデル.xlsx の R5_発注書 を参照</p>")
    return "F 発注書（来週の分）", "来週の予定と在庫から、何をどのくらい発注するかを決める", "".join(body)


def build(model, settings, out_html):
    abc_m = abc_monthly(model, settings)
    pages = [page_a(model, settings, abc_m), page_b(model, settings), page_c(model, settings), page_d(model, settings), page_e(model, settings),
             page_f(model, settings)]
    months = sorted(set(model.T_月次.loc[model.T_月次["出納データあり"], "月"]))
    meta = f"対象月 {months[-1]:%Y/%m}　期間 {settings.start:%Y/%m/%d}～{settings.end:%Y/%m/%d}　基準日 {settings.base_date:%Y/%m/%d}　作成 {dt.datetime.now():%Y/%m/%d}　※架空データ"
    parts = []
    for i, (title, goal, body) in enumerate(pages, 1):
        parts.append(f"<section class='page'><div class='hdr'><h1>{esc(title)}</h1><span class='goal'>{esc(goal)}</span><span class='meta'>{esc(meta)}　{i}/{len(pages)}</span></div>{body}</section>")
    doc = f"<!DOCTYPE html><html lang='ja'><head><meta charset='utf-8'><title>在庫監視レポート</title><style>{CSS}</style></head><body>{''.join(parts)}</body></html>"
    with open(out_html, "w", encoding="utf-8") as f:
        f.write(doc)


def to_pdf(html_path, pdf_path):
    if not os.path.exists(EDGE):
        print("Edge が見つからないので PDF は作らない"); return
    url = "file:///" + html_path.replace("\\", "/")
    subprocess.run([EDGE, "--headless", "--disable-gpu", "--no-pdf-header-footer", f"--print-to-pdf={pdf_path}", url],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=120)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tables", required=True)
    ap.add_argument("--out", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "out"))
    ap.add_argument("--pdf", action="store_true")
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    settings = Settings(a.tables)
    model = Model(settings)
    hp = os.path.join(a.out, "監視レポート.html")
    build(model, settings, hp)
    print("出力：", hp)
    if a.pdf:
        pp = os.path.join(a.out, "監視レポート.pdf")
        to_pdf(hp, pp); print("PDF：", pp)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    main()
