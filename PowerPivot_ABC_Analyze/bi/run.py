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
# 区分（0.5.0 から model.py の abc_table で計算。ここは互換のための入口）
# ------------------------------------------------------------------------------
def abc_monthly(model, settings):
    """T_ABC月次：月 × まとまり の区分（model.T_ABC月次 と同じ）"""
    return model.T_ABC月次


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
                         直近の払出=now["払出"], 構成比=now["構成比"], 累積=now["累積"], 月平均払出=now["月平均払出"], 量の区分=now["量の区分"],
                         患者の最少月=now["患者の最少月"], 今月の区分=now["区分"], 前月の区分=before,
                         変更=("" if before in (None, now["区分"]) else f"{before}→{now['区分']}"),
                         在庫日数=r["在庫日数"], 需要の型=r["需要の型"], 最終在庫=r["最終在庫"],
                         傾向=r.get("まとまりの傾向"), 年間比=r.get("まとまりの年間比"), 患者の傾向=r.get("患者の傾向")))
    out = pd.DataFrame(rows)
    return out.sort_values(["累積", "需要のまとまり", "薬品名"]).reset_index(drop=True)


def r2_dead(model, settings):
    d, rs = model.M_薬品, model.T_判定理由
    horizon = settings.num("当面の予定日数")
    s_codes = rs[rs["コード"].isin(["S01", "S02", "S05", "S06"])]
    if s_codes.empty:
        return pd.DataFrame()
    msgs = s_codes.groupby("薬品名")["文言"].apply(lambda x: "／".join(x))
    rows = []
    for _, r in d[d["薬品名"].isin(msgs.index)].iterrows():
        per_day = r["直近の1日平均払出"] if not pd.isna(r["直近の1日平均払出"]) else 0.0
        stock = r["最終在庫"] if not pd.isna(r["最終在庫"]) else 0.0
        surplus = max(0.0, stock - per_day * horizon)
        rows.append(dict(薬品名=r["薬品名"], 略称=r["略称"], 棚番=r["棚番"], 最終在庫=stock, 最後に動いた月=r["最後に動いた月"],
                         動きのない月数=r["動きのない月数"], 在庫日数=r["在庫日数"], 直近の在庫日数=r["直近の在庫日数"],
                         **{"直近の1日平均払出": per_day}, 余った量=surplus,
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
    # G02：仕入 0 が 2 ヶ月以上続き、払出があり、在庫が月の払出を下回った（ほとんど出ない薬は除く：箱単位の発注なら仕入 0 の月は普通）
    tiny_out = settings.num("極小需要の月の払出")
    for n, g in tt.sort_values("月").groupby("薬品名"):
        g = g.reset_index(drop=True)
        for i in range(1, len(g)):
            a, b = g.loc[i - 1], g.loc[i]
            if (a["仕入数量"] or 0) == 0 and (b["仕入数量"] or 0) == 0 and (b["払出数量"] or 0) >= tiny_out and (a["払出数量"] or 0) > 0 \
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


def r5_order(model, settings):
    """R5 発注書（㉔ R_発注書 と同じ定義）：来週の予定と在庫から発注の目安・判定・注意。区分 → 略称 → 薬品名
       0.9.0：同じ需要のまとまり（屋号が 2 つ以上）は予定・在庫を合算し、発注先（M_薬品）の 1 薬に寄せる
       0.12.0：発注の目安は区分ごと（A 上限まで／B 倍率×見込み／C 必要量／D 人が判断）。目安か不足のある薬だけ載せる"""
    import bisect
    from model import calendar
    S = settings
    gap_r, near, b_mult = S.num("ずれの注意割合"), S.num("期限切迫の日数"), S.num("区分Bの倍率")
    n_txt = str(int(S.num("直近の月数")))
    pct = lambda x: "－" if (x is None or (isinstance(x, float) and math.isnan(x))) else f"{fmt0(x * 100)}%"
    up = lambda x: math.ceil(round(max(0.0, x), 6))
    # 来週の予定・患者・在庫（一番新しい日の発注残込み）
    fcg = {}
    fc = model.予測データ
    if len(fc):
        for n, g in fc.groupby("薬品名"):
            last = g.sort_values("日付").iloc[-1]
            st = last["在庫量_発注残込み"] if not pd.isna(last["在庫量_発注残込み"]) else last["在庫量"]
            fcg[n] = (float(g["数量"].fillna(0).sum()), float(g["患者数"].fillna(0).sum()), None if pd.isna(st) else float(st))
    md = model.M_薬品
    g_ = lambda r, k: (None if (k not in r.index or pd.isna(r[k])) else r[k])
    rows_by_name = {r["薬品名"]: r for _, r in md.iterrows()}
    key_of = {r["薬品名"]: (r["需要のまとまり"] if isinstance(g_(r, "発注先"), str) else r["薬品名"]) for _, r in md.iterrows()}
    abbr_of = {r["薬品名"]: g_(r, "略称") for _, r in md.iterrows()}
    stock_of, plan_of, pats_of = {}, {}, {}
    for _, r in md.iterrows():
        n = r["薬品名"]
        plan, pats, st = fcg.get(n, (0.0, 0.0, None))
        stock_of[n] = st if st is not None else (g_(r, "最終在庫") or 0.0)
        plan_of[n], pats_of[n] = plan, pats
    g_plan, g_stock, g_pats = defaultdict(float), defaultdict(float), defaultdict(float)
    for n, k in key_of.items():
        g_plan[k] += plan_of[n]; g_stock[k] += stock_of[n]; g_pats[k] += pats_of[n]
    # 納品と納品の間（翌営業日納品）に出る予定の最大（まとまりは合算）
    cal = calendar(S.start, S.end + dt.timedelta(days=14), S)
    biz = sorted(cal.loc[cal["卸の営業日"], "日付"])
    def next_biz(d):
        i = bisect.bisect_right(biz, d)
        return biz[i] if i < len(biz) else None
    mx = {}
    day = model.T_日次
    if len(day):
        day = day.assign(納品日=day["日付"].map(next_biz), キー=day["薬品名"].map(lambda n: key_of.get(n, n)))
        blk = day.groupby(["キー", "納品日"]).agg(量=("予定数量", lambda x: float(x.fillna(0).sum())), 始=("日付", "min"), 終=("日付", "max")).reset_index()
        for k, g in blk.groupby("キー"):
            r = g.sort_values(["量", "始"], ascending=[False, True]).iloc[0]
            mx[k] = (float(r["量"]), r["始"], r["終"])
    adj = getattr(model, "T_まとまり月次", None)
    adj_names = set()
    if adj is not None and len(adj) and "G01" in adj.columns:
        a = adj[adj["月"] == adj["月"].max()]
        adj_names = set(a.loc[(a["G01"] == True) | (a["G02"] == True), "薬品名"])
    lots = model.T_ロット
    lot_names = set()
    if len(lots) and "残り日数" in lots.columns:
        lot_names = set(lots.loc[((lots["残り日数"].notna()) & (lots["残り日数"] <= near)) | (lots["返す量"].fillna(0) > 0), "薬品名"])
    order_of = {"A": 1, "B": 2, "C": 3, "D": 4}
    rows = []
    for _, r in md.iterrows():
        n = r["薬品名"]
        g = lambda k: g_(r, k)
        tgt = g("発注先")
        pooled = isinstance(tgt, str)
        target = (not pooled) or tgt == n
        t = rows_by_name.get(tgt, r) if pooled else r           # 発注の考え方・判定を決める薬（まとまりなら発注先）
        tg = lambda k: g_(t, k)
        k = key_of[n]
        plan, pats, stock = plan_of[n], pats_of[n], stock_of[n]
        gp = g_plan[k] if pooled else plan
        gs = g_stock[k] if pooled else stock
        gap = (g("まとまりのずれの割合") if pooled else g("ずれの割合")) or 0.0
        factor = max(0.0, 1 + gap)
        exp = gp * factor
        short = max(0.0, gp - gs)
        mq = math.ceil(round((mx[k][0] if k in mx else 0.0) * factor, 6))
        cls, sup = g("区分"), g("供給状況")
        tcls = tg("区分")
        excl = bool(r["発注除外"]); trans = bool(r["経過措置"])
        sw_to, sw_from = g("切替先"), g("切替元")
        no_order = bool(t["発注除外"]) or tg("供給状況") in ("供給停止", "薬価削除予定") or bool(tg("切替先"))
        careful = bool(t["経過措置"]) or tg("供給状況") == "限定" or gap <= -gap_r or (bool(sw_from) and not pooled)
        need = up(exp - gs)
        if tcls == "A" and mq > 0:
            policy, oq_p = "上限まで", up(mq - gs)
        elif tcls == "B":
            policy, oq_p = f"{b_mult:g}倍を在庫", up(exp * b_mult - gs)
        elif tcls == "D":
            policy, oq_p = "人が判断", need
        else:
            policy, oq_p = "必要量", need
        oq_g = 0 if no_order else oq_p
        oq = oq_g if target else 0
        over = target and tcls == "A" and mq > 0 and gs + oq > mq
        grp_txt = f"予定 {fmt0(g_plan[k])}・在庫 {fmt0(g_stock[k])} を合算"
        notes = []
        if excl: notes.append("発注除外")
        if sup in ("供給停止", "薬価削除予定"): notes.append(sup)
        if sup == "限定": notes.append("限定出荷")
        if n in adj_names: notes.append("出荷調整の疑い")
        if n in lot_names: notes.append("期限の近いロットを先に使う")
        if gap <= -gap_r: notes.append(f"予定より {fmt0(-gap * 100)}% 少なく出る")
        if cls == "D": notes.append("患者が少ない（区分 D）")
        if trans: notes.append("経過措置品（切替先があれば発注しない）")
        if r["出どころ"] == "出納だけ": notes.append("予定にない薬")
        if sw_to: notes.append("切替元（在庫を使い切る）" if pooled else f"切替元（「{sw_to}」へ寄せる。在庫を使い切る）")
        if sw_from: notes.append("切替先（切替元の在庫は合算済み）" if pooled else f"切替先（切替元の在庫 {fmt0(g('切替元の在庫'))} を使い切るまで控えめに）")
        if over: notes.append(f"上限を超える（在庫＋目安 {fmt0(gs + oq)} ＞ {fmt0(mq)}）")
        trend = g("まとまりの傾向") if pooled else g("傾向")
        tr1 = g("まとまりの前期比") if pooled else g("前期比")
        tr2 = g("まとまりの年間比") if pooled else g("年間比")
        trend_notes = ([f"{trend}傾向（直近 {n_txt} ヶ月は前の {n_txt} ヶ月の {pct(tr1)}、年間の {pct(tr2)}。患者は {g('患者の傾向') or '不明'}）。"
                        f"目安を{'増やす' if trend == '増加' else '減らす'}か検討"] if trend in ("増加", "減少") else [])
        pool_notes = [] if not pooled else (
            [f"まとまり「{g('需要のまとまり')}」の発注先（{int(g('まとまりの品目数'))} 品目の {grp_txt}）"] if target
            else [f"「{tgt}」に寄せる（まとまり「{g('需要のまとまり')}」：{grp_txt}）"])
        note = "、".join(pool_notes + trend_notes + notes)
        verdict = ("寄せる" if (pooled and not target) else "発注しない" if no_order
                   else "要判断" if (oq > 0 and tcls == "D") else "控えめに" if (oq > 0 and careful) else "発注する" if oq > 0 else "不要")
        if not (oq_g > 0 or short > 0):
            continue
        period = "" if (k not in mx or not target) else f"{mx[k][1]:%Y/%m/%d}～{mx[k][2]:%Y/%m/%d}（{(mx[k][2] - mx[k][1]).days + 1}日）"
        rows.append(dict(区分=cls, 発注の考え方=policy, 略称=g("略称"), 薬品名=n, 来週の予定=plan, 来週の患者=pats, 在庫=stock,
                         需要のまとまり=(g("需要のまとまり") if pooled else None),
                         まとまりの予定=(g_plan[k] if pooled else None), まとまりの在庫=(g_stock[k] if pooled else None),
                         不足=(short if target else None), ずれ=gap, 傾向=(trend if trend in ("増加", "減少") else None),
                         見込み=(float(round(exp)) if target else None), 発注の目安=oq,
                         必要な最大在庫=(mq if target else None), 最大の期間=period, 判定=verdict, 注意=note,
                         並び=order_of.get(cls, 5), 並び略=((abbr_of.get(tgt) if pooled else g("略称")) or ""), キー=k, 並び先=(0 if target else 1)))
    cols = ["区分", "発注の考え方", "略称", "薬品名", "来週の予定", "来週の患者", "在庫", "需要のまとまり", "まとまりの予定", "まとまりの在庫",
            "不足", "ずれ", "傾向", "見込み", "発注の目安", "必要な最大在庫", "最大の期間", "判定", "注意"]
    if not rows:
        return pd.DataFrame(columns=cols)
    df = pd.DataFrame(rows).sort_values(["並び", "並び略", "キー", "並び先", "薬品名"]).reset_index(drop=True)
    return df[cols]


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
                         A=int((a == "A").sum()), B=int((a == "B").sum()), C=int((a == "C").sum()), D=int((a == "D").sum())))
    return pd.DataFrame(rows)


def k1_stock_bins(model):
    d = model.M_薬品.copy()
    d["在庫日数の帯"] = d["直近の在庫日数"].map(stock_bin)       # 直近 N ヶ月の払出で見た在庫日数
    order = ["払出なし"] + [b[2] for b in STOCK_BINS]
    g = d.groupby("在庫日数の帯").agg(薬品数=("薬品名", "count"), 在庫数量=("最終在庫", "sum")).reindex(order).fillna(0).reset_index()
    return g


def k2_quadrant(model, settings):
    d = model.M_薬品.copy()
    tiny_out = settings.num("極小需要の月の払出")
    d["回転"] = d["直近の1日平均払出"].map(lambda x: "払出なし" if pd.isna(x) or x <= 0 else ("回る" if x * 30 >= tiny_out else "回らない"))
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
    body += ["<h2>K2 4 象限（需要の型 × 回転。回転は直近の1日平均払出 × 30 が「極小需要の月の払出」以上）：薬品数</h2>", df_html(k2n.reset_index()),
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
    R = dict(R1=r1_abc(model, settings, abc_m), R2=r2_dead(model, settings), R3=r3_expiry(model, settings), R4=r4_shortage(model, settings),
             R5=r5_order(model, settings))
    # 薬品グループ案：出荷調整の疑い（R4 の最新月の G01・G02）も載せる理由に入れる
    from model import propose_groups
    r4 = R["R4"]
    adj = set(r4.loc[r4["月"] == r4["月"].max(), "薬品名"]) if len(r4) else set()
    model.薬品グループ案 = propose_groups(model.M_薬品, adj)
    K = dict(K0=k0_summary(model, settings, abc_m), K1=k1_stock_bins(model), K2=k2_quadrant(model, settings), K4=k4_loss(model), K6=k6_share(model))
    sheets = [("T_日次", model.T_日次), ("T_月次", model.T_月次), ("M_日付", model.M_日付), ("M_薬品", model.M_薬品),
              ("M_理由コード", model.M_理由コード), ("T_判定理由", model.T_判定理由), ("T_ロット", model.T_ロット), ("T_ABC月次", abc_m),
              ("薬品グループ案", model.薬品グループ案),
              ("R1_ABC区分表", R["R1"]), ("R2_不動薬の処理候補", R["R2"]), ("R3_経過措置・期限切迫", R["R3"]), ("R4_出荷調整の可能性", R["R4"]), ("R5_発注書", R["R5"]),
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
