"""EDA：需要を「量（月平均払出）」と「広がり（月平均の延べ患者数）」の 2 軸で見て、対数の需要スコアと「偏った需要」を試す
使い方（プロジェクトのフォルダで）：python bi/eda_demand.py [--model bi/out/モデル.xlsx] [--out docs]
  入力：bi/run.py が書いた モデル.xlsx の T_月次（払出数量＝出納、予定患者数＝予定データの延べ患者数）
  出力：docs/EDA_需要スコア_散布.png と標準出力（スコアの分布、区分の数、偏った需要の例）
  定義（案）
    量 q  ＝ 出納のある月の払出の月平均（錠／月）
    広がり p ＝ 予定データのある月の延べ患者数の月平均（人／月。定期患者だけ。来局 1 回を 1 と数える）
    需要スコア S ＝ log10(q) ＋ log10(p)（＝ log10(q×p)。量も患者も多いほど高い。患者が少ないと量が多くても下がる）
    1 人あたり   ＝ q ÷ p（錠／人・月。大きいほど少数の患者に偏った需要）
"""
import os, sys, math, argparse
import numpy as np
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--model", default=os.path.join(ROOT, "bi", "out", "モデル.xlsx"))
    ap.add_argument("--out", default=os.path.join(ROOT, "docs"))
    ap.add_argument("--p-line", type=float, default=5.0, help="偏った需要とみなす延べ患者数／月の線")
    a = ap.parse_args()

    x = pd.ExcelFile(a.model)
    t = pd.read_excel(x, "T_月次")
    sui = t[t["出納データあり"] == True]
    pln = t[t["予定データあり"] == True]
    q = sui.groupby("薬品名")["払出数量"].sum().fillna(0) / sui["月"].nunique()
    p = pln.groupby("薬品名")["予定患者数"].sum().fillna(0) / pln["月"].nunique()
    d = pd.DataFrame({"q": q}).join(p.rename("p"), how="left").fillna({"p": 0})
    d["per"] = np.where(d["p"] > 0, d["q"] / d["p"], np.nan)
    ok = (d["q"] > 0) & (d["p"] > 0)
    d["lq"] = np.where(d["q"] > 0, np.log10(d["q"].clip(lower=1e-9)), np.nan)
    d["lp"] = np.where(d["p"] > 0, np.log10(d["p"].clip(lower=1e-9)), np.nan)
    d["S"] = d["lq"] + d["lp"]
    print(f"薬品 {len(d)} 品目：払出あり {int((d['q']>0).sum())}、患者あり {int((d['p']>0).sum())}、両方あり {int(ok.sum())}、"
          f"払出あり患者なし（予定に出ない薬） {int(((d['q']>0)&(d['p']<=0)).sum())}")
    print(f"延べ患者数／月：中央値 {d.loc[ok,'p'].median():.1f}、10% 点 {d.loc[ok,'p'].quantile(.1):.1f}、90% 点 {d.loc[ok,'p'].quantile(.9):.1f}")
    print(f"1 人あたりの月量（q/p）：中央値 {d.loc[ok,'per'].median():.0f} 錠、90% 点 {d.loc[ok,'per'].quantile(.9):.0f} 錠")
    r = np.corrcoef(d.loc[ok, "lq"], d.loc[ok, "lp"])[0, 1]
    print(f"log q と log p の相関：{r:.2f}")

    # ---- スコアの分布と 4 区分（平均±σ） ----
    s = d.loc[ok, "S"]
    mu, sd = float(s.mean()), float(s.std(ddof=0))
    print(f"需要スコア S = log10(q×p)：平均 {mu:.2f}、σ {sd:.2f}、最小 {s.min():.2f}、最大 {s.max():.2f}")
    def cls(v):
        if pd.isna(v):
            return "D"
        z = (v - mu) / sd
        return "A" if z > 1 else "B" if z > 0 else "C" if z > -1 else "D"
    d["区分S"] = d["S"].map(cls)
    # 比較：量だけ（log q の平均±σ）
    lq = d.loc[d["q"] > 0, "lq"]; mq, sq = float(lq.mean()), float(lq.std(ddof=0))
    d["区分q"] = d["lq"].map(lambda v: "D" if pd.isna(v) else ("A" if (v-mq)/sq > 1 else "B" if (v-mq)/sq > 0 else "C" if (v-mq)/sq > -1 else "D"))
    print("\n■ 区分ごとの薬品数：スコア S（量×患者）と 量だけ の比較")
    ct = pd.crosstab(d["区分q"], d["区分S"], rownames=["量だけ"], colnames=["スコアS"])
    print(ct.to_string())
    moved = d[d["区分q"] != d["区分S"]]
    print(f"区分が変わる薬：{len(moved)} 品目（量だけ→スコア）。うち下がる {int((moved['区分q'] < moved['区分S']).sum())}、上がる {int((moved['区分q'] > moved['区分S']).sum())}")

    # ---- 偏った需要：量はあるのに患者が少ない ----
    d["偏り"] = (d["q"] > 0) & (d["p"] < a.p_line)
    bias = d[d["偏り"] & (d["q"] >= 10 ** (mq - sq))].sort_values("per", ascending=False)
    print(f"\n■ 偏った需要（延べ患者 {a.p_line:.0f} 人／月未満で、量が『量だけ』の D 線以上）：{len(bias)} 品目。1 人あたりの多い順に 10 件")
    show = bias.head(10).copy()
    show["q"] = show["q"].round(0); show["p"] = show["p"].round(1); show["per"] = show["per"].round(0)
    print(show[["q", "p", "per", "区分q", "区分S"]].rename(columns={"q": "払出/月", "p": "患者/月", "per": "1人あたり"}).to_string())
    print("\n■ 量が多いのに患者で見ると下がる例（区分q が A/B、区分S が 1 つ以上下）")
    ex = d[(d["区分q"].isin(["A", "B"])) & (d["区分S"] > d["区分q"])].sort_values("q", ascending=False).head(8).copy()
    ex["q"] = ex["q"].round(0); ex["p"] = ex["p"].round(1); ex["per"] = ex["per"].round(0)
    print(ex[["q", "p", "per", "区分q", "区分S"]].rename(columns={"q": "払出/月", "p": "患者/月", "per": "1人あたり"}).to_string())

    # ---- 図：散布（対数×対数）。色は区分S を一色の濃淡で ----
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager
    for f in ("Yu Gothic", "Meiryo", "MS Gothic"):
        if any(ff.name == f for ff in font_manager.fontManager.ttflist):
            plt.rcParams["font.family"] = f; break
    plt.rcParams["axes.unicode_minus"] = False
    RAMP = {"A": "#08519C", "B": "#3182BD", "C": "#6BAED6", "D": "#C6DBEF"}
    INK = "#333333"; GRID = "#E5E5E5"
    fig, ax = plt.subplots(figsize=(7.2, 5.6), dpi=150)
    dd = d[ok]
    for c in "ABCD":
        g = dd[dd["区分S"] == c]
        ax.scatter(g["lq"], g["lp"], s=22, color=RAMP[c], edgecolor="white", linewidth=0.5, label=f"区分 {c}（{len(g)}）")
    ax.axvline(mq - sq, color=INK, linewidth=1, linestyle="--"); ax.axvline(mq, color=INK, linewidth=1, linestyle="--"); ax.axvline(mq + sq, color=INK, linewidth=1, linestyle="--")
    ax.axhline(math.log10(a.p_line), color="#C0504D", linewidth=1, linestyle=":")
    ax.text(ax.get_xlim()[1], math.log10(a.p_line), f" 患者 {a.p_line:.0f} 人/月", va="bottom", ha="right", fontsize=8, color="#C0504D")
    for v, lab in ((mq - sq, "D|C"), (mq, "C|B"), (mq + sq, "B|A")):
        ax.text(v, ax.get_ylim()[1], lab, ha="center", va="bottom", fontsize=8, color=INK)
    ax.set_xlabel("log10（月平均払出）　破線＝量だけの区分の線（平均±σ）")
    ax.set_ylabel("log10（月平均の延べ患者数）")
    ax.set_title("量 × 患者の広がり（架空データ、両方ある薬）。色＝需要スコア S の区分", fontsize=10)
    ax.grid(color=GRID); ax.set_axisbelow(True)
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False)
    ax.legend(frameon=False, fontsize=8, loc="lower right")
    fig.tight_layout()
    os.makedirs(a.out, exist_ok=True)
    png = os.path.join(a.out, "EDA_需要スコア_散布.png")
    fig.savefig(png)
    print("\n図：", png)


if __name__ == "__main__":
    main()
