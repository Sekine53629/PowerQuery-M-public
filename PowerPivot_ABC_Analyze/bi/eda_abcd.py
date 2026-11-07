"""EDA：使用量（月平均払出）を対数にして ABCD の 4 区分に分ける切り方を比べる
使い方（プロジェクトのフォルダで）：python bi/eda_abcd.py [--model bi/out/モデル.xlsx] [--out docs]
  入力：bi/run.py が書いた モデル.xlsx の T_月次（出納のある月）と M_薬品
  出力：docs/EDA_ABCD区分_分布.png（分布と線）、標準出力に各方式の区分の数・構成比・月ごとの区分の入れ替わり
  ※ 架空データ（tests/out_clinic、1 年・約 200 品目）で見る。線の案は docs/EDA_ABCD区分.md に書く
"""
import os, sys, math, argparse
import numpy as np
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


# ------------------------------------------------------------------------------
# 区分の方式（どれも 月平均払出 x を受け取り "A"/"B"/"C"/"D" を返す。x <= 0 は D）
# ------------------------------------------------------------------------------
def by_log_fixed(x, lines=(1000, 100, 10)):
    """対数の等間隔：10 の累乗の線。A ≧ 1000／月、B ≧ 100、C ≧ 10、D それ未満（0 を含む）"""
    if x is None or x <= 0:
        return "D"
    return "A" if x >= lines[0] else "B" if x >= lines[1] else "C" if x >= lines[2] else "D"


def by_log_sigma(x, mu, sd):
    """対数の平均±標準偏差：A > μ+σ、B μ～μ+σ、C μ−σ～μ、D < μ−σ"""
    if x is None or x <= 0:
        return "D"
    z = (math.log10(x) - mu) / sd
    return "A" if z > 1 else "B" if z > 0 else "C" if z > -1 else "D"


def by_cum_share(values, a=0.7, b=0.9):
    """累積構成比（今の T_ABC月次）：払出の多い順に累積し、A ≦ 70%、B ≦ 90%、残り C。0 は D"""
    s = pd.Series(values).fillna(0).sort_values(ascending=False)
    tot = s.sum()
    cum = s.cumsum() / tot if tot else s * 0
    cls = {}
    for n, c, v in zip(s.index, cum, s):
        cls[n] = "D" if v <= 0 else "A" if c <= a else "B" if c <= b else "C"
    return cls


def jenks_breaks(values, k=4):
    """自然分類（Jenks）：クラス内の分散が最小になる k 区分の境界（小さい方からの上限値を返す）"""
    v = sorted(values)
    n = len(v)
    if n <= k:
        return v
    lower = [[0] * (k + 1) for _ in range(n + 1)]
    var = [[float("inf")] * (k + 1) for _ in range(n + 1)]
    for i in range(1, k + 1):
        lower[1][i] = 1; var[1][i] = 0.0
    for l in range(2, n + 1):
        s1 = s2 = 0.0; w = 0
        for m in range(1, l + 1):
            i3 = l - m + 1
            val = v[i3 - 1]
            s2 += val * val; s1 += val; w += 1
            vv = s2 - (s1 * s1) / w
            i4 = i3 - 1
            if i4 != 0:
                for j in range(2, k + 1):
                    if var[l][j] >= vv + var[i4][j - 1]:
                        lower[l][j] = i3; var[l][j] = vv + var[i4][j - 1]
        lower[l][1] = 1; var[l][1] = vv
    breaks = [0.0] * k
    kk = n
    for j in range(k, 1, -1):
        idx = int(lower[kk][j]) - 2
        breaks[j - 2] = v[idx]
        kk = int(lower[kk][j]) - 1
    breaks[k - 1] = v[-1]
    return breaks


def by_jenks(x, br):
    """自然分類の境界（対数値、昇順の上限 3 つ）：下から D・C・B・A"""
    if x is None or x <= 0:
        return "D"
    lx = math.log10(x)
    return "D" if lx <= br[0] else "C" if lx <= br[1] else "B" if lx <= br[2] else "A"


def summarize(name, cls_map, usage):
    """区分ごとの薬品数・払出の構成比・月平均の範囲"""
    df = pd.DataFrame({"区分": pd.Series(cls_map), "月平均": pd.Series(usage)}).dropna(subset=["区分"])
    tot = df["月平均"].sum()
    rows = []
    for c in "ABCD":
        g = df[df["区分"] == c]
        rows.append(dict(方式=name, 区分=c, 薬品数=len(g), 構成比=(g["月平均"].sum() / tot if tot else 0),
                         最小=(g["月平均"].min() if len(g) else None), 最大=(g["月平均"].max() if len(g) else None)))
    return pd.DataFrame(rows)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--model", default=os.path.join(ROOT, "bi", "out", "モデル.xlsx"))
    ap.add_argument("--out", default=os.path.join(ROOT, "docs"))
    a = ap.parse_args()

    x = pd.ExcelFile(a.model)
    t = pd.read_excel(x, "T_月次")
    t = t[t["出納データあり"] == True]
    months = sorted(t["月"].unique())
    n_months = len(months)
    usage = t.groupby("薬品名")["払出数量"].sum().fillna(0) / n_months       # 月平均払出（出納のある月で割る）
    monthly = t.pivot_table(index="薬品名", columns="月", values="払出数量", aggfunc="sum").fillna(0)
    pos = usage[usage > 0]
    logs = np.log10(pos)
    mu, sd = float(logs.mean()), float(logs.std(ddof=0))
    print(f"薬品 {len(usage)} 品目、出納の月 {n_months}、払出 0 の薬 {int((usage <= 0).sum())} 品目")
    print(f"log10(月平均払出)：平均 {mu:.2f}（＝{10**mu:,.0f}/月）、標準偏差 {sd:.2f}、最小 {logs.min():.2f}、最大 {logs.max():.2f}")
    q = np.percentile(logs, [10, 25, 50, 75, 90])
    print("  分位（10/25/50/75/90%）：" + " / ".join(f"{10**v:,.0f}" for v in q) + " /月")
    br = jenks_breaks(list(logs), 4)
    print(f"自然分類（Jenks）の境界：" + " / ".join(f"{10**b:,.0f}" for b in br[:3]) + " /月")

    methods = {
        "累積構成比 70/90（今）": by_cum_share(usage),
        "対数の等間隔 1000/100/10": {n: by_log_fixed(v) for n, v in usage.items()},
        "対数の平均±σ": {n: by_log_sigma(v, mu, sd) for n, v in usage.items()},
        "自然分類（Jenks）": {n: by_jenks(v, br) for n, v in usage.items()},
    }
    tbl = pd.concat([summarize(k, m, usage) for k, m in methods.items()], ignore_index=True)
    pd.set_option("display.width", 200)
    print("\n■ 区分ごとの薬品数と払出の構成比（年の月平均で分けたとき）")
    print(tbl.to_string(index=False, formatters={"構成比": "{:.0%}".format, "最小": "{:,.0f}".format, "最大": "{:,.0f}".format}))

    # ---- 月ごとに分け直したときの入れ替わり（安定性） ----
    print("\n■ 月ごとの払出で分け直したとき、前の月から区分が変わった薬の数（12 か月の平均）")
    for name in methods:
        changes = []
        prev = None
        for m in months:
            u = monthly[m]
            if name.startswith("累積"):
                cur = by_cum_share(u)
            elif name.startswith("対数の等"):
                cur = {n: by_log_fixed(v) for n, v in u.items()}
            elif name.startswith("対数の平"):
                cur = {n: by_log_sigma(v, mu, sd) for n, v in u.items()}
            else:
                cur = {n: by_jenks(v, br) for n, v in u.items()}
            if prev is not None:
                changes.append(sum(1 for n in cur if cur[n] != prev.get(n)))
            prev = cur
        print(f"  {name:<22} 平均 {np.mean(changes):5.1f} 品目/月（最大 {max(changes)}）")

    # ---- 図：分布と線 ----
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager
    for f in ("Yu Gothic", "Meiryo", "MS Gothic"):
        if any(ff.name == f for ff in font_manager.fontManager.ttflist):
            plt.rcParams["font.family"] = f; break
    plt.rcParams["axes.unicode_minus"] = False
    BLUE = "#3B6EA5"; INK = "#333333"; MUTED = "#777777"; GRID = "#E5E5E5"
    fig, axes = plt.subplots(1, 2, figsize=(11, 4.2), dpi=150)
    ax = axes[0]
    bins = np.arange(math.floor(logs.min() * 4) / 4, math.ceil(logs.max() * 4) / 4 + 0.25, 0.25)
    ax.hist(logs, bins=bins, color=BLUE, edgecolor="white", linewidth=0.8)
    for v, lab in ((3, "1000"), (2, "100"), (1, "10")):
        ax.axvline(v, color=INK, linewidth=1, linestyle="--")
        ax.text(v, ax.get_ylim()[1] * 0.97, f"{lab}/月", ha="center", va="top", fontsize=8, color=INK)
    for v in (mu - sd, mu, mu + sd):
        ax.axvline(v, color=MUTED, linewidth=1, linestyle=":")
    ax.text(mu, ax.get_ylim()[1] * 0.80, "μ", ha="center", fontsize=8, color=MUTED)
    ax.set_xlabel("log10（月平均払出）　点線＝平均±σ、破線＝10の累乗")
    ax.set_ylabel("薬品数")
    ax.set_title(f"月平均払出の分布（払出ありの {len(pos)} 品目、{n_months} か月）", fontsize=10)
    ax.grid(axis="y", color=GRID); ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax2 = axes[1]
    s = usage.sort_values(ascending=False)
    cum = s.cumsum() / s.sum()
    ax2.plot(range(1, len(s) + 1), cum.values, color=BLUE, linewidth=2)
    for y, lab in ((0.7, "70%"), (0.9, "90%")):
        ax2.axhline(y, color=INK, linewidth=1, linestyle="--")
        k = int((cum <= y).sum())
        ax2.text(len(s), y, f" {lab}：{k} 品目", va="bottom", ha="right", fontsize=8, color=INK)
    ax2.set_xlabel("払出の多い順の薬品数")
    ax2.set_ylabel("累積構成比")
    ax2.set_ylim(0, 1.02)
    ax2.set_title("累積構成比（今の ABC の切り方）", fontsize=10)
    ax2.grid(color=GRID); ax2.set_axisbelow(True)
    for sp in ("top", "right"):
        ax2.spines[sp].set_visible(False)
    fig.tight_layout()
    os.makedirs(a.out, exist_ok=True)
    png = os.path.join(a.out, "EDA_ABCD区分_分布.png")
    fig.savefig(png)
    print("\n図：", png)


if __name__ == "__main__":
    main()
