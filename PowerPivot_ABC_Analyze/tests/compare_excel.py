"""Excel の結果（tools/excel_dump.ps1 が書いた CSV）を tests/out/期待値.xlsx と突き合わせる
使い方（tests フォルダで）：python compare_excel.py <CSV のフォルダ> [期待値.xlsx]
  CSV：M_薬品.csv・T_月次.csv・T_判定理由.csv（excel_dump.ps1 の出力。UTF-8 BOM、日付 yyyy-MM-dd、真偽 True/False）
  突き合わせ：キーで結び、数値は相対差 0.1%（絶対差 0.001）まで許容、文字・日付は完全一致
  ※ 実データは使わない。期待値は make_testdata.py の架空データから計算したもの
"""
import sys, os, csv, math, datetime as dt
import openpyxl

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# 期待値の列名 → Excel（M_薬品）の列名
M_COLS = {"出どころ": "出どころ", "予定日数": "予定日数", "平均間隔": "平均間隔", "CV2": "ばらつき（CV²）", "需要の型": "需要の型",
          "動いた月数": "動いた月数", "払出合計": "払出合計", "一日平均払出": "1日平均払出", "最後に動いた月": "最後に動いた月",
          "最終在庫": "最終在庫", "動きのない月数": "動きのない月数", "在庫日数": "在庫日数", "その他の減少": "その他の減少",
          "ずれの割合": "ずれの割合", "ずれの大きさ": "ずれの大きさ", "予定外の月数": "予定外の月数", "経過措置期限": "経過措置期限",
          "直近の払出合計": "直近の払出合計", "直近の在庫日数": "直近の在庫日数"}
T_COLS = ["出納に載った", "払出数量", "予定数量", "予定とのずれ", "ずれの割合", "在庫数量", "その他の増減"]


def read_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    hdr = rows[0]
    return [dict(zip(hdr, r)) for r in rows[1:] if any(c != "" for c in r)]


def rows_of(ws):
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    hdr = [str(c) if c is not None else "" for c in rows[0]]
    return [dict(zip(hdr, r)) for r in rows[1:] if any(c is not None for c in r)]


def norm(v):
    """CSV の文字列も期待値の値も、同じ型にそろえる"""
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return None if (isinstance(v, float) and math.isnan(v)) else float(v)
    s = str(v).strip()
    if s == "":
        return None
    if s in ("True", "TRUE"):
        return True
    if s in ("False", "FALSE"):
        return False
    try:
        return dt.date.fromisoformat(s)
    except ValueError:
        pass
    try:
        return float(s)
    except ValueError:
        return s


def serial_to_date(n):
    """Excel の日付の通し番号（1900 年式）→ 日付"""
    return dt.date(1899, 12, 30) + dt.timedelta(days=int(n))


def same(a, b):
    a, b = norm(a), norm(b)
    if a is None and b is None:
        return True
    # 片方が日付、片方が数値なら、数値を Excel の通し番号とみなす（CSV に書き出すと日付が数値になることがある）
    if isinstance(a, dt.date) and isinstance(b, float):
        b = serial_to_date(b)
    if isinstance(b, dt.date) and isinstance(a, float):
        a = serial_to_date(a)
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b
    if isinstance(a, float) and isinstance(b, float):
        return abs(a - b) <= max(0.001, 0.001 * max(abs(a), abs(b)))
    return a == b


def main():
    folder = sys.argv[1]
    exp_path = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "out", "期待値.xlsx")
    exp = openpyxl.load_workbook(exp_path, data_only=True, read_only=True)
    bad = 0

    # ---- M_薬品 ----
    g = {r["薬品名"]: r for r in read_csv(os.path.join(folder, "M_薬品.csv"))}
    e = {r["薬品名"]: r for r in rows_of(exp["M_薬品"])}
    print(f"M_薬品：Excel {len(g)} 行、期待値 {len(e)} 行")
    for n in sorted(set(e) - set(g)):
        print("  期待値にだけある薬:", n); bad += 1
    for n in sorted(set(g) - set(e)):
        print("  Excel にだけある薬:", n); bad += 1
    for n in sorted(set(e) & set(g)):
        for ec, gc in M_COLS.items():
            if gc not in g[n]:
                print(f"  列がない: {gc}"); bad += 1; continue
            if not same(e[n][ec], g[n][gc]):
                print(f"  {n} / {gc}: 期待 {e[n][ec]!r} ≠ Excel {g[n][gc]!r}"); bad += 1

    # ---- T_判定理由（薬品×コードの集合） ----
    g2 = {(r["薬品名"], r["コード"]) for r in read_csv(os.path.join(folder, "T_判定理由.csv"))}
    e2 = {(r["薬品名"], r["コード"]) for r in rows_of(exp["T_判定理由"])}
    print(f"T_判定理由：Excel {len(g2)} 行、期待値 {len(e2)} 行")
    for k in sorted(e2 - g2):
        print("  期待値にだけ:", k); bad += 1
    for k in sorted(g2 - e2):
        print("  Excel にだけ:", k); bad += 1

    # ---- T_月次（期待値にある行だけ） ----
    def month_key(v):
        v = norm(v)
        return serial_to_date(v) if isinstance(v, float) else v
    g3 = {(month_key(r["月"]), r["薬品名"]): r for r in read_csv(os.path.join(folder, "T_月次.csv"))}
    e3 = rows_of(exp["T_月次"])
    print(f"T_月次：Excel {len(g3)} 行、期待値 {len(e3)} 行（期待値にある行だけ比べる）")
    for r in e3:
        k = (month_key(r["月"]), r["薬品名"])
        if k not in g3:
            print("  Excel にない行:", k); bad += 1; continue
        for c in T_COLS:
            if not same(r[c], g3[k][c]):
                print(f"  {k} / {c}: 期待 {r[c]!r} ≠ Excel {g3[k][c]!r}"); bad += 1

    print("違い:", bad)
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
