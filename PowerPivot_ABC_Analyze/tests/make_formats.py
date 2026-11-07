"""架空データ（tests/out の .xlsx）を、ほかの形式（.csv／.txt）に書き直す。Python 版で形式ごとの結果が同じかも確かめる
使い方（プロジェクトのフォルダで）：
  python tests/make_testdata.py          … 先に .xlsx を作る（tests/out）
  python tests/make_formats.py           … tests/out/形式別/csv と tests/out/形式別/txt を作る
  pwsh -File tools/make_xls.ps1          … tests/out/形式別/xls を作る（Excel を使う。自宅 PC だけ）
  python tests/make_formats.py --check   … 形式ごとに Python 版（bi/model.py）で組み立て、.xlsx のときと同じ結果か確かめる
作るもの（tests/out/形式別/<形式>/ の下。フォルダの構成は tests/out と同じ）
  csv（カンマ区切り・Shift_JIS）
    実績/薬品使用量分析データYYYYMMDD.csv … 1 日 1 ファイル（日付はファイル名の 8 桁）
    出納/…_数量一覧.csv                   … 見出しより上の行（出力日・期間）もそのまま
    システムリスト/…                      … 3 つのリスト
  txt（タブ区切り）
    実績/薬品使用量分析データYYYYMM.txt    … 1 か月 1 ファイル。先頭に「日付」の列（UTF-8・BOM あり）
    出納/…_数量一覧.txt                   … UTF-8・BOM なし
    システムリスト/…                      … UTF-16（Excel の「Unicode テキスト」と同じ）
  どの形式にも 手入力テーブル.xlsx（フォルダパスだけその形式のフォルダに変えたもの）を置く
※ エリシアSが実際に出す .csv／.txt の並びは未確認。ここは「読み方の決まり」を確かめるための形
"""
import os, re, sys, csv, io, shutil, datetime as dt
import openpyxl

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
OUT = os.path.join(HERE, "out")
FMT = os.path.join(OUT, "形式別")
LISTS = os.path.join("システムリスト")
SETTING_PATHS = {"フォルダパス": "実績", "出納フォルダパス": "出納",
                 "期限切迫フォルダパス": os.path.join("システムリスト", "期限切れ"),
                 "経過措置フォルダパス": os.path.join("システムリスト", "経過措置薬品"),
                 "不動品フォルダパス": os.path.join("システムリスト", "不動薬")}


def cell(v):
    """セルの値 → 文字（空欄は空文字、整数は小数点なし、日付は yyyy/mm/dd）"""
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.strftime("%Y/%m/%d")
    if isinstance(v, dt.date):
        return v.strftime("%Y/%m/%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def write_text(path, rows, delim, enc):
    buf = io.StringIO()
    w = csv.writer(buf, delimiter=delim, lineterminator="\r\n")
    for r in rows:
        w.writerow([cell(c) for c in r])
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding=enc, newline="") as f:
        f.write(buf.getvalue())


def sheets_of(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    return [(ws.title, [list(r) for r in ws.iter_rows(values_only=True)]) for ws in wb.worksheets]


def trim(rows):
    """行の後ろの空欄を落とす（表題の行が短い、実際の文字のファイルに近づける）"""
    out = []
    for r in rows:
        r = list(r)
        while r and r[-1] is None:
            r.pop()
        out.append(r)
    return out


def hand_tables(dst_root):
    """手入力テーブル.xlsx を写し、フォルダパスだけ書き換える"""
    src = os.path.join(OUT, "手入力テーブル.xlsx")
    dst = os.path.join(dst_root, "手入力テーブル.xlsx")
    wb = openpyxl.load_workbook(src)
    ws = wb["設定"]
    for row in ws.iter_rows(min_row=2):
        if row[0].value in SETTING_PATHS:
            row[1].value = os.path.join(dst_root, SETTING_PATHS[row[0].value])
    wb.save(dst)


def make():
    for kind in ("csv", "txt"):
        d = os.path.join(FMT, kind)
        if os.path.isdir(d):
            shutil.rmtree(d)
        os.makedirs(d)
    C, T = os.path.join(FMT, "csv"), os.path.join(FMT, "txt")
    n = 0
    # ---- 実績 ----
    for f in sorted(os.listdir(os.path.join(OUT, "実績"))):
        m = re.match(r"(.+?)(\d{4})(\d{2})\.xlsx$", f)
        if not m:
            continue
        base, fy, fm = m.group(1), int(m.group(2)), int(m.group(3))
        long_rows = None
        for title, rows in sheets_of(os.path.join(OUT, "実績", f)):
            mm = re.match(r"(\d{1,2})月(\d{1,2})日", title)
            if not mm:
                continue                                   # 「期間すべて」は書かない
            d = dt.date(fy, int(mm.group(1)), int(mm.group(2)))
            write_text(os.path.join(C, "実績", f"{base}{d:%Y%m%d}.csv"), rows, ",", "cp932"); n += 1
            if long_rows is None:
                long_rows = [["日付"] + list(rows[0])]
            long_rows += [[d.strftime("%Y/%m/%d")] + list(r) for r in rows[1:] if any(c is not None for c in r)]
        if long_rows:
            write_text(os.path.join(T, "実績", f"{base}{fy:04d}{fm:02d}.txt"), long_rows, "\t", "utf-8-sig"); n += 1
    # ---- 出納 ----
    for f in sorted(os.listdir(os.path.join(OUT, "出納"))):
        if not f.endswith(".xlsx"):
            continue
        rows = trim(sheets_of(os.path.join(OUT, "出納", f))[0][1])
        stem = f[:-5]
        write_text(os.path.join(C, "出納", stem + ".csv"), rows, ",", "cp932")
        write_text(os.path.join(T, "出納", stem + ".txt"), rows, "\t", "utf-8"); n += 2
    # ---- 3 つのリスト ----
    for sub in sorted(os.listdir(os.path.join(OUT, LISTS))):
        for f in sorted(os.listdir(os.path.join(OUT, LISTS, sub))):
            if not f.endswith(".xlsx"):
                continue
            rows = sheets_of(os.path.join(OUT, LISTS, sub, f))[0][1]
            stem = f[:-5]
            write_text(os.path.join(C, LISTS, sub, stem + ".csv"), rows, ",", "cp932")
            write_text(os.path.join(T, LISTS, sub, stem + ".txt"), rows, "\t", "utf-16"); n += 2
    for kind in ("csv", "txt"):
        hand_tables(os.path.join(FMT, kind))
    if os.path.isdir(os.path.join(FMT, "xls")):
        hand_tables(os.path.join(FMT, "xls"))
    print(f"できた：{FMT}（csv・txt 合わせて {n} ファイル）")


def check():
    """形式ごとに Python 版で組み立て、.xlsx のときと同じ結果か確かめる（ファイル名・シート名の列は比べない）"""
    sys.path.insert(0, os.path.join(ROOT, "bi"))
    import model
    import pandas as pd
    skip = {"ファイル名", "シート名"}
    names = ["日次データ", "予測データ", "出納", "経過措置該当", "期限切れリスト", "不動品リスト", "T_日次", "T_月次", "T_ABC月次", "M_薬品", "薬品グループ案", "T_ロット", "T_判定理由"]
    base = model.Model(model.Settings(os.path.join(OUT, "手入力テーブル.xlsx")))
    bad = 0
    for kind in ("xls", "csv", "txt"):
        tables = os.path.join(FMT, kind, "手入力テーブル.xlsx")
        if not os.path.isdir(os.path.join(FMT, kind)):
            print(f"{kind}：フォルダがないので飛ばす"); continue
        if not os.path.exists(tables):
            hand_tables(os.path.join(FMT, kind))
        m = model.Model(model.Settings(tables))
        diffs = []
        for nme in names:
            a, b = getattr(base, nme), getattr(m, nme)
            cols = [c for c in a.columns if c not in skip]
            a2 = a[cols].reset_index(drop=True)
            b2 = b[[c for c in cols if c in b.columns]].reset_index(drop=True)
            try:
                pd.testing.assert_frame_equal(a2, b2, check_dtype=False, check_exact=False, rtol=1e-9)
            except AssertionError as e:
                diffs.append(f"  {nme}：{str(e).splitlines()[0]} … {len(a2)} 行と {len(b2)} 行")
        print(f"{kind}：" + ("同じ（" + "・".join(f"{n} {len(getattr(m, n))}" for n in ("日次データ", "出納", "M_薬品", "T_月次", "T_判定理由", "T_ロット")) + "）"
                             if not diffs else "ちがう"))
        for d in diffs:
            print(d)
        bad += len(diffs)
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    if "--check" in sys.argv:
        check()
    else:
        make()
