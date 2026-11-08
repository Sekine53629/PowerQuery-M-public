"""PowerPivot_ABC と同じデータモデルを Python（pandas）で組み立てる
   入力：薬品使用量予測データ（実績フォルダ）、出納（出納フォルダ）、手入力テーブル.xlsx（設定ほか）
   出力：T_日次・T_月次・M_日付・M_薬品・M_理由コード・T_判定理由（列名・定義は M クエリ ⑤～⑪ に合わせる）
   ※ Power Pivot 側の結果と突き合わせるための「正解」役。定義を変えるときは M クエリと両方直す
"""
import os, re, io, csv, math, glob, unicodedata, datetime as dt
from collections import defaultdict
import pandas as pd
import numpy as np
import openpyxl
try:
    import xlrd          # .xls（Excel 97-2003）を読むときだけ使う
except ImportError:
    xlrd = None

# ------------------------------------------------------------------------------
# 共通
# ------------------------------------------------------------------------------
def nz(v):
    """fx_正規化：全角英数→半角、半角カナ→全角、前後の空白を除く"""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    s = unicodedata.normalize("NFKC", str(v)).strip()
    return s if s != "" else None


MARKS = "劇向毒麻覚習生"


def strip_mark(v):
    """品名の先頭の記号（劇 ／向 など 1 文字＋空白）と先頭空白を外す。戻り値：(薬品名, 記号)"""
    s = nz(v)
    if s is None:
        return None, None
    if len(s) >= 2 and s[0] in MARKS and s[1] == " ":
        return s[2:].strip() or None, s[0]
    return s, None


def norm_col(v):
    """列名の照合：正規化し、空白・改行を除き、かっこを全角にそろえる"""
    s = nz(v) or ""
    for ch in (" ", "　", "\n", "\r", "\t"):
        s = s.replace(ch, "")
    return s.replace("(", "（").replace(")", "）")


def num(v):
    """数値：全角数字・桁区切り・単位を除いて数値にする"""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = re.sub(r"[ ,]", "", unicodedata.normalize("NFKC", str(v)))
    m = re.search(r"-?\d+(\.\d+)?", t)
    return float(m.group(0)) if m else None


def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = unicodedata.normalize("NFKC", str(v)).strip().replace("年", "/").replace("月", "/").replace("日", "").replace("-", "/").replace(".", "/")
    m = re.match(r"(\d{4})/(\d{1,2})/(\d{1,2})", s)
    if m:
        return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.match(r"(\d{4})(\d{2})(\d{2})$", s)
    if m:
        return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def month_start(d):
    return dt.date(d.year, d.month, 1)


def month_end(d):
    n = dt.date(d.year + (d.month == 12), d.month % 12 + 1, 1)
    return n - dt.timedelta(days=1)


def add_years(d, n):
    try:
        return d.replace(year=d.year + n)
    except ValueError:                     # 2/29
        return d.replace(year=d.year + n, day=28)


def months_between(a, b):
    return (b.year - a.year) * 12 + b.month - a.month


# ------------------------------------------------------------------------------
# 設定（手入力テーブル.xlsx）
# ------------------------------------------------------------------------------
DEFAULTS = dict(不動とみなす月数=3, 在庫過多とみなす日数=60,
                ずれの注意割合=0.3, ずれの大きさの注意割合=0.5,
                # 月次レポート用（0.2.0 で設定テーブルに追加する項目。空欄なら初期値）
                当面の予定日数=30, ABCの月数=3,
                # 0.5.0：区分は量の固定線（錠／月）と予定の患者数で付ける（累積構成比の線 ABCのAの線・ABCのBの線 は廃止）
                区分Aの線=1000, 区分Bの線=300, 区分Dの患者数=2,
                予測日数=7,                      # 0.6.0：発注書 F が読む来週の日数
                出荷調整の構成比の変化=0.3, 出荷調整の合計の許容=0.3, 期限切迫の日数=180, 直近の月数=3,
                # 0.3.0：包装単位を使わない。極小需要の線は「月の払出」1 本（G02・A の回転・E）
                極小需要の月の払出=100, 傾向の線=0.2)


class Settings:
    def __init__(self, path):
        wb = openpyxl.load_workbook(path, data_only=True)
        self.raw = {}
        for r in wb["設定"].iter_rows(min_row=2, values_only=True):
            if r[0] is not None:
                self.raw[nz(r[0])] = r[1]
        self.tables = {}
        for name in ("経過措置リスト", "発注除外", "在庫情報", "追加休業日", "薬品グループ"):
            if name in wb.sheetnames:
                ws = wb[name]
                rows = list(ws.iter_rows(values_only=True))
                self.tables[name] = pd.DataFrame(rows[1:], columns=[nz(c) for c in rows[0]]) if len(rows) > 1 else pd.DataFrame(columns=[nz(c) for c in rows[0]])
            else:
                self.tables[name] = pd.DataFrame()
        self.folder = self.text("フォルダパス")
        self.sui_folder = self.text("出納フォルダパス")
        self.kigen_folder = self.text("期限切迫フォルダパス")
        self.keika_folder = self.text("経過措置フォルダパス")
        self.fudo_folder = self.text("不動品フォルダパス")
        self.base_date = to_date(self.raw.get("発注日")) or dt.date.today()
        # 実績を読む期間（① 設定値と同じ）：設定に開始日・終了日があればその値。なければ
        #   終了日 ＝ 予測開始日（発注日の翌日）の前日 ＝ 発注日、開始日 ＝ 終了日の 1 年前の翌日が属する月の 1 日
        fc_start = to_date(self.raw.get("予測開始日")) or (self.base_date + dt.timedelta(days=1))
        self.end = to_date(self.raw.get("終了日")) or (fc_start - dt.timedelta(days=1))
        self.start = to_date(self.raw.get("開始日")) or month_start(add_years(self.end, -1) + dt.timedelta(days=1))
        # 発注書 F：来週（予測開始～予測終了）。予測フォルダパスが空欄なら実績フォルダ（当月の予測データに月末までの予定がある）
        self.fc_start = fc_start
        self.fc_end = fc_start + dt.timedelta(days=int(self.num("予測日数")) - 1)
        self.fc_folder = self.text("予測フォルダパス") or self.folder
        ye_s, ye_e = self.raw.get("年末年始の休み始め") or "12/29", self.raw.get("年末年始の休み終わり") or "1/3"
        self.ye_start = tuple(int(x) for x in str(ye_s).replace("-", "/").split("/")[-2:])
        self.ye_end = tuple(int(x) for x in str(ye_e).replace("-", "/").split("/")[-2:])

    def text(self, key):
        v = self.raw.get(key)
        return None if v is None else str(v).strip().strip('"').rstrip("\\/ ")

    def num(self, key, dflt=None):
        v = self.raw.get(key)
        if v is None or (isinstance(v, str) and v.strip() == ""):
            return DEFAULTS.get(key, dflt) if dflt is None else dflt
        return float(v)


def first_table(path, label):
    """先頭 30 行に見出し label がある最初の表の行（なければ 1 枚目）。.xls はシートが名前の順に並ぶため（㉒ と同じ）"""
    tables = open_tables(path)
    for _, rows, _ in tables:
        if any(norm_col(label) in [norm_col(c) for c in r] for r in rows[:30]):
            return rows
    return tables[0][1] if tables else []


# ------------------------------------------------------------------------------
# 薬品使用量予測データ（予定）：fx_使用量読込 ＋ ③ 日次データ
# ------------------------------------------------------------------------------
COLS = dict(薬品名="薬品名", 数量="予定使用量", 在庫量="在庫数量（発注残を含まない）", 在庫量込み="在庫数量",
            薬品区分="製剤", 患者数="患者数", 略称="略称")
WD = "月火水木金土日"


def _sheet_rows(ws):
    return [list(r) for r in ws.iter_rows(values_only=True)]


# ------------------------------------------------------------------------------
# ファイルを開く：㉒ fx_ファイルを開く と同じ（.xls／.xlsx／.csv／.txt。形式は中身で見分ける）
# ------------------------------------------------------------------------------
SUPPORTED_EXT = (".xls", ".xlsx", ".xlsm", ".csv", ".txt", ".tsv")


def list_files(folder):
    """フォルダの中の、読める形式のファイル（Excel の一時ファイル「~$…」を除く）"""
    if not folder or not os.path.isdir(folder):
        return []
    return [os.path.join(folder, f) for f in sorted(os.listdir(folder))
            if os.path.splitext(f)[1].lower() in SUPPORTED_EXT and not f.startswith("~$")
            and os.path.isfile(os.path.join(folder, f))]


def open_tables(path):
    """戻り値：[(シート名, 行のリスト, 形式)]。文字のファイルは 1 つだけ（シート名は空、値はすべて文字、空欄は None）"""
    name = os.path.basename(path)
    with open(path, "rb") as f:
        data = f.read()
    if not data:
        raise ValueError(f"「{name}」は空のファイルです")
    if data[:4] == b"\xD0\xCF\x11\xE0":                      # Excel 97-2003（.xls）
        if xlrd is None:
            raise ValueError(f"「{name}」：.xls を読むには xlrd が必要です（pip install xlrd）")
        wb = xlrd.open_workbook(file_contents=data)
        out = []
        for sh in wb.sheets():
            rows = []
            for i in range(sh.nrows):
                r = []
                for c in sh.row(i):
                    if c.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK) or c.value == "":
                        r.append(None)
                    elif c.ctype == xlrd.XL_CELL_DATE:
                        r.append(xlrd.xldate_as_datetime(c.value, wb.datemode))
                    elif c.ctype == xlrd.XL_CELL_NUMBER and float(c.value).is_integer():
                        r.append(int(c.value))                # .xls の数値はすべて小数で返るので、整数は整数に戻す
                    else:
                        r.append(c.value)
                rows.append(r)
            out.append((sh.name, rows, "Excel（.xls）"))
        return out
    if data[:2] == b"PK":                                      # .xlsx
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        return [(ws.title, _sheet_rows(ws), "Excel（.xlsx）") for ws in wb.worksheets]
    body = data[3:] if data[:3] == b"\xef\xbb\xbf" else data
    if body[:1] == b"<":
        raise ValueError(f"「{name}」は中身が HTML か XML です。Excel で開いて .xlsx か .csv で保存し直してください")
    # 文字コード：BOM があればそれに従う。なければ UTF-8 として読んでみて、読めなければ Shift_JIS
    if data[:3] == b"\xef\xbb\xbf":
        enc, label = "utf-8-sig", "UTF-8"
    elif data[:2] in (b"\xff\xfe", b"\xfe\xff"):
        enc, label = "utf-16", "UTF-16"
    else:
        try:
            data.decode("utf-8"); enc, label = "utf-8", "UTF-8"
        except UnicodeDecodeError:
            enc, label = "cp932", "Shift_JIS"
    txt = data.decode(enc)
    # 区切り：先頭 50 行（空行を除く）で、タブとカンマの多い方
    head = [l for l in txt.splitlines() if l.strip() != ""][:50]
    tabs, commas = sum(l.count("\t") for l in head), sum(l.count(",") for l in head)
    delim = "\t" if tabs > 0 and tabs >= commas else ","
    rows = [[(c if c != "" else None) for c in r] for r in csv.reader(io.StringIO(txt, newline=""), delimiter=delim)]
    n = max((len(r) for r in rows), default=0)
    rows = [r + [None] * (n - len(r)) for r in rows]
    kind = f"文字（{'カンマ区切り' if delim == ',' else 'タブ区切り'}・{label}）"
    return [("", rows, kind)]


def _plan_parts(sheet, rows, kind, fy, fm, d8, start, end):
    """1 枚の表を (日付, 行のリスト) に分ける。
       Excel のシート：シート名「10月1日(水)」とファイルの年月から日付を決める
       文字のファイル：見出しに「日付」の列があればその列で日ごとに分ける。なければファイル名の 8 桁の日付の 1 日分"""
    if kind.startswith("文字"):
        hidx = next((i for i, r in enumerate(rows[:30]) if norm_col(COLS["薬品名"]) in [norm_col(c) for c in r]), None)
        hdr = [norm_col(c) for c in rows[hidx]] if hidx is not None else []
        if "日付" in hdr:
            di, groups = hdr.index("日付"), {}
            for r in rows[hidx + 1:]:
                d = to_date(r[di]) if di < len(r) else None
                if d and start <= d <= end:
                    groups.setdefault(d, [rows[hidx]]).append(r)
            return list(groups.items())
        return [(d8, rows)] if d8 and start <= d8 <= end else []
    mm = re.match(r"\s*(\d{1,2})月(\d{1,2})日", sheet)
    if not mm or sheet.startswith("期間"):
        return []
    mo, dd = int(mm.group(1)), int(mm.group(2))
    if mo == fm:
        y = fy
    elif mo == (1 if fm == 12 else fm + 1):
        y = fy + 1 if fm == 12 else fy
    else:
        return []
    try:
        d = dt.date(y, mo, dd)
    except ValueError:
        return []
    return [(d, rows)] if start <= d <= end else []


def load_plan(folder, start, end, prefix="薬品使用量", keep_zero=False):
    """日次データ：日付／薬品区分／略称／薬品名／患者数／数量／在庫量／在庫量_発注残込み／ファイル名／シート名
       keep_zero：True なら数量も患者数も 0 の行も残す（在庫だけ見たいとき。② の keepZero と同じ）"""
    files = [f for f in list_files(folder) if os.path.basename(f).startswith(prefix)]
    recs = {}   # (date) -> (sortkey, rows)
    for f in sorted(files):
        name = os.path.basename(f)
        m = re.search(r"\d+", os.path.splitext(nz(name[len(prefix):]) or "")[0])     # 先頭の後ろで最初に出てくる数字の並び
        if not m or len(m.group(0)) < 6:
            continue
        dg = m.group(0)
        fy, fm = int(dg[:4]), int(dg[4:6])
        d8 = to_date(dg[:8]) if len(dg) >= 8 else None
        sortkey = (os.path.getmtime(f), name)
        for s, rows0, kind in open_tables(f):
            for d, rows in _plan_parts(s, rows0, kind, fy, fm, d8, start, end):
                if d in recs and recs[d][0] > sortkey:
                    continue
                hidx = next((i for i, r in enumerate(rows) if norm_col(COLS["薬品名"]) in [norm_col(c) for c in r]), None)
                if hidx is None:
                    continue
                hdr = [norm_col(c) for c in rows[hidx]]
                idx = {k: (hdr.index(norm_col(v)) if norm_col(v) in hdr else None) for k, v in COLS.items()}
                if idx["略称"] is None:
                    cand = [i for i, h in enumerate(hdr) if "略称" in h]
                    idx["略称"] = cand[0] if cand else (1 if len(hdr) >= 2 else None)
                out = []
                for r in rows[hidx + 1:]:
                    get = lambda k: (r[idx[k]] if idx[k] is not None and idx[k] < len(r) else None)
                    drug = nz(get("薬品名"))
                    if not drug or norm_col(drug) in ("合計", "総合計", "小計", norm_col(COLS["薬品名"])):
                        continue
                    q, p = num(get("数量")), num(get("患者数"))
                    if not keep_zero and not ((q is not None and q > 0) or (p is not None and p > 0)):
                        continue
                    out.append(dict(日付=d, 薬品区分=nz(get("薬品区分")), 略称=(str(get("略称")).strip() if get("略称") is not None else None),
                                    薬品名=drug, 患者数=p, 数量=q, 在庫量=num(get("在庫量")), 在庫量_発注残込み=num(get("在庫量込み")),
                                    ファイル名=name, シート名=s))
                recs[d] = (sortkey, out)
    rows = [r for _, (_, rs) in sorted(recs.items()) for r in rs]
    return pd.DataFrame(rows, columns=["日付", "薬品区分", "略称", "薬品名", "患者数", "数量", "在庫量", "在庫量_発注残込み", "ファイル名", "シート名"])


# ------------------------------------------------------------------------------
# 出納（実際）：⑤
# ------------------------------------------------------------------------------
SUI_COLS = ["製剤コード", "製剤", "コード", "品名", "繰越数量", "仕入数量", "払出数量", "在庫数量", "単位", "棚番"]


def _dates_in(v):
    if v is None:
        return []
    if isinstance(v, (dt.date, dt.datetime)):
        return [v.date() if isinstance(v, dt.datetime) else v]
    if isinstance(v, (int, float)):
        s = str(int(v))
        return [to_date(s)] if len(s) == 8 and to_date(s) else []
    s = unicodedata.normalize("NFKC", str(v))
    for ch in "年月-.":
        s = s.replace(ch, "/")
    s = s.replace("日", "")
    toks = re.split(r"[^0-9/]+", s)
    return [d for d in (to_date(t) for t in toks if t) if d and 1990 <= d.year <= 2100]


def load_suitou(folder, start=None, end=None, mark="数量一覧"):
    """出納：期間が start～end にかかるファイルだけ読む（⑤ と同じ。None なら全部）"""
    files = [f for f in list_files(folder) if mark in nz(os.path.basename(f))]
    by_period = {}
    for f in files:
        name = os.path.basename(f)
        rows = first_table(f, "品名")
        hidx = next((i for i, r in enumerate(rows) if "品名" in [norm_col(c) for c in r]), None)
        if hidx is None:
            continue
        period = None
        ds = [d for r in rows[:hidx][:30] for c in r for d in _dates_in(c)]
        for i in range(len(ds) - 1):
            if ds[i] <= ds[i + 1]:
                period = (ds[i], ds[i + 1]); break
        if period is None:
            toks = [t for t in re.split(r"[_\-. （）()]", nz(name)) if len(t) == 8 and t.isdigit()]
            if len(toks) >= 2 and to_date(toks[0]) and to_date(toks[1]) and to_date(toks[0]) <= to_date(toks[1]):
                period = (to_date(toks[0]), to_date(toks[1]))
        if period is None:
            continue
        if (start and period[1] < start) or (end and period[0] > end):
            continue
        key = period
        sortkey = os.path.getmtime(f)
        if key in by_period and by_period[key][0] > sortkey:
            continue
        hdr = [norm_col(c) for c in rows[hidx]]
        idx = {c: (hdr.index(c) if c in hdr else None) for c in SUI_COLS}
        out = []
        for r in rows[hidx + 1:]:
            get = lambda c: (r[idx[c]] if idx[c] is not None and idx[c] < len(r) else None)
            raw = get("品名")
            pname, pmark = strip_mark(raw)
            if not pname or norm_col(pname) in ("合計", "総合計", "小計", "品名"):
                continue
            lead = len(str(raw)) - len(str(raw).lstrip(" 　")) if raw is not None else 0
            rec = dict(期間開始=period[0], 期間終了=period[1], 日数=(period[1] - period[0]).days + 1,
                       製剤コード=(str(get("製剤コード")).strip() if get("製剤コード") is not None else None),
                       製剤=nz(get("製剤")), コード=(str(get("コード")).strip() if get("コード") is not None else None),
                       品名=pname, 品名の先頭空白=lead, 品名の記号=pmark,
                       繰越数量=num(get("繰越数量")), 仕入数量=num(get("仕入数量")), 払出数量=num(get("払出数量")),
                       在庫数量=num(get("在庫数量")), 単位=nz(get("単位")),
                       棚番=(str(get("棚番")).strip() if get("棚番") is not None else None), ファイル名=name)
            st = rec["在庫数量"]
            rec["その他の増減"] = None if st is None else round(st - ((rec["繰越数量"] or 0) + (rec["仕入数量"] or 0) - (rec["払出数量"] or 0)), 3)
            out.append(rec)
        by_period[key] = (sortkey, out)
    rows = [r for _, (_, rs) in sorted(by_period.items()) for r in rs]
    cols = ["期間開始", "期間終了", "日数", "製剤コード", "製剤", "コード", "品名", "品名の先頭空白", "品名の記号", "繰越数量", "仕入数量",
            "払出数量", "在庫数量", "その他の増減", "単位", "棚番", "ファイル名"]
    return pd.DataFrame(rows, columns=cols)


# ------------------------------------------------------------------------------
# エリシアSのリスト（経過措置薬品リスト／期限切れリスト／不動品リスト）：⑳㉑⑭
#   見出し 1 行目、データ 2 行目から。フォルダの中で一番新しいファイル 1 本を読む
# ------------------------------------------------------------------------------
def load_system_list(folder, marker):
    """フォルダの中で名前に marker を含む一番新しいファイル（.xls／.xlsx／.csv／.txt）を読む。戻り値：DataFrame（列名は正規化）か None"""
    if not folder or not os.path.isdir(folder):
        return None
    files = [f for f in list_files(folder) if marker in nz(os.path.basename(f))]
    if not files:
        return None
    f = max(files, key=lambda x: (os.path.getmtime(x), os.path.basename(x)))
    rows = first_table(f, "品名")
    hidx = next((i for i, r in enumerate(rows) if "品名" in [norm_col(c) for c in r]), None)
    if hidx is None:
        return None
    hdr = [norm_col(c) for c in rows[hidx]]
    data = [list(r) + [None] * (len(hdr) - len(r)) for r in rows[hidx + 1:] if any(c is not None for c in r)]
    df = pd.DataFrame([r[:len(hdr)] for r in data], columns=hdr)
    df["ファイル名"] = os.path.basename(f)
    names = df["品名"].map(strip_mark)
    df["薬品名"] = names.map(lambda t: t[0])
    df["品名の記号"] = names.map(lambda t: t[1])
    return df[df["薬品名"].notna()].reset_index(drop=True)


def keika_list(folder):
    """経過措置薬品リスト → 薬品名／経過措置期限／在庫数量／単位／記号"""
    df = load_system_list(folder, "経過措置")
    if df is None:
        return pd.DataFrame(columns=["薬品名", "経過措置期限", "在庫数量", "単位", "記号", "ファイル名"])
    out = pd.DataFrame(dict(薬品名=df["薬品名"], 経過措置期限=df["経過措置"].map(to_date), 在庫数量=df["在庫数量"].map(num),
                            単位=df["単位"].map(nz), 記号=df["記号"].map(nz) if "記号" in df else None, ファイル名=df["ファイル名"]))
    return out.drop_duplicates("薬品名").reset_index(drop=True)


def kigen_list(folder):
    """期限切れリスト → 薬品名／ロット／使用期限／数量（期限切れ数量）／単位／仕入先"""
    df = load_system_list(folder, "期限切れ")
    if df is None:
        return pd.DataFrame(columns=["薬品名", "ロット", "使用期限", "数量", "単位", "仕入先", "ファイル名"])
    qty = df["期限切れ数量"] if "期限切れ数量" in df else df.get("数量")
    return pd.DataFrame(dict(薬品名=df["薬品名"], ロット=df["ロット番号"].map(lambda v: None if v is None else str(v).strip()),
                             使用期限=df["期限"].map(to_date), 数量=qty.map(num), 単位=df["単位"].map(nz),
                             仕入先=df["仕入先"].map(nz) if "仕入先" in df else None, ファイル名=df["ファイル名"]))


def fudo_list(folder):
    """不動品リスト → 薬品名／ロット／期限日／最終払出日／不動日数／数量／単位／棚番／薬価／YJコード／JANコード／基準日"""
    df = load_system_list(folder, "不動品")
    cols = ["薬品名", "ロット", "期限日", "最終払出日", "不動日数", "数量", "単位", "棚番", "薬価", "YJコード", "JANコード", "基準日", "ファイル名"]
    if df is None:
        return pd.DataFrame(columns=cols)
    g = lambda c: df[c] if c in df else pd.Series([None] * len(df))
    return pd.DataFrame(dict(薬品名=df["薬品名"], ロット=g("ロット番号").map(lambda v: None if v is None else str(v).strip()),
                             期限日=g("期限日").map(to_date), 最終払出日=g("最終払出日").map(to_date), 不動日数=g("不動日数").map(num),
                             数量=g("数量").map(num), 単位=g("単位").map(nz), 棚番=g("棚番").map(lambda v: None if v is None else str(v).strip()),
                             薬価=g("薬価").map(num), YJコード=g("YJコード").map(lambda v: None if v is None else str(v).strip()),
                             JANコード=g("JANコード").map(lambda v: None if v is None else str(v).strip()), 基準日=g("基準日").map(to_date),
                             ファイル名=df["ファイル名"]))


# ------------------------------------------------------------------------------
# 卸の休み（④）
# ------------------------------------------------------------------------------
def jp_holidays(y):
    def nth_mon(m, n):
        f = dt.date(y, m, 1)
        return f + dt.timedelta(days=(7 - f.weekday()) % 7 + 7 * (n - 1))
    k = y - 1980
    shun = int(20.8431 + 0.242194 * k - k // 4)
    shu = int(23.2488 + 0.242194 * k - k // 4)
    base = {dt.date(y, 1, 1): "元日", nth_mon(1, 2): "成人の日", dt.date(y, 2, 11): "建国記念の日", dt.date(y, 2, 23): "天皇誕生日",
            dt.date(y, 3, shun): "春分の日", dt.date(y, 4, 29): "昭和の日", dt.date(y, 5, 3): "憲法記念日", dt.date(y, 5, 4): "みどりの日",
            dt.date(y, 5, 5): "こどもの日", nth_mon(7, 3): "海の日", dt.date(y, 8, 11): "山の日", nth_mon(9, 3): "敬老の日",
            dt.date(y, 9, shu): "秋分の日", nth_mon(10, 2): "スポーツの日", dt.date(y, 11, 3): "文化の日", dt.date(y, 11, 23): "勤労感謝の日"}
    return base


def holiday_table(lo, hi, extra=None):
    base = {}
    for y in range(lo.year, hi.year + 1):
        base.update(jp_holidays(y))
    hol = dict(base)
    for d, n in base.items():                      # 振替休日
        if d.weekday() == 6:
            x = d + dt.timedelta(days=1)
            while x in base:
                x += dt.timedelta(days=1)
            hol.setdefault(x, "振替休日")
    d = lo
    while d <= hi:                                 # 国民の休日
        if d not in base and (d - dt.timedelta(days=1)) in base and (d + dt.timedelta(days=1)) in base:
            hol.setdefault(d, "国民の休日")
        d += dt.timedelta(days=1)
    for d, n in (extra or {}).items():
        hol[d] = n
    return hol


def calendar(lo, hi, settings):
    extra = {}
    t = settings.tables.get("追加休業日")
    if t is not None and len(t):
        for _, r in t.iterrows():
            d = to_date(r.get("日付"))
            if d:
                extra[d] = (str(r.get("理由")).strip() if r.get("理由") not in (None, "") else "追加休業日")
    hol = holiday_table(lo, hi, extra)
    (s_m, s_d), (e_m, e_d) = settings.ye_start, settings.ye_end

    def year_end(d):
        md, s, e = d.month * 100 + d.day, s_m * 100 + s_d, e_m * 100 + e_d
        return (s <= md <= e) if s <= e else (md >= s or md <= e)
    rows = []
    d = lo
    while d <= hi:
        # ④ fx_卸の休み と同じ順：追加休業日 → 祝日 → 年末年始 → 土曜 → 日曜（土曜も卸は休み）
        why = None
        if d in extra:
            why = extra[d]
        elif d in hol:
            why = hol[d]
        elif year_end(d):
            why = "年末年始"
        elif d.weekday() == 5:
            why = "土曜"
        elif d.weekday() == 6:
            why = "日曜"
        rows.append(dict(日付=d, 曜日=WD[d.weekday()], 卸の営業日=why is None, 休みの理由=why))
        d += dt.timedelta(days=1)
    return pd.DataFrame(rows)


# ------------------------------------------------------------------------------
# データモデル
# ------------------------------------------------------------------------------
REASON_CODES = [
    ("S01", "在庫・不動", "警告", 1, "{値}ヶ月、不動です。（在庫 {在庫}）"),
    ("S02", "在庫・不動", "注意", 2, "在庫が {値}日分あります。（目安 {線}日）"),
    ("S03", "在庫・不動", "欠番", 9, "（欠番）0.3.0 で廃止。包装単位を使わなくなったため"),
    ("S04", "在庫・不動", "情報", 3, "返品・廃棄・移動などで {値} 減っています。"),
    ("D01", "需要", "情報", 3, "需要は安定しています。（平均 {値}日に1回）"),
    ("D02", "需要", "情報", 3, "需要は量がばらつきます。（平均 {値}日に1回）"),
    ("D03", "需要", "情報", 3, "需要は間欠的です。（平均 {値}日に1回）"),
    ("D04", "需要", "注意", 2, "需要はまばらで、出るときは大口です。（平均 {値}日に1回）"),
    ("D05", "需要", "注意", 2, "予定にない払出だけの月が {値}ヶ月あります。（臨時・新患・他科）"),
    # 0.5.0：区分 D（患者が極端に少ない。出納だけの薬は Q01 のまま。切替元は E05 だけ）
    ("D06", "需要", "注意", 2, "予定の患者が月 {値}人しかいません。（線 {線}人）区分 D：在庫は手で管理する薬です"),
    # 0.10.0：傾向（直近 N ヶ月 vs その前 N ヶ月 vs 年間）
    ("D07", "需要", "注意", 2, "増加傾向です。直近の払出は前の期の {値}%、年間の {線}% です。発注量を増やす余地を見てください。"),
    ("D08", "需要", "注意", 2, "減少傾向です。直近の払出は前の期の {値}%、年間の {線}% です。発注量を減らすことを考えてください。"),
    # 0.7.0：薬品グループの 切替先
    ("E05", "期限・制度", "注意", 2, "切替元です。在庫 {在庫} は切替先の需要で {値}日で使い切る見込みです。返品できる分は返品してください。"),
    ("E06", "期限・制度", "注意", 2, "切替先です。切替元の在庫 {値} がなくなるまで、発注は控えめにしてください。"),
    ("P01", "予定とのずれ", "注意", 2, "予定より {値}% 多く出ています。（臨時・飛び処方・再来局）"),
    ("P02", "予定とのずれ", "注意", 2, "予定より {値}% 少なく出ています。（処方終了・残薬調整）"),
    ("P03", "予定とのずれ", "情報", 3, "月ごとに予定と {値}% ずれます。"),
    ("E01", "期限・制度", "警告", 1, "経過措置品です。期限まで {値}日です。"),
    ("E02", "期限・制度", "警告", 1, "経過措置の期限を {値}日過ぎています。"),
    ("E03", "期限・制度", "警告", 1, "発注除外ですが、{値} 払い出されています。"),
    ("Q01", "データ", "情報", 3, "出納だけに出ます。予定に一度も出ていません。"),
    ("Q02", "データ", "注意", 2, "予定だけに出ます。出納と名前がずれている可能性があります。"),
    # 0.2.0 で追加予定：経過措置品がまとまりの中で一番出ている（推奨の位置にいる）
    ("E04", "期限・制度", "警告", 1, "経過措置品が、まとまりの中で一番出ています。次の屋号を推奨に置き換え、処方の切替を進めてください。"),
    # 0.2.0 で追加予定：直近 N ヶ月の払出で見た在庫（需要が落ちた薬を全期間平均が隠すのを防ぐ）
    ("S05", "在庫・不動", "注意", 2, "直近の払出で見ると、在庫が {値}日分あります。（目安 {線}日）"),
    # 0.2.0 で追加予定：エリシアSの不動品リストに載っている（出納の期間より前から動いていない薬も拾える）
    ("S06", "在庫・不動", "注意", 2, "不動品リストに載っています。（{値}日 払出なし、在庫 {在庫}）"),
    # 0.2.0 で追加予定（まとまり＝同じ成分規格の屋号どうし）
    ("G01", "まとまり", "注意", 2, "同じまとまりの中で、構成比が {値}ポイント入れ替わりました。出荷調整の可能性があります。"),
    ("G02", "まとまり", "注意", 2, "仕入がないまま払い出しが続き、在庫が {値} まで減りました。入荷が止まっている可能性があります。"),
]


def fmt0(n):
    return "" if n is None or (isinstance(n, float) and math.isnan(n)) else f"{round(n):,}"


# ------------------------------------------------------------------------------
# 薬品グループ案（㉕）：薬品名を 成分＋剤形＋規格「屋号」 に分解して、需要のまとまりの案を出す（0.7.0）
# ------------------------------------------------------------------------------
FORMS = ["口腔内崩壊錠", "フィルムコーティング錠", "フィルムコート錠", "チュアブル錠", "ザイディス錠", "ザイディス",
         "徐放カプセル", "徐放錠", "腸溶カプセル", "腸溶錠", "糖衣錠", "舌下錠", "バッカル錠",
         "ODフィルム", "OD錠", "D錠", "CR錠", "CRカプセル", "SR錠", "SRカプセル", "LA錠", "LAカプセル", "ER錠", "XR錠", "R錠", "錠",
         "ソフトカプセル", "カプセル", "フィルム", "ドライシロップ", "シロップ", "点眼液", "点鼻液", "吸入液", "外用液", "内用液", "液",
         "細粒", "顆粒", "散", "末", "丸", "トローチ", "テープ", "パップ", "ゲル", "軟膏", "クリーム", "ローション", "坐剤", "注"]
TABLET_FORMS = {"口腔内崩壊錠", "フィルムコーティング錠", "フィルムコート錠", "チュアブル錠", "ザイディス錠", "ザイディス", "糖衣錠",
                "ODフィルム", "OD錠", "D錠", "錠", "ソフトカプセル", "カプセル", "フィルム", "丸"}
POWDER_FORMS = {"散", "顆粒", "細粒", "末"}
LIQUID_FORMS = {"内用液", "シロップ", "液"}
SLOW_FORMS = {"徐放カプセル", "徐放錠", "CR錠", "CRカプセル", "SR錠", "SRカプセル", "LA錠", "LAカプセル", "ER錠", "XR錠", "R錠"}
ENTERIC_FORMS = {"腸溶カプセル", "腸溶錠"}
LETTERS = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")
NUM_CHARS = set("0123456789.,")
SUF_CHARS = LETTERS | set("μ%/単位個万千")
UNIT_LETTERS = LETTERS | set("μ%単個")
UNIT_MAP = [("㎎", "mg"), ("㎍", "μg"), ("㎖", "mL"), ("㎕", "μL"), ("㎏", "kg"), ("㎜", "mm"), ("㎝", "cm"), ("µ", "μ")]
UNIT_FIX = [("ml", "mL"), ("mcg", "μg"), ("ug", "μg"), ("iu", "IU"), ("meq", "mEq")]
CLOSERS = {"]": "[", ")": "(", "）": "（", "】": "【", "〕": "〔"}


def _strip_tail(s):
    s = s.strip()
    while s and s[-1] in CLOSERS:
        p = s.rfind(CLOSERS[s[-1]])
        if p < 0:
            break
        s = s[:p].strip()
    return s


def _canon_unit(u):
    u = u.lower()
    for a_, b_ in UNIT_FIX:
        u = u.replace(a_, b_)
    return u


def _runs(core):
    """名前の中の「数字＋単位」の並び（位置・数・単位・単位あり）。前が数字のもの（7.5 の 5）は拾わない"""
    out = []
    for i, c in enumerate(core):
        if not c.isdigit() or (i > 0 and core[i - 1] in NUM_CHARS):
            continue
        j = i
        while j < len(core) and core[j] in NUM_CHARS:
            j += 1
        k = j
        while k < len(core) and core[k] in SUF_CHARS:
            k += 1
        n = core[i:j].rstrip(".,").replace(",", "")
        u = _canon_unit(core[j:k].rstrip("/."))
        out.append(dict(位置=i, 数=n, 単位=u, 単位あり=any(ch in UNIT_LETTERS for ch in u)))
    return out


def parse_drug_name(name):
    """戻り値：成分案／剤形／規格／屋号／成分規格案／まとまり案／先発後発案／確からしさ（㉕ と同じ決まり。0.10.1）"""
    n0 = nz(name) or ""
    for a_, b_ in UNIT_MAP:
        n0 = n0.replace(a_, b_)
    has_b = "「" in n0 and "」" in n0
    brand = n0[n0.index("「") + 1:n0.index("」")].strip() if has_b else None
    core = _strip_tail(n0[:n0.index("「")] if has_b else n0)
    form = next((f for f in FORMS if f in core), None)
    f_pos = core.index(form) if form else -1
    runs = _runs(core)
    with_u = [r for r in runs if r["単位あり"]]
    after_f = [] if form is None else [r for r in runs if r["位置"] == f_pos + len(form)]
    pick = with_u[0] if with_u else (after_f[0] if after_f else None)
    str0 = None if pick is None else pick["数"] + pick["単位"]
    ok_str = pick is not None
    d_pos = pick["位置"] if pick else -1
    cuts = [x for x in (f_pos, d_pos) if x >= 0]
    ingr = (core if not cuts else core[:min(cuts)]).replace(" ", "").replace("　", "")
    if not ingr or not ok_str or form is None:
        grp = None
    elif form in TABLET_FORMS:
        grp = f"{ingr} {str0}"
    elif form in SLOW_FORMS:
        grp = f"{ingr}徐放 {str0}"
    elif form in ENTERIC_FORMS:
        grp = f"{ingr}腸溶 {str0}"
    elif form in POWDER_FORMS:
        grp = f"{ingr}散剤 {str0}"
    elif form in LIQUID_FORMS:
        grp = f"{ingr}液剤 {str0}"
    else:
        grp = f"{ingr}{form} {str0}"
    spec = None if not ingr else f"{ingr}{form or ''}{str0 if ok_str else ''}"
    if brand is not None and form is not None and ok_str:
        conf = "後発（屋号あり）"
    elif brand is None and form is not None and ok_str:
        conf = "要確認：屋号なし（先発か旧名称。成分名で結ぶ）"
    elif form is None:
        conf = "要確認：剤形が読めない"
    else:
        conf = "要確認：規格が読めない"
    return dict(成分案=(ingr or None), 剤形=form, 規格=(str0 if ok_str else None), 屋号=brand, 成分規格案=spec, まとまり案=grp,
                先発後発案=("後発" if brand is not None else None), 確からしさ=conf)


def propose_groups(m_drug, adj_names=None):
    """薬品グループ案（㉕）。載せるのは、まとめる必要のある薬だけ（0.8.0）：
       同じまとまり（案）が 2 品目以上／経過措置／出荷調整の疑い（adj_names：⑬ の最新月の G01・G02 の薬品名）／
       供給状況が通常以外／切替先あり・切替先に指定／上の薬と同じまとまり（案）。
       先頭 6 列は手入力テーブル 薬品グループ と同じ並び（薬品名・成分規格・需要のまとまり・先発後発・供給状況・切替先）"""
    adj_names = adj_names or set()
    cols = ["薬品名", "成分規格", "需要のまとまり", "先発後発", "供給状況", "切替先", "載せた理由", "確からしさ", "同じまとまりの屋号数",
            "略称", "成分（案）", "剤形", "規格", "屋号", "需要のまとまり（案）", "未登録", "違い"]
    rows = []
    for _, r in m_drug.iterrows():
        p = parse_drug_name(r["薬品名"])
        g = lambda k: (r[k] if k in r.index and isinstance(r[k], str) else None)
        rows.append(dict(薬品名=r["薬品名"], 略称=g("略称"), **p, 登録済みのまとまり=g("需要のまとまり"), 登録_成分規格=g("成分規格"),
                         登録_先発後発=g("先発後発"), 供給状況=g("供給状況"), 切替先=g("切替先"), 切替元=g("切替元"),
                         経過措置=bool(r["経過措置"]) if "経過措置" in r.index else False))
    if not rows:
        return pd.DataFrame(columns=cols)
    df = pd.DataFrame(rows)
    cnt = df["まとまり案"].value_counts()
    df["同じまとまりの屋号数"] = df["まとまり案"].map(lambda g: int(cnt[g]) if isinstance(g, str) else None)
    sq = lambda t: (nz(t) or "").replace(" ", "").replace("　", "")

    def reason(r):
        out = []
        if (r["同じまとまりの屋号数"] or 0) >= 2: out.append(f"同じ成分規格が {int(r['同じまとまりの屋号数'])} 品目")
        if r["経過措置"]: out.append("経過措置")
        if r["薬品名"] in adj_names: out.append("出荷調整の疑い")
        if r["供給状況"] and r["供給状況"] != "通常": out.append(r["供給状況"])
        if r["切替先"] or r["切替元"]: out.append("切替")
        return "、".join(out)
    df["理由1"] = df.apply(reason, axis=1)
    hit = set(df.loc[(df["理由1"] != "") & df["まとまり案"].notna(), "まとまり案"])
    df["載せた理由"] = [a if a else ("同じ成分規格の候補" if isinstance(b, str) and b in hit else "") for a, b in zip(df["理由1"], df["まとまり案"])]
    df = df[df["載せた理由"] != ""].copy()
    df["成分規格"] = [a or b for a, b in zip(df["登録_成分規格"], df["成分規格案"])]
    df["需要のまとまり"] = [a or b for a, b in zip(df["登録済みのまとまり"], df["まとまり案"])]
    df["先発後発"] = [a or b for a, b in zip(df["登録_先発後発"], df["先発後発案"])]
    df["未登録"] = df["登録済みのまとまり"].map(lambda g: "" if isinstance(g, str) else "○")
    df["違い"] = [("○" if isinstance(a, str) and isinstance(b, str) and sq(a) != sq(b) else "") for a, b in zip(df["登録済みのまとまり"], df["まとまり案"])]
    df["並び"] = df["まとまり案"].map(lambda g: 0 if isinstance(g, str) else 1)
    df = df.sort_values(["並び", "まとまり案", "薬品名"], na_position="last")
    df = df.rename(columns={"成分案": "成分（案）", "まとまり案": "需要のまとまり（案）"})
    return df[cols].reset_index(drop=True)


# ------------------------------------------------------------------------------
# 区分（⑫ T_ABC月次）：需要のまとまり単位。未登録の薬は薬品名単体（0.5.0）
#   量の区分：直近 n ヶ月の月平均払出 ≧ 区分Aの線 → A、≧ 区分Bの線 → B、それ未満 → C
#   区分    ：窓の各月（予定データのある月だけ）の、まとまりの予定の延べ患者数の最少が 区分Dの患者数 未満 → D。
#             予定の行がない月は 0 人。窓に予定データの月がなければ 量の区分 のまま
# ------------------------------------------------------------------------------
def abc_table(t_month, grp_map, n, a_line, b_line, d_line):
    """戻り値：T_ABC月次（月・需要のまとまり・払出・構成比・累積・月平均払出・量の区分・患者の最少月・月平均患者数・区分・前月の区分・変更）"""
    cols = ["月", "需要のまとまり", "払出", "構成比", "累積", "月平均払出", "量の区分", "患者の最少月", "月平均患者数", "区分", "前月の区分", "変更"]
    if t_month.empty:
        return pd.DataFrame(columns=cols)
    t = t_month.assign(まとまり=t_month["薬品名"].map(lambda x: grp_map.get(x) or x))
    sui = t[t["出納データあり"]]
    months = sorted(set(sui["月"]))
    pl = t[t["予定データあり"]]
    pat = pl.groupby(["月", "まとまり"])["予定患者数"].sum()
    plan_months = set(pl["月"])
    rows, prev_cls = [], {}
    for i, m in enumerate(months):
        win = months[max(0, i + 1 - n):i + 1]
        w = sui[sui["月"].isin(win)]
        s = w.groupby("まとまり")["払出数量"].sum()
        # その月までに一度でも出たまとまりは、窓に払出がなくても行を作る（払出 0）
        seen = set(sui.loc[sui["月"] <= m, "まとまり"]) | set(pl.loc[pl["月"] <= m, "まとまり"])
        vals = {k: (float(s[k]) if k in s.index and not pd.isna(s[k]) else 0.0) for k in seen}
        items = sorted(vals.items(), key=lambda kv: (-kv[1], kv[0]))
        total = sum(v for _, v in items)
        win_p = [x for x in win if x in plan_months]
        cum, cur_cls = 0.0, {}
        for k, v in items:
            share = (v / total) if total else 0.0
            cum += share
            avg = v / len(win) if win else 0.0
            q = "A" if avg >= a_line else "B" if avg >= b_line else "C"
            if win_p:
                vals = [float(pat.get((pm, k), 0.0) or 0.0) for pm in win_p]
                pmin, pavg = min(vals), sum(vals) / len(win_p)
            else:
                pmin = pavg = None
            cls = "D" if (pmin is not None and pmin < d_line) else q
            before = prev_cls.get(k)
            rows.append(dict(月=m, 需要のまとまり=k, 払出=v, 構成比=share, 累積=(cum if total else 0.0), 月平均払出=avg, 量の区分=q,
                             患者の最少月=pmin, 月平均患者数=pavg, 区分=cls, 前月の区分=before,
                             変更=("" if before in (None, cls) else f"{before}→{cls}")))
            cur_cls[k] = cls
        prev_cls = cur_cls
    return pd.DataFrame(rows, columns=cols)


class Model:
    def __init__(self, settings: Settings):
        self.s = settings
        S = settings
        self.日次データ = load_plan(S.folder, S.start, S.end)
        self.予測データ = load_plan(S.fc_folder, S.fc_start, S.fc_end, keep_zero=True)      # 発注書 F（来週）
        self.出納 = load_suitou(S.sui_folder, S.start, S.end)
        self.経過措置該当 = keika_list(S.keika_folder)
        self.期限切れリスト = kigen_list(S.kigen_folder)
        self.不動品リスト = fudo_list(S.fudo_folder)
        self.T_日次 = self._t_day()
        self.T_月次 = self._t_month()
        self.T_ABC月次 = self._t_abc()
        self.M_日付 = self._m_date()
        self.M_理由コード = pd.DataFrame(REASON_CODES, columns=["コード", "分類", "種類", "優先度", "文言"])
        self.M_薬品 = self._m_drug()
        self.薬品グループ案 = propose_groups(self.M_薬品)
        self.T_ロット = self._t_lot()
        self.T_判定理由 = self._reasons()

    # ⑥
    def _t_day(self):
        d = self.日次データ
        if d.empty:
            return pd.DataFrame(columns=["日付", "薬品名", "予定数量", "予定患者数", "在庫量"])
        g = d.groupby(["日付", "薬品名"], as_index=False).agg(予定数量=("数量", "sum"), 予定患者数=("患者数", "sum"), 在庫量=("在庫量", "max"))
        return g.sort_values(["日付", "薬品名"]).reset_index(drop=True)

    # ⑦
    def _t_month(self):
        S0, E0 = self.s.start, self.s.end
        sui = self.出納.copy()
        sui["月"] = sui["期間開始"].map(month_start)
        sui = sui.sort_values(["月", "品名", "期間終了"])
        suim = sui.groupby(["月", "品名"], as_index=False).agg(
            繰越数量=("繰越数量", "sum"), 仕入数量=("仕入数量", "sum"), 払出数量=("払出数量", "sum"),
            その他の増減=("その他の増減", "sum"), 在庫数量=("在庫数量", "last"))
        sui_months = set(suim["月"])
        day = self.T_日次.copy()
        day["月"] = day["日付"].map(month_start)
        daym = day.groupby(["月", "薬品名"], as_index=False).agg(
            予定数量=("予定数量", "sum"), 予定患者数=("予定患者数", "sum"),
            予定日数=("予定数量", lambda x: int((x.fillna(0) > 0).sum())))
        j = suim.merge(daym, left_on=["月", "品名"], right_on=["月", "薬品名"], how="outer", indicator=True)
        rows = []
        for _, r in j.iterrows():
            m = r["月"]
            on_sui = isinstance(r["品名"], str)
            sui_mon = m in sui_months
            plan_mon = (m >= S0) and (month_end(m) <= E0)
            z = lambda v: (0.0 if (pd.isna(v) and sui_mon) else (None if pd.isna(v) else float(v)))
            out = z(r["払出数量"])
            plan = (0.0 if pd.isna(r["予定数量"]) else float(r["予定数量"])) if plan_mon else None
            gap = round(out - plan, 3) if (sui_mon and plan_mon) else None
            other = z(r["その他の増減"])
            rows.append(dict(月=m, 薬品名=(r["品名"] if on_sui else r["薬品名"]), 出納に載った=on_sui, 出納データあり=sui_mon, 予定データあり=plan_mon,
                             繰越数量=(None if pd.isna(r["繰越数量"]) else float(r["繰越数量"])), 仕入数量=z(r["仕入数量"]), 払出数量=out,
                             その他の増減=other, 在庫数量=(None if pd.isna(r["在庫数量"]) else float(r["在庫数量"])),
                             入庫の概算=((z(r["仕入数量"]) or 0) + max(0, other or 0)) if sui_mon else None,
                             予定数量=plan, 予定患者数=((0.0 if pd.isna(r["予定患者数"]) else float(r["予定患者数"])) if plan_mon else None),
                             予定日数=((0 if pd.isna(r["予定日数"]) else int(r["予定日数"])) if plan_mon else None),
                             予定とのずれ=gap, ずれの割合=((gap / plan) if (gap is not None and plan and plan > 0) else None)))
        t = pd.DataFrame(rows)
        return t.sort_values(["月", "薬品名"]).reset_index(drop=True)

    # ⑫（M_薬品 より先に作る。M_薬品 が区分を参照する）
    def _t_abc(self):
        S = self.s
        grp = S.tables["薬品グループ"]
        grp_map = {}
        if len(grp) and "薬品名" in grp.columns and "需要のまとまり" in grp.columns:
            for _, r in grp.iterrows():
                k, v = nz(r["薬品名"]), nz(r["需要のまとまり"])
                if k and k not in grp_map:
                    grp_map[k] = v
        return abc_table(self.T_月次, grp_map, int(S.num("ABCの月数")), S.num("区分Aの線"), S.num("区分Bの線"), S.num("区分Dの患者数"))

    # ⑧
    def _m_date(self):
        mons = list(self.T_月次["月"]) if len(self.T_月次) else []
        lo = month_start(min([self.s.start] + mons))
        hi = month_end(max([self.s.end] + mons))
        cal = calendar(lo, hi, self.s)
        days = set(self.T_日次["日付"])
        cal["年"] = cal["日付"].map(lambda d: d.year)
        cal["月"] = cal["日付"].map(lambda d: d.month)
        cal["年月"] = cal["日付"].map(lambda d: f"{d.year}-{d.month:02d}")
        cal["月初"] = cal["日付"].map(month_start)
        cal["期"] = cal["日付"].map(lambda d: ["1-3月", "4-6月", "7-9月", "10-12月"][(d.month - 1) // 3])
        cal["曜日番号"] = cal["日付"].map(lambda d: d.weekday() + 1)
        cal["処方のあった日"] = cal["日付"].map(lambda d: d in days)
        return cal[["日付", "年", "月", "年月", "月初", "期", "曜日", "曜日番号", "処方のあった日", "卸の営業日", "休みの理由"]]

    # ⑨
    def _m_drug(self):
        S = self.s
        day, mon, sui = self.T_日次, self.T_月次, self.出納
        # 対象の薬：予定か出納に出た薬 ＋ 不動品リスト・経過措置薬品リストにだけ載る薬（⑨ と同じ。出納の期間より前から動いていない薬を拾う）
        names = sorted(set(day["薬品名"].dropna()) | set(mon["薬品名"].dropna())
                       | set(self.不動品リスト["薬品名"].dropna()) | set(self.経過措置該当["薬品名"].dropna()))
        base = pd.DataFrame({"薬品名": names})
        # 略称・製剤（予定）
        src = self.日次データ[["日付", "薬品名", "略称", "薬品区分"]]
        src = src[src["略称"].notna() | src["薬品区分"].notna()].sort_values("日付")
        abbr = src.groupby("薬品名").agg(略称=("略称", lambda x: nz(x.dropna().iloc[-1]) if x.notna().any() else None),
                                        製剤_予定=("薬品区分", lambda x: x.dropna().iloc[-1] if x.notna().any() else None)).reset_index()
        # 出納の属性（最新期間）
        sl = sui.sort_values("期間終了").groupby("品名").last().reset_index()[["品名", "製剤", "単位", "棚番", "コード", "品名の先頭空白"]]
        # 手入力テーブル
        hand = S.tables
        trans = hand["経過措置リスト"].copy() if len(hand["経過措置リスト"]) else pd.DataFrame(columns=["薬品名", "経過措置期限"])
        trans["薬品名"] = trans["薬品名"].map(nz); trans["経過措置期限"] = trans["経過措置期限"].map(to_date)
        trans = trans.dropna(subset=["薬品名"]).drop_duplicates("薬品名")
        # システムの経過措置薬品リストが優先（手入力は補助）
        sysk = self.経過措置該当[["薬品名", "経過措置期限"]] if len(self.経過措置該当) else pd.DataFrame(columns=["薬品名", "経過措置期限"])
        trans = pd.concat([sysk, trans], ignore_index=True).dropna(subset=["薬品名"]).drop_duplicates("薬品名")
        # 不動品リスト：薬ごとにまとめる（最終払出日は一番新しいもの、不動日数はその日から）
        fd = self.不動品リスト
        fudo = {}
        if len(fd):
            for n, g in fd.groupby("薬品名"):
                last = g["最終払出日"].dropna().max() if g["最終払出日"].notna().any() else None
                idle = g["不動日数"].dropna().min() if g["不動日数"].notna().any() else None
                fudo[n] = dict(最終払出日=last, 不動日数=(None if idle is None else float(idle)), 数量=float(g["数量"].fillna(0).sum()),
                               薬価=(float(g["薬価"].dropna().max()) if g["薬価"].notna().any() else None),
                               YJコード=(g["YJコード"].dropna().iloc[0] if g["YJコード"].notna().any() else None))
        stock = hand["在庫情報"].copy() if len(hand["在庫情報"]) else pd.DataFrame(columns=["薬品名", "列数"])
        stock["薬品名"] = stock["薬品名"].map(nz); stock = stock.dropna(subset=["薬品名"]).drop_duplicates("薬品名")
        excl = set(hand["発注除外"]["薬品名"].map(nz).dropna()) if len(hand["発注除外"]) else set()
        grp = hand["薬品グループ"].copy() if len(hand["薬品グループ"]) else pd.DataFrame(columns=["薬品名", "成分規格", "需要のまとまり", "先発後発", "供給状況", "切替先"])
        for c in ("供給状況", "切替先"):
            if c not in grp.columns:
                grp[c] = None
        for c in grp.columns:
            grp[c] = grp[c].map(nz)
        grp = grp.dropna(subset=["薬品名"]).drop_duplicates("薬品名")
        # 予定の指標
        day_n = day["日付"].nunique()
        plan_rows = {}
        for n, g in day.groupby("薬品名"):
            q = g["予定数量"].dropna(); q = q[q > 0]
            cnt = len(q); tot = float(q.sum()) if cnt else None
            mean = tot / cnt if cnt else None
            sd = float(q.std(ddof=1)) if cnt >= 2 else None
            cv2 = (sd / mean) ** 2 if (sd is not None and mean and mean > 0) else None
            adi = day_n / cnt if cnt else None
            if adi is None:
                typ = None
            elif cv2 is None:
                typ = "安定" if adi < 1.32 else "間欠"
            elif adi < 1.32:
                typ = "安定" if cv2 < 0.49 else "ばらつき大"
            else:
                typ = "間欠" if cv2 < 0.49 else "まばら・大口"
            plan_rows[n] = dict(予定日数=cnt, 予定数量合計=tot, 予定患者数合計=float(g["予定患者数"].sum()),
                                最終予定日=(g.loc[g["予定数量"].fillna(0) > 0, "日付"].max() if cnt else None),
                                平均間隔=adi, **{"ばらつき（CV²）": cv2}, 需要の型=typ, **{"1日平均予定": (tot / day_n if day_n and tot is not None else None)})
        # 区分（⑫ の一番新しい月）と患者の指標（予定データのある月）
        abc = self.T_ABC月次
        latest = abc[abc["月"] == abc["月"].max()] if len(abc) else abc
        abc_map = {r["需要のまとまり"]: r for _, r in latest.iterrows()}
        pl = mon[mon["予定データあり"]]
        plan_months = sorted(set(pl["月"]))
        abc_n = int(S.num("ABCの月数"))
        recent_plan = plan_months[-abc_n:] if plan_months else []
        prev_plan = plan_months[-2 * abc_n:-abc_n] if len(plan_months) >= 2 * abc_n else []
        q0 = S.start
        quarter_of = lambda m: ((m.year - q0.year) * 12 + m.month - q0.month) // 3
        pat_rows = {}
        for n, g in pl.groupby("薬品名"):
            rec = g[g["月"].isin(recent_plan)]
            act = g.loc[g["予定患者数"].fillna(0) > 0, "月"]
            prv = g[g["月"].isin(prev_plan)]
            pat_rows[n] = dict(直近の月平均患者数=(None if not recent_plan else float(rec["予定患者数"].fillna(0).sum()) / len(recent_plan)),
                               前の月平均患者数=(None if not prev_plan else float(prv["予定患者数"].fillna(0).sum()) / len(prev_plan)),
                               年間の月平均患者数=(None if not plan_months else float(g["予定患者数"].fillna(0).sum()) / len(plan_months)),
                               稼働四半期数=len({quarter_of(m) for m in act}))
        pat_default = dict(直近の月平均患者数=(None if not recent_plan else 0.0), 前の月平均患者数=(None if not prev_plan else 0.0),
                           年間の月平均患者数=(None if not plan_months else 0.0), 稼働四半期数=0)
        # 出納の指標
        sm = mon[mon["出納データあり"]]
        sui_months = sorted(set(sm["月"]))
        mon_n = len(sui_months)
        cal_days = sum(month_end(m).day for m in sui_months)
        latest_m = max(sui_months) if sui_months else None
        recent_n = int(S.num("直近の月数"))
        recent_months = sui_months[-recent_n:] if sui_months else []
        recent_days = sum(month_end(m).day for m in recent_months)
        prev_months = sui_months[-2 * recent_n:-recent_n] if len(sui_months) >= 2 * recent_n else []
        prev_days = sum(month_end(m).day for m in prev_months)
        sui_rows = {}
        for n, g in sm.groupby("薬品名"):
            moved = g[g["出納に載った"]]
            out = float(g["払出数量"].fillna(0).sum())
            per_day = out / cal_days if cal_days else None
            rec_out = float(g.loc[g["月"].isin(recent_months), "払出数量"].fillna(0).sum())
            rec_per_day = rec_out / recent_days if recent_days else None
            prv_out = float(g.loc[g["月"].isin(prev_months), "払出数量"].fillna(0).sum())
            rec_both = g[g["月"].isin(recent_months) & g["予定とのずれ"].notna()]
            rec_plan = float(rec_both["予定数量"].fillna(0).sum())
            last = moved.sort_values("月").iloc[-1] if len(moved) else None
            both = g[g["予定とのずれ"].notna()]
            plan_sum = float(both["予定数量"].fillna(0).sum())
            gap_sum = float(both["予定とのずれ"].sum()); abs_sum = float(both["予定とのずれ"].abs().sum())
            stock_v = (None if last is None or pd.isna(last["在庫数量"]) else float(last["在庫数量"]))
            sui_rows[n] = dict(動いた月数=int(len(moved)), 払出合計=out, 払出のあった月数=int((g["払出数量"].fillna(0) > 0).sum()),
                               月平均払出=(out / mon_n if mon_n else None), **{"1日平均払出": per_day},
                               仕入合計=float(g["仕入数量"].fillna(0).sum()), 入庫の概算合計=float(g["入庫の概算"].fillna(0).sum()),
                               その他の減少=float((-g["その他の増減"].fillna(0)).clip(lower=0).sum()),
                               最後に動いた月=(last["月"] if last is not None else None), 最終在庫=stock_v,
                               動きのない月数=(months_between(last["月"], latest_m) if last is not None else None),
                               在庫日数=((stock_v / per_day) if (stock_v is not None and per_day) else None),
                               比較月数=int(len(both)), 比較の予定合計=plan_sum, 比較のずれ合計=gap_sum, ずれの割合=((gap_sum / plan_sum) if plan_sum > 0 else None),
                               ずれの大きさ=((abs_sum / plan_sum) if plan_sum > 0 else None),
                               予定外の月数=int(((both["予定数量"].fillna(0) == 0) & (both["払出数量"].fillna(0) > 0)).sum()),
                               直近の払出合計=rec_out, 直近の1日平均払出=rec_per_day, 前の払出合計=prv_out,
                               直近のずれの割合=((float(rec_both["予定とのずれ"].sum()) / rec_plan) if rec_plan > 0 else None),
                               直近の在庫日数=((stock_v / rec_per_day) if (stock_v is not None and rec_per_day) else None))
        plan_cols = ["予定日数", "予定数量合計", "予定患者数合計", "最終予定日", "平均間隔", "ばらつき（CV²）", "需要の型", "1日平均予定"]
        sui_cols = ["動いた月数", "払出合計", "払出のあった月数", "月平均払出", "1日平均払出", "仕入合計", "入庫の概算合計", "その他の減少",
                    "最後に動いた月", "最終在庫", "動きのない月数", "在庫日数", "比較月数", "比較の予定合計", "比較のずれ合計", "ずれの割合", "ずれの大きさ", "予定外の月数",
                    "直近の払出合計", "直近の1日平均払出", "直近の在庫日数", "前の払出合計", "直近のずれの割合"]
        recs = []
        tr_map = dict(zip(trans["薬品名"], trans["経過措置期限"]))
        st_map = {r["薬品名"]: num(r.get("列数")) for _, r in stock.iterrows()}
        gr_map = {r["薬品名"]: r for _, r in grp.iterrows()}
        # 切替（0.7.0）：切替元（この薬を切替先にしている薬）と、まとまりの直近1日平均払出
        sw_from, sw_stock = {}, {}
        for _, r0 in grp.iterrows():
            if isinstance(r0["切替先"], str):
                sw_from.setdefault(r0["切替先"], []).append(r0["薬品名"])
        for k, lst in sw_from.items():
            sw_stock[k] = float(sum((sui_rows.get(x, {}).get("最終在庫") or 0) for x in lst))
        grp_rate = {}
        for n0, r0 in gr_map.items():
            g0 = r0["需要のまとまり"]
            if isinstance(g0, str):
                grp_rate[g0] = grp_rate.get(g0, 0.0) + float(sui_rows.get(n0, {}).get("直近の1日平均払出") or 0)
        ab_map = {r["薬品名"]: r for _, r in abbr.iterrows()}
        sl_map = {r["品名"]: r for _, r in sl.iterrows()}
        for n in names:
            a, s_, p, m = ab_map.get(n), sl_map.get(n), plan_rows.get(n, {}), sui_rows.get(n, {})
            cols_n = st_map.get(n)
            r = dict(薬品名=n, 略称=(a["略称"] if a is not None else None),
                     製剤=((s_["製剤"] if s_ is not None and s_["製剤"] else None) or (a["製剤_予定"] if a is not None else None)),
                     単位=(s_["単位"] if s_ is not None else None), 棚番=(s_["棚番"] if s_ is not None else None),
                     コード=(s_["コード"] if s_ is not None else None), 品名の先頭空白=(int(s_["品名の先頭空白"]) if s_ is not None else None),
                     経過措置=(n in tr_map), 経過措置期限=tr_map.get(n), 発注除外=(n in excl),
                     不動品リスト=(n in fudo), システム不動日数=(fudo[n]["不動日数"] if n in fudo else None),
                     システム最終払出日=(fudo[n]["最終払出日"] if n in fudo else None), システム在庫=(fudo[n]["数量"] if n in fudo else None),
                     薬価=(fudo[n]["薬価"] if n in fudo else None), YJコード=(fudo[n]["YJコード"] if n in fudo else None),
                     成分規格=(gr_map[n]["成分規格"] if n in gr_map else None), 需要のまとまり=(gr_map[n]["需要のまとまり"] if n in gr_map else None),
                     先発後発=(gr_map[n]["先発後発"] if n in gr_map else None), 供給状況=(gr_map[n]["供給状況"] if n in gr_map else None),
                     切替先=(gr_map[n]["切替先"] if n in gr_map and isinstance(gr_map[n]["切替先"], str) else None),
                     切替元=("、".join(sorted(sw_from[n])) if n in sw_from else None), 切替元の在庫=sw_stock.get(n),
                     列数=cols_n,
                     出どころ=("予定・出納" if (p.get("予定日数", 0) or 0) > 0 and (m.get("動いた月数", 0) or 0) > 0
                               else "予定だけ" if (p.get("予定日数", 0) or 0) > 0 else "出納だけ"))
            key = (gr_map[n]["需要のまとまり"] if n in gr_map and isinstance(gr_map[n]["需要のまとまり"], str) else None) or n
            ab = abc_map.get(key)
            r["区分"] = "D" if r["切替先"] else (ab["区分"] if ab is not None else None)        # 切替元は D
            own_rate = float(m.get("直近の1日平均払出") or 0)
            r["まとまりの直近1日平均払出"] = grp_rate.get(key, own_rate) if key != n else own_rate
            st0 = float(m.get("最終在庫") or 0)
            r["切替の使い切り日数"] = (st0 / r["まとまりの直近1日平均払出"]) if (r["切替先"] and r["まとまりの直近1日平均払出"] > 0) else None
            r["量の区分"] = ab["量の区分"] if ab is not None else None
            r["患者の最少月"] = (None if ab is None or pd.isna(ab["患者の最少月"]) else float(ab["患者の最少月"]))
            for c in plan_cols:
                r[c] = p.get(c)
            for c, v in (pat_rows.get(n) or pat_default).items():
                r[c] = v
            for c in sui_cols:
                r[c] = m.get(c)
            recs.append(r)
        # 発注の収束（0.9.0）：まとまりの品目数・発注先・まとまりのずれの割合（⑨ と同じ決まり）
        by_g = {}
        for r in recs:
            if isinstance(r["需要のまとまり"], str):
                by_g.setdefault(r["需要のまとまり"], []).append(r)

        def pick(rows):
            if not rows:
                return None
            return sorted(rows, key=lambda x: (-(x.get("直近の払出合計") or 0), x["薬品名"]))[0]["薬品名"]
        for r in recs:
            g0 = r["需要のまとまり"] if isinstance(r["需要のまとまり"], str) else None
            rows = by_g[g0] if g0 else [r]
            r["まとまりの品目数"] = len(rows)
            tgt = None
            if len(rows) >= 2:
                tgt = (pick([x for x in rows if x["切替元"]])
                       or pick([x for x in rows if not x["切替先"] and (not isinstance(x["供給状況"], str) or x["供給状況"] == "通常")
                                and not x["発注除外"] and not x["経過措置"]])
                       or pick([x for x in rows if not x["切替先"]]))
            r["発注先"] = tgt
            if g0:
                w = sum((x.get("比較の予定合計") or 0) for x in rows)
                r["まとまりのずれの割合"] = (sum((x.get("比較のずれ合計") or 0) for x in rows) / w) if w > 0 else None
            else:
                r["まとまりのずれの割合"] = r.get("ずれの割合")
        # 傾向（0.10.0）：直近 N ヶ月 vs その前 N ヶ月 vs 年間（⑨ と同じ決まり）
        trend_line, tiny, d_line = S.num("傾向の線"), S.num("極小需要の月の払出"), S.num("区分Dの患者数")
        enough = len(prev_months) == recent_n and recent_days > 0

        def judge(r1, r2, rp, pre):
            if pre is not None:
                return pre
            if r1 is None or r2 is None:
                return "増加" if rp > 0 else "横ばい"
            if r1 >= 1 + trend_line and r2 >= 1 + trend_line:
                return "増加"
            if r1 <= 1 - trend_line and r2 <= 1 - trend_line:
                return "減少"
            return "横ばい"

        def trend_of(rec, prv, al, avg_m):
            rp = rec / recent_days if recent_days else 0.0
            pp = prv / prev_days if prev_days else 0.0
            ap = al / cal_days if cal_days else 0.0
            r1 = rp / pp if pp > 0 else None
            r2 = rp / ap if ap > 0 else None
            pre = "月不足" if not enough else "少量" if (avg_m or 0) < tiny else None
            return r1, r2, judge(r1, r2, rp, pre)

        def pat_trend(rec, prv, al):
            r1 = rec / prv if (prv is not None and prv > 0 and rec is not None) else None
            r2 = rec / al if (al is not None and al > 0 and rec is not None) else None
            pre = "月不足" if (prv is None or al is None or rec is None) else "少量" if al < d_line else None
            return r1, r2, judge(r1, r2, (rec or 0), pre)
        for r in recs:
            g0 = r["需要のまとまり"] if isinstance(r["需要のまとまり"], str) else None
            rows = by_g[g0] if g0 else [r]
            z = lambda k: sum((x.get(k) or 0) for x in rows)
            r["前期比"], r["年間比"], r["傾向"] = trend_of(r.get("直近の払出合計") or 0, r.get("前の払出合計") or 0, r.get("払出合計") or 0, r.get("月平均払出"))
            r["まとまりの前期比"], r["まとまりの年間比"], r["まとまりの傾向"] = trend_of(z("直近の払出合計"), z("前の払出合計"), z("払出合計"), z("月平均払出"))
            r["患者の前期比"], r["患者の年間比"], r["患者の傾向"] = pat_trend(r.get("直近の月平均患者数"), r.get("前の月平均患者数"), r.get("年間の月平均患者数"))
        df = pd.DataFrame(recs)
        return df.sort_values(["払出合計", "薬品名"], ascending=[False, True], na_position="last").reset_index(drop=True)

    # ⑭ T_ロット：期限切れリストと不動品リストのロットを合わせ、期限の早い順に直近の払出ペースで割り当てる
    def _t_lot(self):
        base = self.s.base_date
        a = self.期限切れリスト.assign(出どころ="期限切れリスト")[["薬品名", "ロット", "使用期限", "数量", "単位", "出どころ", "ファイル名"]] if len(self.期限切れリスト) else None
        b = self.不動品リスト.rename(columns={"期限日": "使用期限"}).assign(出どころ="不動品リスト")[["薬品名", "ロット", "使用期限", "数量", "単位", "出どころ", "ファイル名"]] if len(self.不動品リスト) else None
        parts = [x for x in (a, b) if x is not None and len(x)]
        cols = ["薬品名", "ロット", "使用期限", "数量", "単位", "出どころ", "残り日数", "それより早いロットの数量", "期限までに使い切れる量", "返す量", "ファイル名"]
        if not parts:
            return pd.DataFrame(columns=cols)
        t = pd.concat(parts, ignore_index=True)
        t = t[t["数量"].fillna(0) > 0].drop_duplicates(["薬品名", "ロット"], keep="first")
        pd_map = dict(zip(self.M_薬品["薬品名"], self.M_薬品["直近の1日平均払出"]))
        rows = []
        for n, g in t.groupby("薬品名"):
            g = g.sort_values(["使用期限", "ロット"], na_position="last")
            per_day = pd_map.get(n); per_day = 0.0 if per_day is None or pd.isna(per_day) else float(per_day)
            prior = 0.0
            for _, r in g.iterrows():
                q = float(r["数量"])
                left = (r["使用期限"] - base).days if r["使用期限"] else None
                if left is None:
                    usable = q
                else:
                    cap = per_day * max(0, left)
                    usable = max(0.0, min(q, cap - prior))
                rows.append(dict(薬品名=n, ロット=r["ロット"], 使用期限=r["使用期限"], 数量=q, 単位=r["単位"], 出どころ=r["出どころ"], 残り日数=left,
                                 それより早いロットの数量=prior, 期限までに使い切れる量=usable, 返す量=q - usable, ファイル名=r["ファイル名"]))
                prior += q
        return pd.DataFrame(rows, columns=cols)

    # ⑪
    def _reasons(self):
        S = self.s
        still_m, over_d = S.num("不動とみなす月数"), S.num("在庫過多とみなす日数")
        gap_r, gap_s, base = S.num("ずれの注意割合"), S.num("ずれの大きさの注意割合"), S.base_date
        d_line = S.num("区分Dの患者数")
        codes = {c[0]: c for c in REASON_CODES}
        # まとまりの中で直近の払出が一番多い薬（経過措置品がここにいれば E04）
        md = self.M_薬品
        top_in_group = set()
        for k, g in md[md["需要のまとまり"].notna()].groupby("需要のまとまり"):
            if len(g) >= 2:
                top_in_group.add(g.sort_values("直近の払出合計", ascending=False).iloc[0]["薬品名"])
        rows = []
        for _, r in self.M_薬品.iterrows():
            g = lambda k: (None if pd.isna(r[k]) else r[k])
            stock = g("最終在庫") or 0; still = g("動きのない月数"); days = g("在庫日数")
            less = g("その他の減少") or 0; adi = g("平均間隔"); typ = g("需要の型"); gr = g("ずれの割合"); gs = g("ずれの大きさ")
            trans = g("経過措置期限"); left = (trans - base).days if trans else None; out = g("払出合計") or 0
            hits = []
            if stock > 0 and still is not None and still >= still_m: hits.append(("S01", still, still_m))
            if days is not None and days >= over_d: hits.append(("S02", days, over_d))
            rdays = g("直近の在庫日数")
            if rdays is not None and rdays >= over_d: hits.append(("S05", rdays, over_d))
            if less > 0: hits.append(("S04", less, None))
            if r["不動品リスト"] is True: hits.append(("S06", g("システム不動日数"), None))
            dmap = {"安定": "D01", "ばらつき大": "D02", "間欠": "D03", "まばら・大口": "D04"}
            if typ in dmap: hits.append((dmap[typ], adi, None))
            if (g("予定外の月数") or 0) >= 1: hits.append(("D05", g("予定外の月数"), 1))
            if g("区分") == "D" and r["出どころ"] != "出納だけ" and not g("切替先"): hits.append(("D06", g("患者の最少月"), d_line))
            pct = lambda x: None if x is None else x * 100
            if g("傾向") == "増加": hits.append(("D07", pct(g("前期比")), pct(g("年間比"))))
            if g("傾向") == "減少": hits.append(("D08", pct(g("前期比")), pct(g("年間比"))))
            if g("切替先"): hits.append(("E05", g("切替の使い切り日数"), None))
            if g("切替元"): hits.append(("E06", g("切替元の在庫"), None))
            if gr is not None and gr >= gap_r: hits.append(("P01", gr * 100, gap_r * 100))
            if gr is not None and gr <= -gap_r: hits.append(("P02", -gr * 100, gap_r * 100))
            if gs is not None and gs >= gap_s: hits.append(("P03", gs * 100, gap_s * 100))
            if left is not None and left >= 0: hits.append(("E01", left, None))
            if left is not None and left < 0: hits.append(("E02", -left, None))
            if r["発注除外"] and out > 0: hits.append(("E03", out, None))
            if trans is not None and r["薬品名"] in top_in_group: hits.append(("E04", None, None))
            if r["出どころ"] == "出納だけ": hits.append(("Q01", None, None))
            if r["出どころ"] == "予定だけ": hits.append(("Q02", None, None))
            for code, val, line in hits:
                c = codes[code]
                msg = c[4].replace("{値}", fmt0(val)).replace("{線}", fmt0(line)).replace("{在庫}", fmt0(g("最終在庫")))
                rows.append(dict(薬品名=r["薬品名"], コード=code, 分類=c[1], 種類=c[2], 優先度=c[3], 文言=msg, 値=val, 線=line))
        t = pd.DataFrame(rows, columns=["薬品名", "コード", "分類", "種類", "優先度", "文言", "値", "線"])
        return t.sort_values(["優先度", "コード", "薬品名"]).reset_index(drop=True)
