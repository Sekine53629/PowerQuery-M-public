# -*- coding: utf-8 -*-
"""src/_data/先発成分.csv → src/26_先発成分.m（㉖ 先発名 → 成分名 の対応表。接続のみ）を作る
使い方（プロジェクトのフォルダで）：python tools/make_brand_table.py
  CSV は UTF-8（BOM なし）、列 先発名・成分名・分類。空行と # で始まる行は飛ばす。
  同じ先発名が 2 行あればエラー。作ったあとは python build.py で dist を組み立てる
"""
import csv
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CSV = os.path.join(ROOT, "src", "_data", "先発成分.csv")
OUT = os.path.join(ROOT, "src", "26_先発成分.m")

HEAD = """// ================================================================================
// ㉖ クエリ名：先発成分（接続のみ。シートにもデータモデルにも読み込まない）
//    先発名 → 成分名（後発品の名前の付け方に合わせた成分名）の対応表。主な精神科の薬を中心に、店で扱いの多い薬も入れている（0.11.0）
//    ・㉕ 薬品グループ案 が、屋号「　」のない薬（先発品）の成分名をここで引き、後発品と同じ「成分 規格」のまとまりの案を出すのに使う
//      （名前の先頭が 先発名 と一致したら引く。長い名前から順に照合するので「アタラックスP」は「アタラックス」より先に当たる）
//    ・ブックに手入力テーブル「先発成分」（列 先発名・成分名）を作れば、㉕ はそちらを優先して足す・上書きする（このクエリは直さなくてよい）
//    ・このクエリは src/_data/先発成分.csv から tools/make_brand_table.py で作る。手で直さず、CSV を直して作り直す
//    ・成分名は後発品の名前の付け方（塩の表記を含む。例：ゾルピデム酒石酸塩、ロキソプロフェンNa）。表記が店の後発品と違えば
//      同じまとまりにならないので、その薬は手入力テーブル「先発成分」か「薬品グループ」で合わせる
// ================================================================================
let
    Source = #table(type table [先発名 = text, 成分名 = text, 分類 = text], {
"""
TAIL = """    })
in
    Source
"""


def q(s):
    return '"' + s.replace('"', '""') + '"'


def main():
    rows = []
    seen = {}
    with open(CSV, encoding="utf-8-sig", newline="") as f:
        for i, r in enumerate(csv.DictReader(f), start=2):
            b = (r.get("先発名") or "").strip()
            g = (r.get("成分名") or "").strip()
            c = (r.get("分類") or "").strip()
            if not b or b.startswith("#"):
                continue
            if not g:
                raise SystemExit(f"{CSV} 行 {i}: 成分名が空です（{b}）")
            if b in seen:
                raise SystemExit(f"{CSV} 行 {i}: 先発名 {b} が重複しています（行 {seen[b]}）")
            seen[b] = i
            rows.append((b, g, c))
    body = ",\n".join(f"        {{{q(b)}, {q(g)}, {q(c)}}}" for b, g, c in rows) + "\n"
    with open(OUT, "w", encoding="utf-8", newline="\n") as f:
        f.write(HEAD + body + TAIL)
    print(f"{OUT}: {len(rows)} 件")


if __name__ == "__main__":
    main()
