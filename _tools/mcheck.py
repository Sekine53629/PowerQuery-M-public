"""M 言語の簡易チェック（括弧の対応と、let のトップレベルの束縛の区切り）
使い方：python mcheck.py <組み立て済みの .txt>
  ・「// ===」の次の行が「// ◯ クエリ名：…」の見出しで、クエリを区切る
  ・文字列とコメントを除いてから、括弧の対応と let の束縛ごとの「名前 =」を数える
  ・意味（列名・型・関数名の誤り）は見ない。Excel での動作確認の代わりにはならない
  ・問題があれば終了コード 1
"""
import re, sys
src = open(sys.argv[1], encoding="utf-8").read().replace("\r\n", "\n")
parts = re.split(r"// =+\n// (.+?クエリ名：[^\n]*)\n", src)
def strip(code):
    out, i, n = [], 0, len(code)
    while i < n:
        c = code[i]
        if code.startswith("//", i):
            j = code.find("\n", i); i = n if j < 0 else j; continue
        if code.startswith("/*", i):
            j = code.find("*/", i); i = j + 2; continue
        if c == '"':
            j = i + 1
            while True:
                k = code.find('"', j)
                if k + 1 < n and code[k + 1] == '"': j = k + 2; continue
                break
            out.append('"S"'); i = k + 1; continue
        out.append(c); i += 1
    return "".join(out)
ident = r'(#"S"|[^\W\d][\w]*)'
bad = 0
for k in range(1, len(parts), 2):
    title = parts[k]; code = parts[k + 1]
    # drop trailing comment-only blocks
    code = strip(code)
    st = []
    pairs = {")": "(", "]": "[", "}": "{"}
    for i, ch in enumerate(code):
        if ch in "([{": st.append((ch, i))
        elif ch in ")]}":
            if not st or st[-1][0] != pairs[ch]:
                print("UNBALANCED", title, code[max(0, i-80):i+5]); bad += 1; break
            st.pop()
    if st: print("UNCLOSED", title, st[-1]); bad += 1
    # top-level let bindings: depth-0 commas
    m = re.search(r"\blet\b", code)
    if not m: print("no let", title); continue
    depth, segs, cur, i = 0, [], [], m.end()
    # find matching top-level 'in' : track nested let/in at depth 0
    nest = 0
    body_end = None
    toks = re.finditer(r"\blet\b|\bin\b|[()\[\]{}]|,", code[m.end():])
    last = m.end(); segs = []
    for t in toks:
        g = t.group(); pos = m.end() + t.start()
        if g in "([{": depth += 1
        elif g in ")]}": depth -= 1
        elif depth == 0 and g == "let": nest += 1
        elif depth == 0 and g == "in":
            if nest == 0: segs.append(code[last:pos]); body_end = pos; break
            nest -= 1
        elif depth == 0 and nest == 0 and g == ",":
            segs.append(code[last:pos]); last = pos + 1
    for sgm in segs:
        lines = [l for l in sgm.split("\n") if l.strip()]
        # binding start lines at depth 0 within segment
        d = 0; starts = 0
        pre = re.split(r"\blet\b", sgm)[0]
        for l in pre.split("\n"):
            if d == 0 and re.match(r"\s*" + ident + r"\s*=(?![=>])", l): starts += 1
            for ch in l:
                if ch in "([{": d += 1
                elif ch in ")]}": d -= 1
        if starts != 1:
            print("BINDING?", title, "|", sgm.strip()[:120].replace("\n", " ")); bad += 1
    print(f"{title[:40]:<42} bindings={len(segs)}")
print("problems:", bad)
sys.exit(1 if bad else 0)
