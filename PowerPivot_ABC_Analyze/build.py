"""PowerPivot_ABC を組み立てる
使い方（このフォルダで）：python build.py [manifest ...]
  引数なしのときは manifest.txt（本体 ⓪～⑪）と manifest_init.txt（初期テーブル A～F）を順に組み立てる
  1. manifest の順に HEADER と各クエリ（.m）を連結する
  2. OUTPUT（dist/…_<版>.txt）に UTF-8・CRLF で書き出す（Excel の詳細エディターに貼る用）
  3. ../_tools/mcheck.py で括弧と束縛をチェックする（問題があれば終了コード 1）
"""
import os, sys, subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_MANIFESTS = ("manifest.txt", "manifest_init.txt")
CHECKER = os.path.join(HERE, "..", "_tools", "mcheck.py")


def read(path):
    with open(os.path.join(HERE, path), encoding="utf-8") as f:
        return f.read().replace("\r\n", "\n")


def build(manifest):
    """manifest を読み、組み立てた txt の絶対パスを返す"""
    conf, parts = {}, []
    for line in read(manifest).split("\n"):
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if "=" in line and not line.endswith(".m"):
            k, v = line.split("=", 1)
            conf[k.strip()] = v.strip()
        else:
            parts.append(line)
    ver = conf["VERSION"]
    out_rel = conf["OUTPUT"].replace("{VERSION}", ver)
    text = read(conf["HEADER"]).rstrip("\n") + "\n\n\n"
    for p in parts:
        text += read(p).rstrip("\n") + "\n\n\n"
    out = os.path.join(HERE, out_rel)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with open(out, "w", encoding="utf-8", newline="") as f:
        f.write(text.rstrip("\n").replace("\n", "\r\n") + "\r\n")
    print("書き出し:", out_rel, f"（{len(parts)} クエリ）")
    return out


def main():
    # Windows の cp932 コンソールでも ⓪ などを表示できるようにする
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    manifests = sys.argv[1:] or [m for m in DEFAULT_MANIFESTS if os.path.exists(os.path.join(HERE, m))]
    rc = 0
    for m in manifests:
        out = build(m)
        rc |= subprocess.run([sys.executable, CHECKER, out]).returncode
    sys.exit(rc)


if __name__ == "__main__":
    main()
