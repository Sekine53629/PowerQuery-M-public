"""PowerPivot_ABC を組み立てる
使い方（このフォルダで）：python build.py
  1. manifest.txt の順に _header.txt と各クエリ（.m）を連結する
  2. dist/PowerPivot_ABC_<版>.txt に UTF-8・CRLF で書き出す（Excel の詳細エディターに貼る用）
  3. ../_tools/mcheck.py で括弧と束縛をチェックする（問題があれば終了コード 1）
"""
import os, sys, subprocess

HERE = os.path.dirname(os.path.abspath(__file__))


def read(path):
    with open(os.path.join(HERE, path), encoding="utf-8") as f:
        return f.read().replace("\r\n", "\n")


def main():
    conf, parts = {}, []
    for line in read("manifest.txt").split("\n"):
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
    chk = os.path.join(HERE, "..", "_tools", "mcheck.py")
    r = subprocess.run([sys.executable, chk, out])
    sys.exit(r.returncode)


if __name__ == "__main__":
    main()
