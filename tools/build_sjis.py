#!/usr/bin/env python3
"""src/ (UTF-8) から dist/ (Shift-JIS) を生成する。

VBE の［ファイル］→［ファイルのインポート］は Shift-JIS を前提に読み込むため、
GitHub 上で読みやすい UTF-8 版とは別に、インポート用の Shift-JIS 版を用意している。

    python3 tools/build_sjis.py          # 生成
    python3 tools/build_sjis.py --check  # 差分があれば exit 1 (CI 用)
"""
import io, sys, pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
JOBS = [
    ("src/ShiftClick.bas",   "dist/ShiftClick_sjis.bas",   'Attribute VB_Name = "ShiftClick"\r\n'),
    ("src/SheetModule.txt",  "dist/SheetModule_sjis.txt",  ""),
]

def build(check=False):
    ng = []
    for src, dst, header in JOBS:
        text = header + io.open(ROOT / src, encoding="utf-8", newline="").read()
        try:
            data = text.encode("cp932")
        except UnicodeEncodeError as e:
            print(f"[NG] {src}: Shift-JIS にできない文字があります -> {e}")
            return 1
        out = ROOT / dst
        if check:
            if not out.exists() or out.read_bytes() != data:
                ng.append(dst)
        else:
            out.write_bytes(data)
            print(f"[OK] {dst}")
    if check:
        if ng:
            print("[NG] dist が古いです: " + ", ".join(ng))
            print("     python3 tools/build_sjis.py を実行してコミットしてください。")
            return 1
        print("[OK] dist は最新です")
    return 0

if __name__ == "__main__":
    sys.exit(build(check="--check" in sys.argv))
