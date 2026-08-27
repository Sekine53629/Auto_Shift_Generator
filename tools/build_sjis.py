#!/usr/bin/env python3
"""src/ (UTF-8) から dist/ (Shift-JIS) を生成する。

VBE の［ファイル］→［ファイルのインポート］は Shift-JIS を前提に読み込むため、
GitHub 上で読みやすい UTF-8 版とは別に、インポート用の Shift-JIS 版を用意している。

対象は src/*.bas を自動で拾う。標準モジュールには Attribute VB_Name を付けて
そのままインポートできる形にし、シートモジュール (Sheet1.bas) は貼り付け用の
テキストとして出力する。

    python3 tools/build_sjis.py          # 生成
    python3 tools/build_sjis.py --check  # 差分があれば exit 1 (CI 用)
"""
import io
import sys
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
SRC_DIR = ROOT / "src"
DIST_DIR = ROOT / "dist"

# シートモジュールは標準モジュールとしてインポートできないため、
# Attribute を付けず貼り付け用のテキストとして出力する。
SHEET_MODULES = {"Sheet1"}

CRLF = "\r\n"
LF = "\n"
CR = "\r"


def to_crlf(text):
    """改行を CRLF に揃える。

    作業ツリーの改行は環境や .gitattributes の適用状況で LF にも CRLF にも
    なりうる。出力がそれに左右されると --check がすり抜けるため、ここで
    正規化する。VBE は CRLF を前提に読み込むので、これが正しい出力でもある。
    """
    return text.replace(CRLF, LF).replace(CR, LF).replace(LF, CRLF)


def jobs():
    """(srcパス, distパス, 先頭に足すヘッダー) を列挙する。"""
    out = []
    for src in sorted(SRC_DIR.glob("*.bas")):
        name = src.stem
        if name in SHEET_MODULES:
            out.append((src, DIST_DIR / f"{name}_sjis.txt", ""))
        else:
            out.append((src, DIST_DIR / f"{name}_sjis.bas",
                        f'Attribute VB_Name = "{name}"' + CRLF))
    return out


def stale_files(expected):
    """dist に残っている、もう生成されないファイル。"""
    if not DIST_DIR.exists():
        return []
    keep = {p.name for p in expected}
    return sorted(p.name for p in DIST_DIR.iterdir()
                  if p.is_file() and p.name not in keep)


def build(check=False):
    todo = jobs()
    if not todo:
        print("[NG] src/*.bas が見つかりません")
        return 1

    ng = []
    for src, dst, header in todo:
        text = to_crlf(header + io.open(src, encoding="utf-8", newline="").read())
        try:
            data = text.encode("cp932")
        except UnicodeEncodeError as e:
            print(f"[NG] {src.name}: Shift-JIS にできない文字があります -> {e}")
            return 1
        if check:
            if not dst.exists() or dst.read_bytes() != data:
                ng.append(dst.name)
        else:
            DIST_DIR.mkdir(exist_ok=True)
            dst.write_bytes(data)
            print(f"[OK] dist/{dst.name}")

    extra = stale_files([d for _, d, _ in todo])
    if check:
        if ng or extra:
            if ng:
                print("[NG] dist が古いです: " + ", ".join(ng))
            if extra:
                print("[NG] dist に不要なファイルがあります: " + ", ".join(extra))
            print("     python3 tools/build_sjis.py を実行してコミットしてください。")
            return 1
        print(f"[OK] dist は最新です ({len(todo)} ファイル)")
    else:
        for name in extra:
            (DIST_DIR / name).unlink()
            print(f"[--] dist/{name} を削除しました (src に対応するファイルがありません)")
    return 0


if __name__ == "__main__":
    sys.exit(build(check="--check" in sys.argv))
