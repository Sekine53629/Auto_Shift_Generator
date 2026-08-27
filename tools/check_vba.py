#!/usr/bin/env python3
"""src/*.bas の静的チェック。VBE でコンパイルする前に落ちる類のミスを拾う。

このプロジェクトは Excel が無い環境でも編集するため、コンパイルに通らない
変更がそのまま入ってしまうことがある。実際に混入した以下を機械的に検出する。

    1. Function の中の Exit Sub (逆も) …………… コンパイルエラー
    2. Sub / Function の終端キーワード違い ……… コンパイルエラー
    3. 戻り値を一度も代入していない Function …… 常に既定値を返す
    4. 無条件 Exit の後ろの到達不能コード ……… 書いたつもりが動かない
    5. モジュール間で重複する Public 名 ………… Ambiguous name detected
    6. エラーハンドラの無いプロシージャ ………… CLAUDE.md Tier 2 違反
    7. どこにも定義が無い識別子 ………………… コンパイルエラー
    8. 宣言セクションの外にある宣言 …………… コンパイルエラー

    python3 tools/check_vba.py
"""
import io
import re
import sys
import pathlib
import collections

ROOT = pathlib.Path(__file__).resolve().parent.parent
SRC_DIR = ROOT / "src"

# 共有モジュール。自分のエラーは自分で記録できないため On Error Resume Next が
# 正しい設計であり、ハンドラ検査の対象外とする。改変もしない。
SKIP_HANDLER = {"ErrorLogger"}

PROC = re.compile(r"^\s*(Public|Private)\s+(Sub|Function)\s+([^\s(]+)")
END = re.compile(r"^\s*End\s+(Sub|Function)\s*$")
PUB_CONST = re.compile(r"^\s*Public\s+Const\s+(\w+)")
ANY_CONST = re.compile(r"^\s*(?:Public|Private)\s+Const\s+(\w+)")
MOD_VAR = re.compile(r"^\s*(?:Public|Private)\s+([A-Za-z_]\w*)(?:\(\))?\s+As\s")
# モジュールレベルの宣言。VBA はこれを先頭の宣言セクションにしか置けない。
MOD_DECL = re.compile(r"^\s*(?:Public|Private|Dim)\s+"
                      r"(?:Const\s+\w+|[A-Za-z_]\w*(?:\(\))?\s+As\s)")
LINENO = re.compile(r"^\s*\d+\s*")

# プロジェクト固有の識別子だけを対象にする(VBA 組み込みは検査しない)
IDENT = re.compile(
    r"\b(AS_[^\s(,)&.\"]+|AP_[^\s(,)&.\"]+|SS_\w+|SC_\w+|SV_\w+"
    r"|Shift(?:Click|Setup|Schema|Survey|Auto)_[^\s(,)&.\"]+"
    r"|Clr[A-Z]\w*|LBL_\w+|IDX_\w+|COL_\w+|SHT_\w+|NM_\w+|FS_\w+|CLR_\w+"
    r"|ST_\w+|KIND_\w+|HEAD_\w+|CFG_\w+|PAL_\w+|MARKER_\w+|LABEL_\w+"
    r"|DOC_\w+|NOTE_\w+|PALETTE_\w+|REQ_\w+|HOL_\w+|LOG_\w+|SYM_\w+"
    r"|MASK_\w+|SCAN_\w+|RPT_\w+|CNT_\w+|CLERK_\w+|MODULE_NAME)\b"
)


def strip_comments(text):
    return "\n".join(re.sub(r"'.*$", "", ln) for ln in text.split("\n"))


def procedures(lines):
    """(可視性, 種別, 名前, 開始行index, 終端行index) を順に返す。"""
    i = 0
    while i < len(lines):
        m = PROC.match(lines[i])
        if not m:
            i += 1
            continue
        j = i + 1
        while j < len(lines) and not END.match(lines[j]):
            j += 1
        yield m.group(1), m.group(2), m.group(3), i, j
        i = j + 1


def check():
    files = sorted(SRC_DIR.glob("*.bas"))
    if not files:
        print("[NG] src/*.bas が見つかりません")
        return 1

    problems = []
    pub_names = collections.defaultdict(set)
    defined = set()
    sources = {}

    for path in files:
        mod = path.stem
        sources[mod] = io.open(path, encoding="utf-8").read()
        for _, _, name, _, _ in procedures(sources[mod].split("\n")):
            defined.add(name)
        for ln in sources[mod].split("\n"):
            m = ANY_CONST.match(ln) or MOD_VAR.match(ln)
            if m:
                defined.add(m.group(1))

    for path in files:
        mod = path.stem
        lines = sources[mod].split("\n")

        # 8. 宣言セクションの外にある宣言
        #    プロシージャが1つでも現れた後の宣言はコンパイルエラーになる。
        seen_proc = False
        inside = False
        for k, ln in enumerate(lines):
            if PROC.match(ln):
                seen_proc = True
                inside = True
                continue
            if END.match(ln):
                inside = False
                continue
            if not inside and seen_proc and MOD_DECL.match(ln):
                problems.append(
                    f"{mod}: 宣言セクションの外に宣言があります "
                    f"(src/{path.name}:{k + 1}: {ln.strip()[:60]})")

        for ln in lines:
            m = PUB_CONST.match(ln)
            if m:
                pub_names[m.group(1)].add(mod)

        for vis, kind, name, start, end in procedures(lines):
            if vis == "Public":
                pub_names[name].add(mod)

            if end >= len(lines):
                problems.append(f"{mod}.{name}: End {kind} がありません")
                continue
            if END.match(lines[end]).group(1) != kind:
                problems.append(
                    f"{mod}.{name}: {kind} を {lines[end].strip()} で閉じています")

            body = lines[start + 1:end]
            code = [re.sub(r"'.*$", "", b) for b in body]

            # 1. Exit のキーワード違い
            wrong = "Function" if kind == "Sub" else "Sub"
            for k, b in enumerate(code):
                if re.search(r"\bExit\s+" + wrong + r"\b", b):
                    problems.append(
                        f"{mod}.{name}: {kind} の中に Exit {wrong} があります "
                        f"(src/{path.name}:{start + 2 + k})")

            # 3. Function が戻り値を一度も代入していない
            if kind == "Function":
                if not re.search(r"\b" + re.escape(name) + r"\s*=", "\n".join(code)):
                    problems.append(f"{mod}.{name}: 戻り値を代入していません")

            # 4. 無条件 Exit の直後の到達不能コード
            for k, b in enumerate(code[:-1]):
                t = LINENO.sub("", b).strip()
                unconditional = (t in ("Exit Sub", "Exit Function")
                                 or (t.endswith((": Exit Sub", ": Exit Function"))
                                     and not t.startswith("If ")))
                if not unconditional:
                    continue
                nxt = k + 1
                while nxt < len(code) and code[nxt].strip() == "":
                    nxt += 1
                if nxt < len(code):
                    n = LINENO.sub("", code[nxt]).strip()
                    if re.match(r"^[\w぀-ヿ一-鿿]+\s*=", n):
                        problems.append(
                            f"{mod}.{name}: 到達しないコードがあります "
                            f"(src/{path.name}:{start + 2 + nxt}: {n})")

            # 6. エラーハンドラ
            if mod not in SKIP_HANDLER:
                if not any("On Error GoTo ErrHandler" in b for b in code):
                    problems.append(f"{mod}.{name}: On Error GoTo ErrHandler がありません")
                if len([b for b in body if b.strip() == "ErrHandler:"]) != 1:
                    problems.append(f"{mod}.{name}: ErrHandler ラベルが1つではありません")

    # 5. モジュール間で重複する Public 名
    for name, mods in sorted(pub_names.items()):
        if len(mods) > 1:
            problems.append(
                f"Public 名 {name} が複数のモジュールにあります: {', '.join(sorted(mods))}")

    # 7. 定義が見つからない識別子
    for mod, text in sources.items():
        for name in sorted(set(IDENT.findall(strip_comments(text)))):
            if name not in defined:
                problems.append(f"{mod}: {name} の定義が見つかりません")

    if problems:
        print(f"[NG] {len(problems)} 件の問題があります")
        for p in problems:
            print("   - " + p)
        return 1

    print(f"[OK] src/*.bas の静的チェックを通過しました ({len(files)} ファイル)")
    return 0


if __name__ == "__main__":
    sys.exit(check())
