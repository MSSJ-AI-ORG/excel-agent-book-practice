# -*- coding: utf-8 -*-
"""P5-01 매크로목록의 기대값 — 16-01·16-02 의 줄마다 값에 더해, 모듈 단위 «규칙 판»(R2 있음이고 R4 전부 통과면 현행).

원본 = practice/부서 관리대장.xlsm. 엑셀을 열지 않고 oletools 로 VBA 모듈을 꺼내 센다.
정답 파일은 보지 않는다. 이 값들이 check.json 에 «expect» 로 박혀 있고, 아래를 돌리면 같은지 보여 준다.

쓰는 법: python -X utf8 exercises/P5-01/_expected.py
"""
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
XLSM = ROOT / "practice" / "부서 관리대장.xlsm"
DECL = re.compile(r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([^\s(]+)")
END = re.compile(r"^\s*End\s+(?:Sub|Function|Property)\b")


def modules():
    """코드가 한 줄이라도 든 모듈 → 줄 목록(Attribute 줄은 뺀다). 속이 빈 시트 모듈은 뺀다."""
    from oletools.olevba import VBA_Parser

    out = {}
    p = VBA_Parser(str(XLSM))
    try:
        for _, _, fname, code in p.extract_macros():
            lines = [ln for ln in code.splitlines() if not ln.startswith("Attribute ")]
            if any(ln.strip() for ln in lines):
                out[Path(fname).stem] = lines
    finally:
        p.close()
    return out


def is_comment(ln):
    return ln.strip().startswith("'")


def procedures(mods=None):
    """[(모듈, 이름, 종류, 선언 윗줄, 본문 줄들)] — 코드에 나온 차례대로."""
    mods = mods or modules()
    out = []
    for m, lines in mods.items():
        cur = None
        for i, ln in enumerate(lines):
            d = DECL.match(ln)
            if d and not is_comment(ln):
                cur = (m, d.group(2), d.group(1).split()[0], lines[i - 1] if i else "", i)
            elif cur and END.match(ln):
                out.append(cur[:4] + (lines[cur[4]:i + 1],))
                cur = None
    return out


def by_name(procs):
    return sorted(procs, key=lambda p: p[1])      # 프로시저 이름 가나다순


def r2(lines):
    return "있음" if any(re.match(r"^\s*Option\s+Explicit\b", ln) for ln in lines) else "없음"


def r4_pass(prev_line):
    s = prev_line.strip()
    return s.startswith("'") and s[1:].strip() != ""


def called_by_other(name, procs):
    pat = re.compile(r"(?<![\w.])" + re.escape(name) + r"(?![\w])")
    for m, n, kind, prev, body in procs:
        if n == name:
            continue
        if any(pat.search(ln) for ln in body if not is_comment(ln)):
            return True
    return False


def compare(expected):
    check = json.loads((Path(__file__).parent / "check.json").read_text(encoding="utf-8"))
    got = {(c["sheet"], c["cell"]): c.get("expect") for c in check["checks"] if c["type"] == "cell_equals"}
    bad = [(k, v, got.get(k, "(없음)")) for k, v in expected.items() if got.get(k, "(없음)") != v]
    return bad


SHEET = "매크로목록"


def expected_cells():
    mods = modules()
    all_procs = procedures(mods)
    verdict = {}
    for m, lines in mods.items():
        ps = [p for p in all_procs if p[0] == m]
        verdict[m] = "현행" if r2(lines) == "있음" and all(r4_pass(p[3]) for p in ps) else "구판"
    exp = {}
    for r, (m, n, kind, prev, body) in enumerate(by_name(all_procs), start=2):
        exp[(SHEET, f"A{r}")] = f"{m}.bas"
        exp[(SHEET, f"B{r}")] = n
        exp[(SHEET, f"C{r}")] = kind
        exp[(SHEET, f"F{r}")] = "예" if called_by_other(n, all_procs) else "아니오"
        exp[(SHEET, f"I{r}")] = verdict[m]
    return exp


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    exp = expected_cells()
    for (sheet, cell), v in exp.items():
        print(f"{sheet}!{cell} = {v!r}")
    bad = compare(exp)
    print("check.json 과 같은가:", not bad)
    for b in bad:
        print("  다름:", b)
