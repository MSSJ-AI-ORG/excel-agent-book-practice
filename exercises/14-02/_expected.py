# -*- coding: utf-8 -*-
"""14-02 채점 규칙 「주석과 빈 줄을 빼면 원본과 같다」의 정규식을 원본에서 만든다.

원본 = practice/부서 관리대장.xlsm 안의 「집계매크로」 모듈(oletools 로 꺼낸다 — 엑셀을 열지 않는다).
주석 줄과 빈 줄을 빼고 남은 코드 줄을 차례대로 잇되, 그 사이에는 주석 줄·빈 줄만 끼어들 수 있게 한다.
줄 끝에 붙인 주석(코드 뒤 ' …)은 허용한다. 줄 안의 공백 개수는 따지지 않는다.

쓰는 법: python -X utf8 exercises/14-02/_expected.py   → check.json 과 같은지 보여 준다
"""
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
XLSM = ROOT / "practice" / "부서 관리대장.xlsm"
MODULE = "집계매크로"


def module_code(name):
    from oletools.olevba import VBA_Parser

    p = VBA_Parser(str(XLSM))
    try:
        for _, _, fname, code in p.extract_macros():
            if Path(fname).stem == name:
                return [ln for ln in code.splitlines() if not ln.startswith("Attribute ")]
    finally:
        p.close()
    raise SystemExit(f"{name} 모듈을 찾지 못했습니다")


def code_lines(lines):
    return [ln.strip() for ln in lines if ln.strip() and not ln.strip().startswith("'")]


def code_same_pattern():
    gap = r"(?:[ \t]*(?:'[^\n]*)?\r?\n)*"          # 주석 줄·빈 줄은 몇 줄이든
    head = r"\A(?:Attribute[^\n]*\n)*" + gap         # 내보낸 .bas 맨 위의 Attribute 줄
    tail = r"[ \t]*(?:'[^\n]*)?(?:\r?\n" + gap + r")?\Z"
    parts = []
    for ln in code_lines(module_code(MODULE)):
        body = r"[ \t]+".join(re.escape(tok) for tok in ln.split())
        parts.append(r"[ \t]*" + body)
    joined = (r"[ \t]*(?:'[^\n]*)?\r?\n" + gap).join(parts)
    return head + joined + tail


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    lines = code_lines(module_code(MODULE))
    print(f"원본 {MODULE} 의 코드 줄(주석·빈 줄 뺌): {len(lines)}줄")
    pat = code_same_pattern()
    check = json.loads((Path(__file__).parent / "check.json").read_text(encoding="utf-8"))
    in_check = check["checks"][0]["pattern"]
    print("check.json 의 정규식과 같은가:", pat == in_check)
