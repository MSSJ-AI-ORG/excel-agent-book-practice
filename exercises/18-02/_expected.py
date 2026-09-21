# -*- coding: utf-8 -*-
"""18-02 채점 규칙 「새 프로시저 이름이 원본 이름과 겹치지 않는다」에 넣을 이름을 원본에서 모은다.

원본 = practice/부서 관리대장.xlsm 의 모든 모듈(oletools 로 꺼낸다 — 엑셀을 열지 않는다).
Sub·Function·Property 선언의 이름을 코드에 나온 차례대로 모은다.

쓰는 법: python -X utf8 exercises/18-02/_expected.py
"""
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
XLSM = ROOT / "practice" / "부서 관리대장.xlsm"
DECL = re.compile(r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([^\s(]+)")


def existing_names():
    from oletools.olevba import VBA_Parser

    p = VBA_Parser(str(XLSM))
    names = []
    try:
        for _, _, _, code in p.extract_macros():
            for ln in code.splitlines():
                m = DECL.match(ln)
                if m and not ln.strip().startswith("'") and m.group(1) not in names:
                    names.append(m.group(1))
    finally:
        p.close()
    return names


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    names = existing_names()
    print("원본의 프로시저 이름:", names)
    check = json.loads((Path(__file__).parent / "check.json").read_text(encoding="utf-8"))
    pat = next(c["pattern"] for c in check["checks"] if "겹치지" in c["what"])
    print("check.json 규칙에 이 이름들이 다 들어 있나:", all(re.escape(n) in pat for n in names))
