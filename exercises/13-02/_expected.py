# -*- coding: utf-8 -*-
"""13-02 채점 규칙 「없는 시트 이름을 부르는 줄이 없다」에 넣을 이름을 원본에서 센다.

원본 코드(practice/부서 관리대장.xlsm 의 모든 모듈)에서 Sheets("…")·Worksheets("…") 로 부르는 이름을 모으고,
그 파일의 실제 시트 목록에 없는 이름만 남긴다. 엑셀을 열지 않는다(oletools·openpyxl 로 읽는다).

쓰는 법: python -X utf8 exercises/13-02/_expected.py
"""
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
XLSM = ROOT / "practice" / "부서 관리대장.xlsm"
REF = re.compile(r'(?:Sheets|Worksheets)\(\s*"([^"]+)"\s*\)')


def referenced_sheet_names():
    from oletools.olevba import VBA_Parser

    p = VBA_Parser(str(XLSM))
    names = []
    try:
        for _, _, _, code in p.extract_macros():
            for line in code.splitlines():
                if line.strip().startswith("'"):
                    continue
                for n in REF.findall(line):
                    if n not in names:
                        names.append(n)
    finally:
        p.close()
    return names


def actual_sheet_names():
    import openpyxl

    return openpyxl.load_workbook(XLSM, read_only=True).sheetnames


def missing_sheet_names():
    actual = set(actual_sheet_names())
    return [n for n in referenced_sheet_names() if n not in actual]


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    print("코드가 부르는 시트:", referenced_sheet_names())
    print("실제 시트        :", actual_sheet_names())
    miss = missing_sheet_names()
    print("없는 이름        :", miss)
    check = json.loads((Path(__file__).parent / "check.json").read_text(encoding="utf-8"))
    pat = check["checks"][1]["pattern"]
    print("check.json 규칙에 이 이름들이 다 들어 있나:", all(n in pat for n in miss))
