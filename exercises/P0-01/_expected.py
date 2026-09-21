"""P0-01 기대값을 practice 원본에서 따로 계산하고, check.json 에 박은 값과 대 본다.

쓰는 법: python -X utf8 exercises/P0-01/_expected.py

채점기에는 아직 「시트 이름 목록」·「값이 있는 마지막 줄 번호」·「수식 칸 수」를 바로 세는 기능이 없다.
그래서 check.json 에는 값을 박아 두었고, 그 값이 원본에서 나온 것인지를 이 파일이 확인한다.
정답 파일은 보지 않는다.
"""
import json
import sys
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
wb = openpyxl.load_workbook(ROOT / "practice" / "부서 관리대장.xlsx")


def filled(v):
    return v not in (None, "")


def is_formula(v):
    return (isinstance(v, str) and v.startswith("=")) or type(v).__name__ in ("ArrayFormula", "DataTableFormula")


want = {}
print(f"{'시트':<10} 마지막 줄  수식 칸")
for i, ws in enumerate(wb.worksheets, start=2):
    cells = [c for row in ws.iter_rows() for c in row if filled(c.value)]
    last_row = max(c.row for c in cells)
    n_formula = sum(1 for c in cells if is_formula(c.value))
    print(f"{ws.title:<10} {last_row:>6} {n_formula:>8}")
    want[f"A{i}"], want[f"B{i}"], want[f"C{i}"] = ws.title, last_row, n_formula
want[f"A{len(wb.worksheets) + 2}"] = ""  # 시트 목록 다음 줄은 비어 있어야 한다

check = json.loads((HERE / "check.json").read_text(encoding="utf-8"))
bad = 0
for rule in check["checks"]:
    if rule.get("sheet") == "뼈대" and "expect" in rule:
        cell = rule["cell"]
        if cell not in want or want[cell] != rule["expect"]:
            bad += 1
            print("어긋남:", cell, "check.json =", rule["expect"], "원본 =", want.get(cell, "(원본에서 계산 안 함)"))
        want.pop(cell, None)
for cell in want:
    bad += 1
    print("check.json 에 없는 칸:", cell, want[cell])

print("OK — check.json 에 박은 값이 원본에서 센 값과 같다" if not bad else f"어긋남 {bad}건")
sys.exit(1 if bad else 0)
