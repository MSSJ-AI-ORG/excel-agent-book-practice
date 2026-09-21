"""05-01 check.json 에 박아 둔 값(표 수·차트 수·이름 차례)을 practice 원본에서 따로 계산해 대 본다.

쓰는 법(저장소 맨 위 폴더에서):  python -X utf8 exercises/05-01/_expected.py
정답 파일은 보지 않는다. 원본 「부서 관리대장.xlsx」만 읽는다.
"""
import json
import sys
import warnings
from pathlib import Path

import openpyxl

warnings.filterwarnings("ignore")
ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / "practice" / "부서 관리대장.xlsx"

wb = openpyxl.load_workbook(SRC)
tables = sum(len(ws.tables) for ws in wb.worksheets)
charts = sum(len(ws._charts) for ws in wb.worksheets)
names = sorted(n for n, d in wb.defined_names.items() if not getattr(d, "hidden", False))
got = {"B3": tables, "B4": charts, **{f"A{i}": n for i, n in enumerate(names, start=7)}}

check = json.loads((Path(__file__).with_name("check.json")).read_text(encoding="utf-8"))
written = {c["cell"]: c["expect"] for c in check["checks"] if c.get("sheet") == "구조" and "expect" in c}

print("원본에서 센 값:", got)
print("check.json 에 적힌 값:", written)
bad = {k: (v, got.get(k)) for k, v in written.items() if got.get(k) != v}
print("다른 것:", bad or "없음")
sys.exit(1 if bad else 0)
