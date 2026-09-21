"""07-02 날짜검산!B1 의 기대값은 채점기가 원본 대장 A4:A19 에서 «글자가 적힌 칸 수»로 센다.
그 칸이 모두 날짜로 읽을 수 있는 글자인지(13월 같은 값이 없는지)를 여기서 따로 확인한다.
형식별 건수(점 형식 2건 등)도 같이 보여 준다.

쓰는 법(저장소 맨 위 폴더에서):  python -X utf8 exercises/07-02/_expected.py
"""
import datetime as dt
import re
import sys
import warnings
from pathlib import Path

import openpyxl

warnings.filterwarnings("ignore")
ROOT = Path(__file__).resolve().parents[2]
ws = openpyxl.load_workbook(ROOT / "practice" / "부서 관리대장.xlsx")["대장"]


def parse(s):
    m = re.fullmatch(r"(\d{4})[-.](\d{1,2})[-.](\d{1,2})", s) or re.fullmatch(r"(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일", s)
    if not m:
        return None
    try:
        return dt.date(*map(int, m.groups()))
    except ValueError:
        return None


kinds, bad, filled = {"하이픈": 0, "점": 0, "한글": 0}, [], 0
for r in range(4, 20):
    v = ws.cell(r, 1).value
    if v in (None, ""):
        continue
    filled += 1
    kinds["점" if "." in v else "한글" if "년" in v else "하이픈"] += 1
    if parse(v.strip()) is None:
        bad.append((r, v))
print("글자가 적힌 칸 수(채점기가 세는 기대값):", filled)
print("형식별 건수:", kinds)
print("날짜로 못 읽는 칸:", bad or "없음")
sys.exit(1 if bad else 0)
