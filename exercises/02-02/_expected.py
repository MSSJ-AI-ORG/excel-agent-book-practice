"""02-02 기대값을 practice 원본에서 따로 계산한다(정답 파일은 보지 않는다).

쓰는 법: python -X utf8 exercises/02-02/_expected.py
잠근 사본은 «저장했는데도 안 바뀐» 파일이라 원본과 바이트까지 같아야 하고, B2 는 원본 그대로여야 한다.
"""
import hashlib
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[2]
src = ROOT / "practice" / "이번 달 매출.xlsx"

print("원본 매출!B2:", openpyxl.load_workbook(src)["매출"]["B2"].value)
print("원본 지문(sha256):", hashlib.sha256(src.read_bytes()).hexdigest())
