"""24-01 기대값을 practice 원본에서 따로 계산한다(정답 파일·정답 프로그램은 보지 않는다).

쓰는 법: python -X utf8 exercises/24-01/_expected.py

이 과제는 CSV(글 파일)를 채점하므로 check.json 에 숫자를 적을 수밖에 없다.
그 숫자(부서별 매출합계)를 원본 대장의 입력 칸에서 지금 정산 수식과 같은 규칙 —
부서 이름이 «정확히» 같은 줄만 더한다(SUMIF) — 으로 따로 계산해 check.json 과 대 본다.
정산 시트에 저장된 계산값은 쓰지 않는다(이 파일에는 0 이 저장돼 있다).
"""
import json
import re
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[2]
wb = openpyxl.load_workbook(ROOT / "practice" / "부서 관리대장.xlsx")
led, st = wb["대장"], wb["정산"]
depts = [st[f"A{r}"].value for r in range(3, 6)]
sales = {d: sum(led[f"D{r}"].value for r in range(4, 20) if led[f"B{r}"].value == d) for d in depts}
for d, v in sales.items():
    print(f"{d},{v}")

saved = openpyxl.load_workbook(ROOT / "practice" / "부서 관리대장.xlsx", data_only=True)["정산"]
print("참고 — 정산 시트에 저장된 계산값:", [saved[f"B{r}"].value for r in range(3, 6)])

cj = json.loads((Path(__file__).parent / "check.json").read_text(encoding="utf-8"))
pats = [c["pattern"] for c in cj["checks"] if c["type"] == "text_contains"]
for d, v in sales.items():
    assert any(re.search(p, f"{d},{v}", flags=re.M) for p in pats), f"check.json 에 {d},{v} 줄이 없다"
print("check.json 에 박은 값이 원본 계산과 맞다.")
