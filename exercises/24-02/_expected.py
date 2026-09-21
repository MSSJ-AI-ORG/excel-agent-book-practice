"""24-02 기대값을 practice 원본에서 따로 계산한다(정답 파일·정답 프로그램은 보지 않는다).

쓰는 법: python -X utf8 exercises/24-02/_expected.py

정산 시트 D열 수식이 하는 일을 그대로 따라간다.
  매출합계 s = 대장 D 를 부서 이름이 정확히 같은 줄만 더한 값(SUMIF)
  목표 t     = 기준 시트에서 그 부서의 목표금액(VLOOKUP)
  요율 r     = 이름 정의 「기준요율」이 가리키는 칸(기준!B8) — 기준 시트 C열 적용요율은 수식이 쓰지 않는다
  가산 g     = 이름 정의 「성과가산」 값
  건수 n     = 그 부서 이름이 정확히 같은 줄 수(COUNTIF)
  정산액     = ROUND( s>=t 이면 s*r + (n>4 이면 g) , 아니면 s*r*0.5 , -3 )
ROUND 는 엑셀처럼 «딱 절반이면 0 에서 먼 쪽» 으로 한다(Decimal ROUND_HALF_UP).
파이썬 기본 round 로 하면 어디가 달라지는지도 같이 찍는다.
"""
import json
import re
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[2]
wb = openpyxl.load_workbook(ROOT / "practice" / "부서 관리대장.xlsx")
led, base, st = wb["대장"], wb["기준"], wb["정산"]


def name_target(name):
    text = wb.defined_names[name].attr_text
    m = re.fullmatch(r"'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+)", text)
    return wb[m.group(1)][f"{m.group(2)}{m.group(3)}"].value if m else float(text)


r, g = name_target("기준요율"), name_target("성과가산")
target = {base[f"A{i}"].value: base[f"B{i}"].value for i in range(3, 6)}


def excel_round(x, digits):
    return int(Decimal(str(x)).quantize(Decimal(1).scaleb(-digits), rounding=ROUND_HALF_UP))


out, out_bank = {}, {}
for d in [st[f"A{i}"].value for i in range(3, 6)]:
    rows = [i for i in range(4, 20) if led[f"B{i}"].value == d]
    s, n = sum(led[f"D{i}"].value for i in rows), len(rows)
    raw = s * r + (g if n > 4 else 0) if s >= target[d] else s * r * 0.5
    out[d], out_bank[d] = excel_round(raw, -3), int(round(raw, -3))
    print(f"{d}: 매출합계 {s:,} · 목표 {target[d]:,} · 건수 {n} · 반올림 전 {raw:,} → 엑셀 ROUND {out[d]:,} / 파이썬 round {out_bank[d]:,}")
print(f"합계 {sum(out.values()):,} (파이썬 round 이면 {sum(out_bank.values()):,})")

cj = json.loads((Path(__file__).parent / "check.json").read_text(encoding="utf-8"))
pats = [c["pattern"] for c in cj["checks"] if c["type"] == "text_contains"]
for label, v in [*out.items(), ("합계", sum(out.values()))]:
    assert any(re.search(p, f"{label},{v}", flags=re.M) for p in pats), f"check.json 에 {label},{v} 줄이 없다"
print("check.json 에 박은 값이 원본 계산과 맞다.")
