"""24-02 정답 예시 — 규칙은 rules.json 에서 읽고, 부서별 매출합계와 정산액을 CSV 로 낸다.

쓰는 법(저장소 뿌리에서):
    python -X utf8 exercises/24-02/solution/screen.py exercises/24-02/solution/rules.json exercises/24-02/solution/out

- 엑셀을 띄우지 않는다(openpyxl 로 파일만 읽는다) · 돌 때 AI 를 부르지 않는다 · 원본을 고치지 않는다.
- 반올림은 엑셀 ROUND 와 같게 한다. 파이썬 기본 round 는 딱 절반일 때 짝수 쪽으로 가서
  172,500 을 172,000 으로 만든다(엑셀은 173,000). 그래서 Decimal 의 ROUND_HALF_UP 을 쓴다.
- rules.json 의 「검산」 칸에 엑셀이 낸 값이 적혀 있으면 돌 때마다 대 보고, 다르면 알린다.
"""
import csv
import hashlib
import json
import sys
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[3]


def excel_round(x, digits):
    return int(Decimal(str(x)).quantize(Decimal(1).scaleb(-digits), rounding=ROUND_HALF_UP))


def main(rules_path, out_dir):
    rules = json.loads(Path(rules_path).read_text(encoding="utf-8"))
    src = ROOT / rules["원본"]["값"]
    before = hashlib.sha256(src.read_bytes()).hexdigest()

    L = rules["대장"]
    wb = openpyxl.load_workbook(src, read_only=True)
    led = wb[L["시트"]]
    rows = [(led[f"{L['부서열']}{r}"].value, led[f"{L['금액열']}{r}"].value) for r in range(L["첫줄"], L["끝줄"] + 1)]
    tgt_ws = wb[rules["목표"]["시트"]]
    target = {row[0].value: row[1].value for row in tgt_ws[rules["목표"]["범위"]]}
    wb.close()

    rate, bonus = rules["기준요율"]["값"], rules["성과가산"]["값"]
    cnt_limit, half = rules["건수기준"]["값"], rules["미달배율"]["값"]
    digits = rules["반올림"]["자릿수"]

    sales, settle = {}, {}
    for d in rules["정산대상"]["값"]:
        mine = [amt for dept, amt in rows if dept == d and isinstance(amt, (int, float))]  # 이름이 정확히 같은 줄만
        s, n = sum(mine), len(mine)
        raw = s * rate + (bonus if n > cnt_limit else 0) if s >= target[d] else s * rate * half
        sales[d], settle[d] = s, excel_round(raw, digits)

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    for name, head, data in (("매출합계.csv", "매출합계", sales), ("정산.csv", "정산액", settle)):
        with open(out_dir / name, "w", encoding="utf-8", newline="") as f:
            w = csv.writer(f, lineterminator="\n")
            w.writerow(["부서", head])
            for d, v in data.items():
                w.writerow([d, v])
            if name == "정산.csv":
                w.writerow(["합계", sum(data.values())])

    for d in settle:
        print(f"{d}: 매출합계 {sales[d]:,} · 정산액 {settle[d]:,}")
    print(f"합계: 매출합계 {sum(sales.values()):,} · 정산액 {sum(settle.values()):,}")

    check = rules.get("검산", {}).get("값", {})
    if check:
        mine = {**settle, "합계": sum(settle.values())}
        diff = {k: (check[k], mine.get(k)) for k in check if check[k] != mine.get(k)}
        print("검산(엑셀 값과 대조):", "전부 같음" if not diff else f"다름 {diff}")
    else:
        print("검산: rules.json 에 엑셀 값이 아직 적혀 있지 않다")

    after = hashlib.sha256(src.read_bytes()).hexdigest()
    print("원본이 그대로인가(지문 전·뒤):", before == after)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
