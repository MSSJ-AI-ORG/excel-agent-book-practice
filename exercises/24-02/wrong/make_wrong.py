"""24-02 오답 예시를 만든다 — 흔한 실수 하나: 정산액을 파이썬 기본 round 로 반올림한다.

쓰는 법(저장소 뿌리에서):
    python -X utf8 exercises/24-02/wrong/make_wrong.py exercises/24-02/solution/rules.json exercises/24-02/wrong/out

정답 프로그램(solution/screen.py)의 반올림 함수만 파이썬 기본 round 로 바꿔 끼워 돌린다.
파이썬 round 는 딱 절반일 때 짝수 쪽으로 간다 — 영업2팀 172,500 이 172,000 이 된다(엑셀 ROUND 는 173,000).
"""
import importlib.util
import sys
from pathlib import Path

spec = importlib.util.spec_from_file_location("screen", Path(__file__).resolve().parents[1] / "solution" / "screen.py")
screen = importlib.util.module_from_spec(spec)
spec.loader.exec_module(screen)
screen.excel_round = lambda x, digits: int(round(x, digits))  # ← 이 한 줄이 실수
screen.main(sys.argv[1], sys.argv[2])
