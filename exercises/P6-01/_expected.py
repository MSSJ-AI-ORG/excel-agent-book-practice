"""P6-01 기대값 — 검사기 여덟 줄은 21-01 과 같다. 21-01 의 계산을 그대로 돌린다.

쓰는 법: python -X utf8 exercises/P6-01/_expected.py
check.json 에 박은 값은 검사기 8번 줄(중복 줄 수) 0 하나뿐이고, 나머지는 채점기가 원본에서 직접 센다.
"""
import runpy
from pathlib import Path

runpy.run_path(str(Path(__file__).resolve().parents[1] / "21-01" / "_expected.py"), run_name="__main__")
