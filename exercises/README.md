# 과제 문서 형식

과제 하나는 폴더 하나입니다. 폴더 이름이 과제 번호입니다.

```
exercises/
├── CATALOGUE.md        전체 과제 목록
└── 00-01/
    ├── README.md       읽는 사람이 보는 과제 문서
    └── check.json      채점기가 읽는 규칙
```

번호: `NN-01` 은 NN장 기본, `NN-02` 는 NN장 응용, `PN-01` 은 N부 종합(그 부의 마지막 장 뒤)입니다.

---

## 1. 과제 문서(README.md) — 읽는 사람을 위한 부분

모든 과제 문서는 같은 차례로 씁니다.

| 칸 | 무엇을 적나 |
|---|---|
| 제목 | `# 00-01 도구가 엑셀 파일을 정말 읽고 쓰는지 실물로 확인하기` |
| 어느 장 | 그 과제가 다시 써 보게 하는 절 이름을 그대로 적습니다. 예: 「0.5 깔고 한 번 말 걸어 보기」 |
| 시작 파일 | `practice/` 의 어느 파일에서 시작하는지. **원본은 열지 말고 사본을 만들어** 시작한다고 꼭 적습니다 |
| 할 일 | 순서대로 번호를 붙여 적습니다. 책에 나온 시킬 말을 쓸 때는 책 그대로 인용하고, 과제에서 새로 붙인 말은 따로 표시합니다 |
| 확인법 | 그 장에서 배운 확인 문장을 인용합니다. 예: 「"했습니다"라는 답을 확인으로 쓰지 않는다」 |
| 채점 | auto 면 채점 명령과 제출 파일 이름, manual 이면 「스스로 확인하는 과제」 |

쉽게 씁니다. 짧은 문장, 용어는 처음 나올 때 풀어 씁니다.

★ **클로드 코드 화면에 나오는 글자(권한을 묻는 문구 같은 것)는 짐작해서 적지 않습니다.**
「이런 창이 뜬다」를 적어야 할 때는 그림 자리만 남깁니다.

```
> [그림 00-01-A] 설치해도 되는지 묻는 화면 (촬영 예정)
```

---

## 2. 채점 규칙(check.json) — 과제를 만드는 사람을 위한 부분

```json
{
  "id": "00-01",
  "title": "…",
  "chapter": "0장 0.5 깔고 한 번 말 걸어 보기",
  "mode": "auto",
  "file_kind": "workbook",
  "start_file": "practice/이번 달 매출.xlsx",
  "submit_file": "00-01_점검.xlsx",
  "checks": [ { "type": "…", "what": "화면에 보일 항목 이름", … } ],
  "checklist": [ "채점기가 못 보는 것 — 사람이 확인할 것" ]
}
```

| 칸 | 뜻 |
|---|---|
| `mode` | `auto`(채점기가 채점) · `manual`(확인 목록만 보여 주고 끝) |
| `file_kind` | `workbook`(xlsx·xlsm, 기본값) · `text`(.bas·.md 같은 글 파일) |
| `checks` | 채점 항목. 위에서부터 차례로 보여 준다 |
| `checklist` | manual 과제의 확인 목록. auto 과제에도 붙일 수 있다(채점 결과 아래에 나온다) |

### 채점 항목 종류 — 통합 문서

| type | 보는 것 | 칸 |
|---|---|---|
| `sheet_exists` | 시트가 있나 | `sheet` |
| `cell_equals` | 칸의 값(수식이면 계산값)이 기대값과 같나 | `sheet` `cell` `expect` 또는 `expect_from` · `tolerance`(선택) · `source_note`(기대값을 어떻게 셌는지 한 줄) |
| `cell_is_formula` | 칸에 수식이 있나 · 특정 함수가 들어 있나 | `sheet` `cell` `functions`(선택) |
| `cell_is_constant` | 칸이 수식이 아니라 손으로 적은 값인가(21.1 「기준은 숫자로」) | `sheet` `cell` |
| `range_same_as_source` | 범위가 원본과 같은가 — 값 칸은 값으로, 수식 칸은 수식 글자로 댄다 | `sheet` `range` `source` `source_sheet`(선택) |
| `practice_untouched` | practice 원본이 처음 받은 그대로인가(지문 대조) | `file` |
| `same_bytes_as` | 제출 파일이 어떤 파일과 바이트까지 같은가(사본·복원 확인) | `file` |
| `font_bold` · `number_format` · `fill_present` | 범위 전체가 굵게 · 그 표시 형식 · 채우기 색이 있나 | `sheet` `range` (`format`) |
| `freeze_panes` | 틀 고정 자리 | `sheet` `cell` |
| `column_width_min` | 열 너비가 얼마 이상인가(`#####` 막기) | `sheet` `columns` `min` |

### 채점 항목 종류 — 글 파일(`file_kind: "text"`)

| type | 보는 것 | 칸 |
|---|---|---|
| `text_contains` | 이 모양의 줄이 있나(정규식) | `pattern` `ignore_comments`(VBA 주석 줄 무시) `fail_hint`(선택) |
| `text_absent` | 이 모양의 줄이 없나 — 있으면 몇 번째 줄인지 알려 준다 | 같음 |

글 파일은 UTF-8 로 읽고, 안 되면 cp949(엑셀 VBA 편집기에서 내보낸 .bas)로 읽습니다.

### 기대값은 원본에서 계산한다 — `expect_from`

★ **기대값을 check.json 에 숫자로 박지 않습니다.** 채점기가 practice 원본 자료에서 직접 계산하게 적습니다.
정답 파일에서 값을 거꾸로 읽어 기대값을 맞추는 일이 없게 하려는 것입니다.
그래서 원본의 **수식 칸**을 재료로 쓰면 채점기가 규칙 오류(?)를 냅니다. 입력 칸(손으로 적은 값)에서 계산하도록 적어야 합니다.

```json
"expect_from": {
  "source": "practice/부서 관리대장.xlsx",
  "sheet": "대장",
  "op": "sum",
  "range": "D4:D19",
  "where": [ { "range": "B4:B19", "equals": "영업1팀", "normalize": "remove_spaces" } ]
}
```

| op | 계산 | 칸 |
|---|---|---|
| `cell` | 칸 하나의 값 | `cell` |
| `sum` · `count` · `count_distinct` | 범위의 합 · 빈칸 아닌 칸 수 · 서로 다른 값의 수 | `range` `where`(선택) `normalize`(count_distinct) |
| `sumproduct` | 여러 범위를 행마다 곱해 더한다(예: 수량×단가) | `ranges` `where`(선택) |
| `lookup` | 한 범위에서 값을 찾아 같은 줄의 다른 값 | `key_range` `return_range` `equals` |
| `sheet_count` · `defined_name_count` | 시트 수 · (숨김 아닌) 이름 정의 수 | — |
| `defined_name_value` | 이름 정의가 가리키는 값(고정값이든 칸이든) | `name` |

`where` 는 여러 개를 겹칠 수 있고(모두 만족하는 줄만), 줄 수가 계산 범위와 같아야 합니다.
조건: `equals`(같다) · `not_blank`(빈칸 아님) · `gte` `gt` `lte` `lt`(크기 비교). `normalize` 는 `strip`(앞뒤 공백 빼기) · `remove_spaces`(공백 모두 빼기).
`equals` 에 `{"from": {…}}` 를 주면 그 값을 다시 계산해 조건으로 씁니다. `"source": "@submitted"` 는 **제출한 파일 자신**을 가리킵니다(예: 독자가 N1 에 적은 담당자 이름).

여러 값을 엮을 때는 `calc` 를 씁니다.

```json
"expect_from": { "calc": {
  "expr": "round_excel(if_(s >= t, s*r, s*r*0.5), -3)",
  "vars": { "s": { … sum … }, "t": { … lookup … }, "r": { "source": "…", "op": "defined_name_value", "name": "기준요율" } }
} }
```

`calc` 식에는 더하기·빼기·곱하기·나누기·비교·`and`·`or` 와 함수 `round_excel` `if_` `min` `max` `abs` 만 쓸 수 있습니다.
`round_excel` 은 엑셀 `ROUND` 와 같게 **딱 절반이면 0 에서 먼 쪽으로** 반올림합니다(파이썬 기본 `round` 는 짝수 쪽으로 가서 천 원이 어긋납니다 — 책 24.4).

### 계산값이 없을 때

제출 파일의 수식 칸에 계산값이 없거나, 파일에 «열 때 다시 계산» 표시가 붙어 있으면 그 항목은 ✗ 가 아니라 `!` 로 나옵니다.
「엑셀에서 한 번 열고 저장한 뒤 다시 채점하세요」라는 안내가 붙습니다.

### 새 규칙 종류를 더할 때

1. `grader/check.py` 에 `rule_…` 함수 하나를 더하고 `RULES`(글 파일이면 `TEXT_RULES`)에 이름을 올린다.
2. `grader/test_check.py` 에 **맞는 파일은 ✓, 겨눈 것만 틀린 파일은 ✗** 두 쪽 시험을 같이 넣는다.
3. 이 표에 한 줄을 더한다.
