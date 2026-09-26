"""작성 점검 — 사업별 설명서가 예산과 작성요령대로 쓰였는지 본다.

근거는 예산과 「2027년도 본예산 사업담당자 전달사항(사업별 설명서 작성요령·세목 등
편성 유의사항)」('26. 8.)이다. 금액 검산(crosscheck)과 다르게 **틀렸다고 단정할 수
없는 것이 많다.** 기준단가를 벗어난 데는 이유가 있을 수 있고, 표기 규칙은 입력본의
값을 바꾸지 않는다. 그래서 전부 '확인 필요'로 싣고 입력본 생성을 막지 않는다.

지키는 원칙
- 요령에 **기준으로 적힌 값만** 기준으로 쓴다. 요령의 산출식 예시(강사수당
  100천원×2명×2시간 등)는 쓰는 법을 보인 것이지 상한이 아니다. 표본 설명서도
  강사수당 160천원을 쓴다. 예시를 상한으로 쓰면 멀쩡한 줄이 전부 걸린다.
- 같은 규칙이 한 사업에서 여러 번 걸리면 한 건으로 묶고 곳 수를 적는다.
  가운뎃점 하나 때문에 목록이 수십 줄이 되면 볼 것이 묻힌다.
- 작년 확정본과 비교할 때는 **사업 + 상위 항목 + 항목명**이 모두 같을 때만 짝짓는다.
  이름만으로 맞추면 남의 사업 '시설임차료'와 견줘 +5,900% 같은 가짜 인상이 나온다.
"""

from __future__ import annotations

import re
from collections import OrderedDict
from dataclasses import dataclass, field
from typing import Iterable, Optional

from crosscheck import Issue, _squeeze
from hwpx_native_parser import is_private_use, native_chunks
from plan_parser import HEADING, _cells, _is_organisation

SEVERITY = "확인 필요"

# 분류 → 화면에 보이는 이름. goal 로 쓴다(ui_common.GOAL_LABEL 에 같은 이름을 둔다).
CATEGORIES = OrderedDict([
    ("작성-금액", "금액 표기"),
    ("작성-표기", "표기법"),
    ("작성-용어", "용어 통일"),
    ("작성-단가", "기준단가"),
    ("작성-산출", "산출내역"),
    ("작성-증감", "증감사유"),
    ("작성-비목", "원가통계비목"),
])
PREFIX = "작성"

SOURCE = "작성요령"      # 문구 끝에 붙는 근거 표시


@dataclass
class WritingReport:
    """작성 점검 결과. crosscheck.Report 와 달리 '확인 필요'를 버리지 않는다."""

    issues: list[Issue] = field(default_factory=list)
    skipped: list = field(default_factory=list)

    def of(self, goal: str) -> list[Issue]:
        return [issue for issue in self.issues if issue.goal.startswith(goal)]


# ---------------------------------------------------------------- 문서 훑기

@dataclass
class _Block:
    project: str          # 사업명(설명서 파서가 붙인 이름). 못 찾으면 제목 줄
    kind: str             # text / table
    lines: list           # text: 줄 목록, table: 칸 목록(글자만)
    rows: list = field(default_factory=list)


def _blocks(plan_path: str, projects: list) -> list[_Block]:
    """문서를 사업 단위로 나눈 조각. 사업의 경계는 조직 표(정책·단위·세부사업)다.

    제목 줄('1. 초등교육과정 운영')만 보면 사업계획 안의 '1. 교육과정 실천…' 절과
    구분되지 않는다. 설명서 파서와 같은 기준(조직 표)을 쓴다.
    """
    chunks = native_chunks(plan_path)
    by_heading = {project.heading: project.name for project in projects if project.heading}
    blocks: list[_Block] = []
    current = ""
    last_heading = ""
    for chunk in chunks:
        text = chunk.get("text") or ""
        if chunk.get("type") == "text":
            lines = [line.strip() for line in text.splitlines() if line.strip()]
            for line in lines:
                if HEADING.match(line):
                    last_heading = line
            blocks.append(_Block(current, "text", lines))
            continue
        rows = _cells(text)
        if not rows:
            continue
        if _is_organisation(rows):
            current = by_heading.get(last_heading, last_heading)
            continue
        blocks.append(_Block(current, "table", [cell for row in rows for cell in row if cell], rows))
    return blocks


def _window(text: str, start: int, end: int, pad: int = 6) -> str:
    """걸린 곳 앞뒤 몇 글자. 한글 Ctrl+F 로 찾을 글자로도 쓴다."""
    left = max(0, start - pad)
    right = min(len(text), end + pad)
    snippet = text[left:right].strip()
    # 사설영역 글머리표는 찾기 칸에 붙여 넣어도 걸리지 않는다. 빼고 준다.
    return "".join(char for char in snippet if not is_private_use(char)).strip()


class _Collector:
    """같은 사업·같은 규칙은 한 건으로 묶는다."""

    def __init__(self) -> None:
        self.groups: "OrderedDict[tuple, dict]" = OrderedDict()

    def hit(self, goal: str, project: str, rule: str, message: str, snippet: str,
            index: Optional[int] = None, left=None, right=None) -> None:
        key = (goal, project, rule)
        group = self.groups.get(key)
        if group is None:
            self.groups[key] = {"message": message, "snippets": [snippet], "index": index,
                                "left": left, "right": right}
        elif snippet not in group["snippets"]:
            group["snippets"].append(snippet)

    def issues(self) -> list[Issue]:
        found = []
        for (goal, project, _rule), group in self.groups.items():
            snippets = group["snippets"]
            message = group["message"]
            if len(snippets) > 1:
                shown = " · ".join(f"'{one}'" for one in snippets[:3])
                message += f" — {len(snippets)}곳: {shown}" + (" 외" if len(snippets) > 3 else "")
            found.append(Issue(goal, SEVERITY, project, snippets[0], message,
                               group["left"], group["right"], index=group["index"], blocking=False))
        return found


# ---------------------------------------------------------------- 표기법 (문자표)

MIDDLE_DOT = re.compile(r"(?<=[가-힣A-Za-z0-9)])[·ㆍ‧∙](?=[가-힣A-Za-z0-9(])")
COLON = re.compile(r"(?<![\s:/])(?<!\d):|:(?![\s/])(?!\d)(?!$)")
TIME = re.compile(r"\d:\d")
ARROW = re.compile(r"\s+→|→\s+")
TILDE = re.compile(r"[∼～〜]")
YEAR_QUOTE = re.compile(r"[‘’`´]\d{2}(?=[.년\s)~]|$)")
DOC_DATE = re.compile(r"-\d+,\s*(\d{4}\.\s+\d{1,2}\.\s*\d{1,2}\.|\d{4}\.\d{1,2}\.\s+\d{1,2}\.)")
AREA_CODE = re.compile(r"(?<!\d)063-?\d{3}")
PERIOD_LINE = re.compile(r"사업기간\s*:\s*(.+)$")
PERIOD_OK = re.compile(r"^\d{4}\. \d{1,2}\. ~ (\d{4}\. )?\d{1,2}\.$")
EMPTY_COLON = re.compile(r"사업(목적|근거)\s*:\s*$")
FORBIDDEN_HEAD = re.compile(r"^(시기|방법)\s*:")
# 7칸 ◆ 글머리표. 한글 문서에는 사설영역 글리프(U+F02EF)로 들어 있다. ▸(U+F02FB) 아래의
# 시기·방법은 요령이 오히려 권하는 자리라 걸면 안 된다.
DIAMOND = ("\U000F02EF", "◆")


def _text_rules(block: _Block, out: _Collector) -> None:
    for line in block.lines:
        project = block.project
        for found in MIDDLE_DOT.finditer(line):
            out.hit("작성-표기", project, "가운뎃점",
                    f"가운뎃점은 문자표 일반구두점 '․'(유니코드 2024)을 씁니다. "
                    f"'{found.group()}'(U+{ord(found.group()):04X})이 쓰였습니다. ({SOURCE}: 문자표 사용)",
                    _window(line, found.start(), found.end(), 4))
        if not TIME.search(line) and "http" not in line:
            for found in COLON.finditer(line):
                out.hit("작성-표기", project, "콜론",
                        f"콜론(:)은 앞뒤를 다 띄웁니다. 예) 단위 : 천원 ({SOURCE}: 문자표 사용)",
                        _window(line, found.start(), found.end(), 5))
        for found in ARROW.finditer(line):
            out.hit("작성-표기", project, "화살표",
                    f"화살표(→)는 앞뒤를 띄우지 않습니다. 예) (2개월→4개월) ({SOURCE}: 문자표 사용)",
                    _window(line, found.start(), found.end(), 5))
        for found in TILDE.finditer(line):
            out.hit("작성-표기", project, "물결",
                    f"내지(~)는 키보드의 ~ 를 씁니다. '{found.group()}'(U+{ord(found.group()):04X})이 "
                    f"쓰였습니다. ({SOURCE}: 문자표 사용)", _window(line, found.start(), found.end(), 5))
        for found in YEAR_QUOTE.finditer(line):
            out.hit("작성-표기", project, "연도",
                    f"연도 줄임('27)의 따옴표는 작은따옴표 '(유니코드 0027)입니다. "
                    f"'{found.group()[0]}'이 쓰였습니다. ({SOURCE}: 문자표 사용)",
                    _window(line, found.start(), found.end(), 4))
        for found in DOC_DATE.finditer(line):
            out.hit("작성-표기", project, "공문대호",
                    f"공문대호 날짜는 띄우지 않습니다. 예) (교육부 OO과-123, 2026.9.20.) ({SOURCE}: 공문대호 날짜)",
                    _window(line, found.start(), found.end(), 8))
        for found in AREA_CODE.finditer(line):
            out.hit("작성-표기", project, "연락처",
                    f"연락처는 국번 없이 씁니다(063 입력 금지). 예) 239-3000 ({SOURCE}: 제목 하단 박스)",
                    _window(line, found.start(), found.end(), 6))
        period = PERIOD_LINE.search(line)
        if period and not PERIOD_OK.match(period.group(1).strip()):
            out.hit("작성-표기", project, "사업기간",
                    f"사업기간은 '2027. 1. ~ 2027. 12.' 모양으로 씁니다. ({SOURCE}: 사업개요)",
                    line.strip()[:40])
        if EMPTY_COLON.search(line):
            out.hit("작성-표기", project, "목적콜론",
                    f"사업목적·사업근거를 '-' 로 나눠 적을 때는 제목 뒤 ':' 를 지웁니다. ({SOURCE}: 사업개요)",
                    line.strip()[:40])
        head = line[:1]
        if head in DIAMOND:
            rest = line[1:].strip()
            if FORBIDDEN_HEAD.match(rest):
                out.hit("작성-표기", project, "시기방법",
                        f"'◆시기', '◆방법' 항목은 쓰지 않습니다. 내용의 하위(▸)에 적습니다. ({SOURCE}: 사업계획)",
                        rest[:20])
        if "사업감액" in line:
            at = line.index("사업감액")
            out.hit("작성-용어", project, "사업감액",
                    f"미편성 사유에 '사업감액'은 쓰지 않습니다. 사업종료·사업완료·~으로 통합 운영 등으로 "
                    f"씁니다. ({SOURCE}: 본예산 미편성 사업)", _window(line, at, at + 4))


# ---------------------------------------------------------------- 금액 표기

MONEY = re.compile(r"(\d[\d,]*\d|\d)\s*(천원|원)")
GOOD_NUMBER = re.compile(r"^\d{1,3}(,\d{3})*$")


def _money_rules(block: _Block, out: _Collector) -> None:
    for line in block.lines:
        for found in MONEY.finditer(line):
            number = found.group(1)
            if "," in number and not GOOD_NUMBER.match(number):
                out.hit("작성-금액", block.project, "쉼표",
                        f"금액의 쉼표 자리가 맞지 않습니다({number}{found.group(2)}). 세 자리마다 찍습니다.",
                        _window(line, found.start(), found.end(), 2))
            elif "," not in number and len(number) >= 5:
                out.hit("작성-금액", block.project, "쉼표없음",
                        f"금액에 천 단위 쉼표가 없습니다({number}{found.group(2)}).",
                        _window(line, found.start(), found.end(), 2))


# ---------------------------------------------------------------- 산출식

FACTOR_UNIT = re.compile(r"[\d.,]+\s*([가-힣][가-힣 ]*)$")
RANK = {}
for _unit in ("명", "인", "교", "개교", "기관", "개기관", "개 기관", "팀", "개팀", "개청", "개원", "학급", "개학급"):
    RANK[_unit] = 1
RANK["시간"] = 2
RANK["일"] = 3
for _unit in ("회", "월"):
    RANK[_unit] = 4
RANK_NAME = {1: "명(교·기관)", 2: "시간", 3: "일", 4: "월(회)"}
BARE_COUNTER = {"교": "개교", "기관": "개 기관", "팀": "개팀"}


def _factors(formula: str) -> list[str]:
    body = formula.split("=")[0]
    return [part.strip() for part in re.split(r"[×xX*]", body) if part.strip()]


def unit_price(formula: str) -> Optional[float]:
    """산출식 맨 앞 단가(원). '160천원×2명' → 160000. 단가가 없으면 None."""
    factors = _factors(formula)
    if not factors:
        return None
    found = re.match(r"^([\d,]+(?:\.\d+)?)\s*(천원|원)", factors[0])
    if not found:
        return None
    value = float(found.group(1).replace(",", ""))
    return value * 1000 if found.group(2) == "천원" else value


def _formula_rules(project, index: int, item, out: _Collector) -> None:
    formula = item.formula or ""
    factors = _factors(formula)
    if not factors:
        return
    name = project.name
    snippet = formula.split("=")[0].strip()
    has_won = [("원" in factor) for factor in factors]
    if not any(has_won) and "%" not in formula:
        out.hit("작성-금액", name, "원없음",
                f"'{item.name}' 산출식에 단가 단위(원·천원)가 없습니다. 산출기초는 '원×명×회'로 씁니다. "
                f"({SOURCE}: 산출기초 순서)", snippet, index)
    elif any(has_won) and not has_won[0]:
        out.hit("작성-산출", name, "단가순서",
                f"'{item.name}' 산출식은 단가(원)를 맨 앞에 씁니다 — 단가×대상×횟수. ({SOURCE}: 산출기초 순서)",
                snippet, index)
    if "개월" in formula:
        out.hit("작성-산출", name, "개월",
                f"산출식 단위는 '개월'이 아니라 '월'입니다. ({SOURCE}: 산출기초 순서)", snippet, index)
    ranks = []
    for factor in factors[1:]:
        unit = FACTOR_UNIT.search(factor)
        if not unit:
            continue
        word = unit.group(1).strip()
        if word in BARE_COUNTER:
            out.hit("작성-산출", name, "단위",
                    f"산출식 단위 '{word}'는 '{BARE_COUNTER[word]}'로 씁니다. 예) ×10개교 ({SOURCE}: 산출식 단위 확인)",
                    snippet, index)
        if word in RANK:
            ranks.append(RANK[word])
    for before, after in zip(ranks, ranks[1:]):
        if after < before:
            out.hit("작성-산출", name, "순서",
                    f"'{item.name}' 산출식 순서가 다릅니다. '원×명(교, 기관)×시간×일×월(회)' 순서로 씁니다 — "
                    f"{RANK_NAME[after]}이 {RANK_NAME[before]} 뒤에 있습니다. ({SOURCE}: 산출기초 순서)",
                    snippet, index)
            break


# ---------------------------------------------------------------- 용어

TERMS = {
    "운영용품": ("행사용품", "운영비"),
    "협의회": ("협의회식비", "간담회"),
    "강사수당": ("강사비", "강사료"),
}


def _clean(name: str) -> str:
    name = re.sub(r"\(.*?\)$", "", name or "").strip()
    return "".join(char for char in name if not is_private_use(char)).strip(" ․·-")


def _term_rules(project, index: int, item, out: _Collector) -> None:
    if not item.formula:
        return
    name = _squeeze(_clean(item.name))
    if name == "운영비":
        # '운영비'는 학교·교육지원청에 주는 지원금(620-03 목적사업비 등) 이름으로도 쓴다.
        # 요령이 운영용품으로 통일하라는 것은 물품 쪽이다 — 일반수용비·물품 단가일 때만 본다.
        price = unit_price(item.formula)
        if item.bimok_code5 not in ("", "210-01") or price is None or price > 100000:
            return
    for right, wrongs in TERMS.items():
        if name in wrongs:
            out.hit("작성-용어", project.name, f"용어-{right}",
                    f"'{_clean(item.name)}'는 '{right}'로 통일합니다. ({SOURCE}: 용어통일)",
                    _clean(item.name), index)


# ---------------------------------------------------------------- 기준단가

EVENT_MEAL = ("학생식비", "학부모식비", "민간인식비")


def _price_rules(project, index: int, item, out: _Collector) -> None:
    if not item.formula:
        return
    price = unit_price(item.formula)
    name = _squeeze(_clean(item.name))
    snippet = item.formula.split("=")[0].strip()
    if price is None:
        return
    won = f"{price:,.0f}원"
    if name == "협의회" and price != 15000:
        out.hit("작성-단가", project.name, "협의회",
                f"협의회 단가는 가급적 15,000원으로 통일합니다. ({won}) ({SOURCE}: 기타 기준단가)",
                snippet, index, price / 1000, 15)
    elif name == "운영용품" and price not in (50000, 100000):
        out.hit("작성-단가", project.name, "운영용품",
                f"운영용품 단가는 50,000원 또는 100,000원입니다. ({won}) ({SOURCE}: 기타 기준단가)",
                snippet, index, price / 1000, 50)
    elif (name in EVENT_MEAL or (name == "식비" and item.bimok_code5 in ("210-12", "310-06"))) and price > 11000:
        out.hit("작성-단가", project.name, "행사식비",
                f"{_clean(item.name)} 기준단가는 1인 1식 11,000원입니다. ({won}) ({SOURCE}: 기타교육운영비·행사실비지원금)",
                snippet, index, price / 1000, 11)
    elif name == "간식비" and price > 3000:
        out.hit("작성-단가", project.name, "간식비",
                f"간식비 기준단가(3천원)보다 높습니다. ({won}) ({SOURCE}: 기타 기준단가)",
                snippet, index, price / 1000, 3)
    elif name.endswith("식비") and name not in EVENT_MEAL and "급식" not in name and "특근" not in name \
            and item.bimok_code5 not in ("210-12", "310-06") and price > 9000:
        out.hit("작성-단가", project.name, "식비",
                f"식비 기준단가(업무추진비 제외 9천원)보다 높습니다. ({won}) ({SOURCE}: 기타 기준단가)",
                snippet, index, price / 1000, 9)
    if "%" in item.formula and re.search(r"보험|부담금", item.name):
        rate = re.search(r"([\d.]+)\s*%", item.formula)
        if rate and abs(float(rate.group(1)) - 11.33) > 0.001:
            out.hit("작성-단가", project.name, "4대보험",
                    f"요령의 4대보험 요율(11.33%)과 다릅니다({rate.group(1)}%). 직종별로 요율이 다르면 그대로 "
                    f"두세요. ({SOURCE}: 공통 작성요령)", snippet, index)


# ---------------------------------------------------------------- 원가통계비목

EVENT_WORDS = re.compile(r"워크숍|연찬회|연수|회의|협의회|간담회|세미나")
MEAL_ALLOWED = re.compile(r"검정고시|수학능력|수능|학력평가|연수원")


def _bimok_rules(project, index: int, item, parents: list, out: _Collector) -> None:
    code = item.bimok_code5
    if not code or not item.formula:
        return
    name = _clean(item.name)
    context = " ".join([project.name] + parents + [name])
    snippet = name
    if code == "210-04":
        if "숙박" in name:
            out.hit("작성-비목", project.name, "급량숙박",
                    f"'{name}'이 급량비(210-04)에 있습니다. 워크숍·연수 등의 숙박비는 시설임차료(210-07)입니다. "
                    f"({SOURCE}: 급량비 편성)", snippet, index)
        elif EVENT_WORDS.search(context) and not MEAL_ALLOWED.search(context):
            out.hit("작성-비목", project.name, "급량행사",
                    f"'{name}'이 급량비(210-04)에 있습니다. 워크숍·연수·회의 등 행사성 식비·간식은 "
                    f"사업추진경비(230-02)입니다. ({SOURCE}: 급량비 편성)", snippet, index)
    squeezed = _squeeze(name)
    if squeezed == "학생식비" and code != "210-12":
        out.hit("작성-비목", project.name, "학생식비",
                f"학생 식비는 기타교육운영비(210-12)입니다. 지금 {code}. ({SOURCE}: 교육운영비 편성)", snippet, index)
    if squeezed in ("학부모식비", "민간인식비") and code != "310-06":
        out.hit("작성-비목", project.name, "민간식비",
                f"학부모·민간인 식비는 행사실비지원금(310-06)입니다. 지금 {code}. ({SOURCE}: 기타보전금)",
                snippet, index)
    if code == "320-01" and "사립" in context:
        out.hit("작성-비목", project.name, "사립보조",
                f"사립학교에 대한 보조금은 620목(사립학교목적사업비 등)에 편성합니다. 지금 민간경상보조(320-01). "
                f"({SOURCE}: 민간이전)", snippet, index)
    if code == "230-02" and "기관운영" in (project.policy or ""):
        out.hit("작성-비목", project.name, "기관운영업추",
                f"정책사업 '기관운영'에는 사업추진업무추진비(230-02)를 편성하지 않습니다. ({SOURCE}: 업무추진비)",
                snippet, index)


# ---------------------------------------------------------------- 같은 묶음 안의 인원·횟수

def _count_of(formula: str, units: Iterable[str]) -> Optional[str]:
    for factor in _factors(formula)[1:]:
        found = re.match(r"^([\d,]+)\s*(\S+)$", factor)
        if found and found.group(2) in units:
            return found.group(1) + found.group(2)
    return None


def _sibling_rules(project, out: _Collector) -> None:
    """같은 사업 안에서 인원·횟수는 맞춘다. 요령의 예: 50명인데 식비·간식비는 30명."""
    groups: dict = {}
    stack: list = []
    for index, item in enumerate(project.items):
        while stack and stack[-1][1].depth >= item.depth:
            stack.pop()
        parent = stack[-1][0] if stack else -1
        groups.setdefault(parent, []).append((index, item))
        stack.append((index, item))
    for parent, members in groups.items():
        meals = [(index, item) for index, item in members
                 if item.formula and _squeeze(_clean(item.name)) in ("식비", "간식비", "참가자식비")]
        if len(meals) < 2:
            continue
        people = {_count_of(item.formula, ("명",)) for _index, item in meals} - {None}
        if len(people) > 1:
            index, item = meals[0]
            out.hit("작성-산출", project.name, f"인원-{parent}",
                    f"같은 묶음의 식비·간식비 인원이 다릅니다({', '.join(sorted(people))}). 같은 사업 안에서는 "
                    f"인원수·횟수를 맞춥니다. ({SOURCE}: 공통 작성요령)",
                    " / ".join(one.formula.split("=")[0] for _i, one in meals)[:60], index)


# ---------------------------------------------------------------- 증감사유

def _numbers(cells: list[str]) -> list[Optional[float]]:
    values = []
    for cell in cells:
        text = cell.replace("천원", "").strip()
        negative = text.startswith("△") or text.startswith("-")
        text = text.lstrip("△-").replace(",", "")
        values.append((-1 if negative else 1) * float(text) if re.fullmatch(r"\d+(\.\d+)?", text) else None)
    return values


REASON = re.compile(r"(증감|증액|감액)사유")


def _any_change(rows) -> bool:
    """표 안 어느 줄이라도 본예산 대비 증감이 있는가. 2레벨 합은 같아도 3레벨끼리
    늘고 준 경우(+400 / −400)는 증감사유가 맞다."""
    for row in rows:
        values = [one for one in _numbers(row[1:]) if one is not None]
        if len(values) >= 4 and values[3]:
            return True
    return False


def _change_rules(blocks: list[_Block], out: _Collector) -> None:
    """사업계획 절마다 증감이 있으면 증감사유가 있어야 하고, 없으면 증감사유를 지운다.

    절의 본문과 그 절의 산출내역 표는 이어서 나온다. 사이에 계속비 연부액 표 같은 다른
    표가 끼기도 하므로, 앞 산출내역 표 이후의 글을 모두 그 절의 본문으로 본다.
    표 첫 줄('1. OOO', A, B, C, A−B, A−C)에서 본예산 대비 증감을 읽는다.
    '감액사유'·'증액사유'로 적었어도 사유는 있는 것으로 본다(표기는 따로 말한다).
    """
    body: list[str] = []
    project = None
    for block in blocks:
        if block.project != project:
            body, project = [], block.project
        if block.kind == "text":
            body.extend(block.lines)
            continue
        rows = block.rows
        if not any("산출내역" in "".join(row).replace(" ", "") for row in rows[:4]):
            continue
        first = next((row for row in rows if row and HEADING.match(row[0])), None)
        text = " ".join(body)
        body = []
        if first is None:
            continue
        values = [one for one in _numbers(first[1:]) if one is not None]
        if len(values) < 4:
            continue
        change = values[3]
        title = first[0].strip()
        has_reason = bool(REASON.search(text))
        if change and not has_reason:
            out.hit("작성-증감", block.project, f"사유없음-{title}",
                    f"'{title}'은 본예산 대비 {change:+,.0f}천원 증감이 있는데 증감사유가 없습니다. "
                    f"개조식으로 수치를 넣어 씁니다. ({SOURCE}: 증감사유 작성 방법)", title[:30],
                    left=values[0], right=values[1])
        elif not change and "증감사유" in text and not _any_change(rows):
            out.hit("작성-증감", block.project, f"사유불필요-{title}",
                    f"'{title}'은 본예산 대비 증감이 없는데 증감사유가 있습니다. "
                    f"'26년 대비 증감이 없으면 지웁니다. ({SOURCE}: 사업계획)", title[:30])
        for found in re.finditer(r"(증액|감액)사유", text):
            out.hit("작성-표기", block.project, "사유이름",
                    f"'{found.group()}'가 아니라 '증감사유'로 씁니다(항목 이름 고정). ({SOURCE}: 사업계획)",
                    found.group())


def _headline(blocks: list[_Block]) -> dict:
    """사업 머리 표('2027년 본예산안(A) | 2026년 본예산(B) | …')에서 {사업: (A, B)}."""
    found: dict = {}
    for block in blocks:
        if block.kind != "table" or block.project in found:
            continue
        flat = "".join(block.lines).replace(" ", "")
        if "본예산안(A)" not in flat or "본예산(B)" not in flat:
            continue
        for row in block.rows:
            values = [one for one in _numbers(row) if one is not None]
            if len(values) >= 2 and len(values) == len(row):
                found[block.project] = (values[0], values[1])
                break
    return found


def _growth_rules(project, headline: dict, out: _Collector) -> None:
    """신규 사업·30% 이상 증액 사업은 결재 파일을 붙인 사업계획서가 필요하다(본청)."""
    if project.name not in headline:
        return
    request, before = headline[project.name]
    if before == 0 and request > 0:
        out.hit("작성-증감", project.name, "신규",
                "신규 사업입니다. 사업편성부터 집행계획까지 3레벨 기준 사업계획서(소속 국장 결재 파일)를 "
                "붙입니다(본청). (편성 유의사항: 신규사업 및 30%이상 증액 사업)", project.heading or project.name,
                left=request, right=before)
    elif before > 0 and (request - before) / before >= 0.3:
        out.hit("작성-증감", project.name, "증액30",
                f"본예산 대비 {(request - before) / before:.0%} 증액입니다. 30% 이상 증액 사업은 사업계획서"
                f"(소속 국장 결재 파일)를 붙입니다(본청). (편성 유의사항: 신규사업 및 30%이상 증액 사업)",
                project.heading or project.name, left=request, right=before)


# ---------------------------------------------------------------- 작년 확정본과 비교

def _path_key(names: list[str]) -> tuple:
    from bimok_resolver import normalise

    return tuple(_squeeze(_clean(normalise(name))) for name in names)


def last_year_index(rows) -> dict:
    """작년 K-에듀파인 파일 → {(사업, …, 항목) 경로: (단가, 비목 5자리)}."""
    index: dict = {}
    stack: list = []
    for row in rows:
        if row.level is None:
            continue
        stack = stack[:row.level - 1] + [row.name]
        price = unit_price(row.basis) if row.basis else None
        code = row.code[:5]
        code5 = f"{code[:3]}-{code[3:5]}" if len(code) >= 5 else ""
        if price is not None or code5:
            index.setdefault(_path_key(stack), (price, code5))
    return index


def _last_year_rules(project, index: int, item, parents: list, last: dict, out: _Collector) -> None:
    if not last or not item.formula:
        return
    key = _path_key([project.name] + parents + [item.name])
    found = last.get(key)
    if not found:
        return
    before, code5 = found
    now = unit_price(item.formula)
    name = _clean(item.name)
    if before and now and now > before:
        out.hit("작성-단가", project.name, f"상향-{index}",
                f"'{name}' 단가가 작년 확정본보다 올랐습니다({before:,.0f}원 → {now:,.0f}원, "
                f"{(now - before) / before:+.0%}). 이유 없는 단가 상향은 전면 재검토 대상입니다 — 증감사유에 "
                f"근거(견적·물가 등)를 적었는지 확인하세요.", item.formula.split("=")[0].strip(), index,
                now / 1000, before / 1000)
    if code5 and item.bimok_code5 and code5 != item.bimok_code5:
        out.hit("작성-비목", project.name, f"비목변경-{index}",
                f"'{name}' 비목이 작년 확정본과 다릅니다(작년 {code5} → 올해 {item.bimok_code5}). "
                f"편성목·원가통계비목을 다시 확인하세요. ({SOURCE}: 편성목 확인)", name, index)


# ---------------------------------------------------------------- 2레벨 이름

def _name_rules(project, out: _Collector) -> None:
    for index, item in enumerate(project.items):
        if item.depth == 1 and _squeeze(_clean(item.name)) == "시도분담금":
            out.hit("작성-산출", project.name, "시도분담금",
                    f"본청은 2레벨에 포괄적인 사업명을 씁니다('시도분담금'으로 2레벨 작성 금지). "
                    f"({SOURCE}: 사업별 설명서 작성 철저)", _clean(item.name), index)


# ---------------------------------------------------------------- 진입점

def _parents(items, index: int) -> list[str]:
    names = []
    depth = items[index].depth
    for back in range(index - 1, -1, -1):
        if items[back].depth < depth:
            names.append(items[back].name)
            depth = items[back].depth
    return list(reversed(names))


def check_writing(plan_path: str, projects: list, last_year_rows=None) -> WritingReport:
    report = WritingReport()
    out = _Collector()
    blocks: list[_Block] = []
    if plan_path:
        try:
            blocks = _blocks(plan_path, projects)
        except Exception as error:  # noqa: BLE001 - 표기 점검 실패가 불러오기를 막으면 안 된다
            report.skipped.append(Issue("작성-표기", "건너뜀", "", "", f"문서 글을 읽지 못해 표기 점검을 건너뜀: {error}"))
    for block in blocks:
        _text_rules(block, out)
        _money_rules(block, out)
    _change_rules(blocks, out)
    last = last_year_index(last_year_rows or [])
    headline = _headline(blocks)
    for project in projects:
        _growth_rules(project, headline, out)
        _name_rules(project, out)
        _sibling_rules(project, out)
        for index, item in enumerate(project.items):
            parents = _parents(project.items, index)
            _formula_rules(project, index, item, out)
            _term_rules(project, index, item, out)
            _price_rules(project, index, item, out)
            _bimok_rules(project, index, item, parents, out)
            _last_year_rules(project, index, item, parents, last, out)
    order = list(CATEGORIES)
    report.issues = sorted(out.issues(), key=lambda issue: order.index(issue.goal))
    return report
