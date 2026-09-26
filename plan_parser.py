"""부서별사업별 설명서(HWPX)를 읽어 사업 트리로 만든다.

이 서식의 '□ 요구내용 및 산출근거' 표는 K-에듀파인 입력에 필요한 것을 모두 담고 있다.

    ․ 운영용품   일반수용비(210-01)   50,000원×5종×2회  =  500 천원
    ▸초과       강사수당(210-06)     90,000원×2명×1시간×2회 = 360 천원

  - 항목명, 산출식, 금액
  - 비목명과 목코드(210-01). 이 5자리가 K-에듀파인 원가통계비목의 앞 5자리다.
  - 단가가 이미 원 단위다. 천원 환산이 필요 없다.

같은 문서의 '산출내역' 표(사업계획 절)는 같은 금액을 천원 단가로 적은 별도 표다.
비목이 없으므로 입력원으로 쓰지 않고, 금액 교차검증에만 쓴다.
"""

from __future__ import annotations

import html
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import Optional

from hwpx_native_parser import is_private_use, native_chunks

BIMOK = re.compile(r"^(.+?)\((\d{3}-\d{2})\)$")
LEVEL1 = re.compile(r"^\d+\.$")
LEVEL2 = re.compile(r"^[①-⑳]$")
NUMBER = re.compile(r"^-?[\d,]+$")
HANDOVER_CODE = "620-02"      # 총액배분사업비. 예산과가 학교로 바로 재배정한다.
YEAR_REQUEST = re.compile(r"(예산안?|추경안?)\(A\)")   # 본예산안(A) · 본예산(A) · 추경안(A)
DOT_BULLETS = "․·ㆍ∙・"      # ․ · ㆍ ∙ ・
ARROW_BULLETS = "▸▶►"                # ▸ ▶ ►
HEADING = re.compile(r"^\s*(\d+)\.\s*(.+?)\s*$")
NOT_TITLE = "□❍◦●○※-–—·․ㆍ∙・"


def _squeeze_name(name: str) -> str:
    return re.sub(r"[\s·․ㆍ∙・]", "", name or "")


def _looks_like_title(line: str) -> bool:
    """조직 표 바로 앞에 놓인 '이 블록의 사업명' 줄인지.

    설명서는 한 사업을 총괄 한 쪽 + 세부 여러 쪽으로 나눠 적는다. 세부 쪽의 제목에는
    번호가 붙지 않는다. 번호 있는 줄만 제목으로 보면 세부 쪽이 전부 앞 사업의 이름을
    뒤집어쓰고, 같은 이름의 사업이 두세 개씩 생겨 UBIS 대조가 통째로 어긋난다.
    """
    if not (2 <= len(line) <= 80):
        return False
    if line[0] in NOT_TITLE or is_private_use(line[0]):
        return False
    if ":" in line or "：" in line:
        return False
    if NUMBER.match(line) or line.endswith("쪽 -"):
        return False
    # '62. 유치원 방과후 과정 운영 지원' 처럼 제목이 두 줄로 접히면 둘째 줄은
    # '(총액배분사업비)' 만 남는다. 그걸 제목으로 받으면 62번 사업이 통째로
    # 61번의 세부 쪽으로 빨려 들어간다. 괄호를 걷어낸 알맹이가 있어야 제목이다.
    core = re.sub(r"\([^)]*\)", "", line).strip()
    if len(core) < 3:
        return False
    return any("가" <= character <= "힣" for character in core)


@dataclass
class PlanItem:
    """산출근거 한 줄."""

    depth: int
    name: str
    bimok_name: str = ""
    bimok_code5: str = ""          # 210-01 형태
    formula: str = ""              # 50,000원×5종×2회
    amount: Optional[float] = None  # 천원
    row: int = 0                   # 설명서 표에서의 행 번호 (오류 안내용)
    leaf: Optional[bool] = None    # 산출내역 표는 산출식 없이 금액만 있는 잎이 있다

    @property
    def is_leaf(self) -> bool:
        """맨 아래 줄인가.

        '요구내용 및 산출근거' 표에서는 산출식이 있으면 잎이다. '산출내역' 표에는
        '① 초등학교 신설교과 현장 이해  7,000천원' 처럼 산출식 없이 금액만 있는 잎이
        있어서, 그쪽은 아래에 더 깊은 줄이 있는지로 판정해 `leaf` 에 적어 둔다.
        """
        return bool(self.formula) if self.leaf is None else self.leaf

    @property
    def handover(self) -> bool:
        """K-에듀파인에 직접 입력하지 않는 줄.

        총액배분사업비는 예산과가 학교로 바로 재배정한다. 이 과에서 입력하지 않으므로
        입력본에 넣으면 오입력이 되고, 비목을 확정하라고 조르면 담당자는 확정할 수
        없는 줄을 붙들게 된다.
        """
        return self.bimok_code5 == HANDOVER_CODE

    @property
    def code5(self) -> str:
        """목코드의 숫자만. 210-01 → 21001"""
        return self.bimok_code5.replace("-", "")


@dataclass
class PlanPart:
    """번호 없이 이어 붙은 세부 쪽.

    설명서는 한 사업을 '총괄 한 쪽 + 세부 여러 쪽'으로 적기도 하고, 번호 없는 별개
    사업들을 앞 번호 아래에 이어 붙이기도 한다. 어느 쪽인지 문서만 보고 단정할 수
    없어서, 산출근거는 한 트리로 합치고 요구액은 후보로만 남긴다.
    """

    name: str
    request: Optional[float] = None
    year_request: Optional[float] = None

    @property
    def base(self) -> Optional[float]:
        return self.year_request if self.year_request is not None else self.request



@dataclass
class PlanProject:
    """설명서의 사업 한 건."""

    name: str
    number: str = ""          # 설명서에 적힌 사업 번호. "15"
    order: int = 0            # 문서에 나온 차례. 번호와 다를 수 있다.
    heading: str = ""         # 원문 그대로. 한글에서 찾을 때 쓴다.
    policy: str = ""
    unit: str = ""
    program: str = ""
    organisation: str = ""
    request: Optional[float] = None       # 총사업비(천원). 사업 전체 금액이라 올해 요구액과 다를 수 있다.
    year_request: Optional[float] = None  # '2027년 본예산안(A)' 표의 올해 요구액(천원)
    funding: dict = field(default_factory=dict)
    items: list[PlanItem] = field(default_factory=list)
    parts: list = field(default_factory=list)      # 번호 없이 이어 붙은 세부 쪽
    outline_total: Optional[float] = None  # 산출내역 표(천원 단가)의 합계
    outline_tables: list = field(default_factory=list)   # 산출내역 표 원본(예비 산출근거)
    outline_items: list = field(default_factory=list)    # 산출내역 표를 읽은 항목. 표 하나가 한 묶음
    from_outline: bool = False        # 산출근거를 산출내역 표에서 읽었는가(= 비목이 없다)
    unnumbered: bool = False          # 번호 없는 제목으로 시작한 블록
    split_description: bool = False   # 설명서가 이 사업 언저리를 여러 쪽으로 나눠 적었다
    had_own_items: bool = False       # 자기 블록에서 산출근거를 읽었는가
    supplement: bool = False          # 산출내역 표가 추경 증감 서식이다(증감분만 적혀 있다)

    @property
    def item_total(self) -> float:
        return sum(item.amount or 0 for item in self.items if item.is_leaf)

    @property
    def charge(self) -> str:
        """'유초등특수교육과, 유보통합담당' → '유보통합담당'.

        프로그램을 과 전체가 나눠 쓰므로, 담당자가 자기 것만 걸러 볼 수 있어야 한다.
        담당이 여럿 적힌 사업은 적힌 대로 둔다(공동 사업이다).
        """
        _, _, rest = self.organisation.partition(",")
        return rest.strip() or self.organisation.strip()

    @property
    def handover(self) -> str:
        """이 과에서 K-에듀파인에 입력하지 않는 사업이면 그 사유. 아니면 빈 문자열.

        판별은 이름이 아니라 비목으로 한다. 표본에서 '특수학교 교육과정 운영지원'은
        이름에 (총액배분사업비)가 없는데 잎이 전부 620-02였다. 이름으로 걸렀으면
        이 사업 5행을 놓쳤을 것이다.
        """
        if "재원배분" in self.name:
            return "재원배분"
        leaves = [item for item in self.items if item.is_leaf]
        if leaves:
            return "총액배분 재배정" if all(item.handover for item in leaves) else ""
        # 산출근거가 아예 없는 사업은 비목으로 가릴 수 없다. 이름으로 본다.
        return "총액배분 재배정" if "총액배분" in self.name else ""

    @property
    def base(self) -> Optional[float]:
        """산출근거 합계·UBIS 요구액과 맞춰야 하는 값.

        머리글 '총사업비'는 여러 해의 합계가 아니다. **올해 그 사업의 사업비**이므로
        '2027년 본예산안(A)' 표의 요구액과 같아야 한다. 그래서 둘 중 아무 쪽이나 써도
        되지만, 표에 적힌 요구액(A)을 먼저 본다. 머리글은 손으로 적는 칸이라 잘못 적히는
        일이 있고(3,240,000 / 1,023,500), 그 어긋남 자체는 _check_request 가 오류로 말한다.
        여기서 요구액(A)을 우선하지 않으면 잘못 적힌 머리글이 기준이 되어 버린다.
        """
        return self.year_request if self.year_request is not None else self.request

    @property
    def amount_candidates(self) -> list[tuple[str, float]]:
        """UBIS 요구액과 맞춰 볼 수 있는 값. 설명서에 실제로 적힌 숫자만 쓴다.

        여러 쪽의 요구액을 더해 만든 합성값은 후보로 쓰지 않는다. 그런 값이 우연히
        무언가와 맞으면 조용히 통과하고, 안 맞으면 설명서 어디에도 없는 숫자를
        '설명서 금액'이라며 화면에 크게 띄우게 된다. 실제로 47번에서 15,680 에
        남의 사업 34,500 이 더해져 50,180 이 만들어졌고, 그걸 UBIS 와 견주어
        오류라고 말했다.
        """
        found: list[tuple[str, float]] = []
        if self.base is not None:
            found.append(("설명서 올해 요구액", self.base))
        if self.items and (self.base is None or abs(self.item_total - self.base) > 1):
            found.append(("산출근거 합계", self.item_total))
        return found

    @property
    def label(self) -> str:
        """'15. 영유아교육내실화 지원' — 담당자가 설명서에서 찾는 이름."""
        return f"{self.number}. {self.name}" if self.number else self.name

    def find_text(self, item: Optional[PlanItem] = None) -> str:
        """한글에서 Ctrl+F 로 붙여넣으면 바로 가는 문자열.

        hwpx 에는 쪽 번호가 없다. 쪽은 글꼴과 여백에 따라 렌더링될 때 정해지므로
        파일만 보고는 알 수 없다. 대신 문서에 그대로 있는 글자를 준다.
        """
        if item is not None and item.formula:
            return item.formula
        if item is not None:
            return item.name
        return self.heading or self.name


def _cells(table_html: str) -> list[list[str]]:
    rows = []
    for row in re.findall(r"<tr>(.*?)</tr>", table_html, re.S):
        rows.append([html.unescape(re.sub(r"<[^>]+>", "", cell)).strip()
                     for cell in re.findall(r"<td[^>]*>(.*?)</td>", row, re.S)])
    return rows


def _is_organisation(rows: list[list[str]]) -> bool:
    text = " ".join(" ".join(row) for row in rows[:2])
    return "(정책사업)" in text and "(단위사업)" in text and "(세부사업)" in text


def _organisation(rows: list[list[str]]) -> tuple[str, str, str, str]:
    text = " ".join(" ".join(row) for row in rows[:2])

    def grab(label: str) -> str:
        found = re.search(rf"\({label}\)\s*([^()]+)", text)
        return found.group(1).strip().rstrip(",") if found else ""

    return grab("정책사업"), grab("단위사업"), grab("세부사업"), grab("조직")


def _is_basis_table(rows: list[list[str]]) -> bool:
    """비목이 붙은 '요구내용 및 산출근거' 표인지."""
    for row in rows[:40]:
        for cell in row:
            if BIMOK.match(cell):
                return True
    return False


def _marker_depth(cells: list[str]) -> tuple[Optional[int], str, str, int]:
    """행의 확정 계층·불릿 문자·항목명·항목명 열 번호.

    확정 계층은 번호가 붙은 1·2단과 점 불릿 3단뿐이다. 그보다 깊은 단계는
    사설영역(PUA) 글리프라 문자만으로는 몇 단인지 알 수 없다. 여기서는 불릿
    문자만 돌려주고, 깊이는 문서 전체를 훑어 정한다(_learn_bullets).
    """
    marker_at = -1
    depth: Optional[int] = None
    for index, cell in enumerate(cells):
        if LEVEL1.match(cell):
            depth, marker_at = 1, index
            break
        if LEVEL2.match(cell):
            depth, marker_at = 2, index
            break
    for index in range(marker_at + 1, len(cells)):
        name = cells[index]
        if not name:
            continue
        bullet = ""
        if depth is None:
            first = name[0]
            if first in DOT_BULLETS:
                depth = 3
            elif first in ARROW_BULLETS or is_private_use(first):
                # 한글 문서의 깊은 단계 불릿은 표준 문자가 아니다. 이 문서에는
                # U+F02FB(4단)와 U+F077(5단)처럼 서로 다른 글리프가 섞여 있어,
                # 전부 같은 단계로 뭉뚱그리면 계층이 한 단 무너진다.
                bullet = first
            else:
                depth = 3
        return depth, bullet, _clean_name(name), index
    return None, "", "", -1


def _learn_bullets(tables: list[list[list[str]]]) -> dict[str, int]:
    """문서 전체를 훑어 불릿 글리프마다 몇 단인지 정한다.

    바로 앞 행의 단계 + 1을 후보로 모으고 가장 많이 나온 값을 택한다.
    한 번만 나타나는 글리프에 끌려다니지 않기 위해서다.
    """
    votes: dict[str, Counter] = defaultdict(Counter)
    for rows in tables:
        previous = 2
        tentative: dict[str, int] = {}
        for cells in rows:
            depth, bullet, name, _ = _marker_depth(cells)
            if not name:
                continue
            if depth is not None:
                previous = depth
                # 번호가 붙은 단계를 만나면 그보다 깊은 불릿 배정은 초기화한다.
                tentative = {mark: level for mark, level in tentative.items() if level <= depth}
            elif bullet:
                if bullet in tentative:
                    level = tentative[bullet]
                else:
                    level = min(previous + 1, 6)
                    tentative[bullet] = level
                # 같은 불릿이 연달아 나오는 것은 형제다. 깊이를 더하면 안 된다.
                votes[bullet][level] += 1
                previous = level
    return {bullet: counts.most_common(1)[0][0] for bullet, counts in votes.items()}


def _clean_name(name: str) -> str:
    while name and (name[0] in DOT_BULLETS or name[0] in ARROW_BULLETS or is_private_use(name[0])):
        name = name[1:]
    return name.strip()


def _parse_basis_rows(rows: list[list[str]], bullets: dict[str, int]) -> list[PlanItem]:
    items: list[PlanItem] = []
    for number, cells in enumerate(rows, start=1):
        if not any(cells):
            continue
        bimok_at = next((index for index, cell in enumerate(cells) if BIMOK.match(cell)), -1)
        depth, bullet, name, name_at = _marker_depth(cells)
        if depth is None and bullet:
            depth = bullets.get(bullet, 4)
        if not name or name_at == bimok_at:
            continue
        # 표 맨 끝 합계 줄은 앞 칸이 전부 비고 금액만 남는다. 그 금액이 항목명으로
        # 잡히면 '33,400'이라는 이름의 빈 자식이 생기고, 부모의 하위 합계가 0이 되어
        # 멀쩡한 사업이 금액 오류로 지적된다. (2027 설명서 67번에서 실제로 났다)
        if NUMBER.match(name) or name in ("계", "합계", "소계", "총계"):
            continue
        if any(keyword in name for keyword in ("산출내역", "요구액", "단위 :", "구    분")) and bimok_at < 0 and depth is None:
            continue
        item = PlanItem(depth=depth or 3, name=name, row=number)
        if bimok_at >= 0:
            found = BIMOK.match(cells[bimok_at])
            item.bimok_name, item.bimok_code5 = found.group(1).strip(), found.group(2)
            after = cells[bimok_at + 1:]
        else:
            after = cells[name_at + 1:]
        for cell in after:
            if "×" in cell or re.search(r"[\d,]+\s*원\s*$", cell):
                item.formula = cell.strip()
                break
        amounts = [cell for cell in after if NUMBER.match(cell)]
        if amounts:
            item.amount = float(amounts[-1].replace(",", ""))
        items.append(item)
    return items


def _year_request(rows: list[list[str]]) -> Optional[float]:
    """사업개요 바로 뒤 '2027년 본예산안(A) · 2026년 본예산(B) · …' 표의 (A) 값.

    머리글이 있는 열을 찾아 그 아래 숫자를 읽는다. 열 위치로 세면 연도가 바뀌어
    열이 늘거나 줄 때 엉뚱한 값을 읽는다.
    """
    for index, cells in enumerate(rows):
        for column, cell in enumerate(cells):
            if not YEAR_REQUEST.search(cell.replace(" ", "")):
                continue
            for below in rows[index + 1:]:
                if column < len(below) and NUMBER.match(below[column].strip()):
                    return float(below[column].replace(",", ""))
            return None
    return None


OUTLINE_LEVEL1 = re.compile(r"^(\d+)\.\s*(.+)$")
OUTLINE_LEVEL2 = re.compile(r"^([①-⑳])\s*(.+)$")
OUTLINE_AMOUNT = re.compile(r"^([\d,]+)\s*천원$")
FORMULA_TAIL = re.compile(r"^(.*?)\s*\(([^()]*[×xX][^()]*)\)\s*$")


def _fold_unnumbered(projects: list) -> list:
    """번호 없는 쪽을 앞 사업에 붙일지, 별개 사업으로 둘지 정한다.

    설명서는 번호 없는 쪽을 두 가지로 쓴다.

      67. 장애영유아교육지원        ← 총괄 쪽(산출근거 없음)
          특수학교 유-보-초 …       ← 세부 쪽(산출근거 있음)  → 붙인다
          특수학교(유치원) 방학 …   ← 세부 쪽(산출근거 있음)  → 붙인다

      47. 유치원 입학관리시스템 연수
          유치원 이음교육 활성화 지원 ← **다른 사업**(산출근거 없음) → 뗀다

    산출근거를 들고 온 쪽만 앞 사업의 일부다. 산출근거 없이 이름만 다른 쪽까지
    붙이면 남의 요구액이 그 사업 금액에 섞인다.
    """
    folded: list = []
    host = None
    for project in projects:
        if not project.unnumbered:
            folded.append(project)
            host = project
            continue
        if project.items and host is not None and not host.had_own_items:
            host.items.extend(project.items)
            host.outline_tables.extend(project.outline_tables)
            host.parts.append(PlanPart(name=project.name, request=project.request,
                                       year_request=project.year_request))
            host.split_description = True
            continue
        if host is not None:
            host.split_description = True
        folded.append(project)
    for order, project in enumerate(folded, start=1):
        project.order = order
    return folded


def _is_outline_table(rows: list[list[str]]) -> bool:
    """사업계획 절의 '산출내역' 표인지."""
    return any("산출내역" in "".join(row).replace(" ", "") for row in rows[:3])


def _is_supplement_table(rows: list[list[str]]) -> bool:
    """추경(추가경정예산) 설명서의 산출내역 표인지.

    추경 표는 '추가경정예산안(A) · 기정예산액(B) · 비교증감(A-B)' 머리에 **증감분만**
    적는다(`․ 개인당(△740천원×7명×6/12개월)  △2,590천원`). 이것을 산출근거로 읽으면
    사업 전체 금액과 견줄 수 없는 조각만 남아, 멀쩡한 사업이 '금액이 비어 있다'
    '합계가 0이다'라는 오류로 가득 찬다. 본예산 표는 '예산(안) · 요구액(A)'이다.
    """
    head = "".join("".join(row) for row in rows[:4]).replace(" ", "")
    return "추가경정" in head or "기정예산" in head


def _outline_marker(text: str) -> tuple[Optional[int], str, str]:
    """산출내역 표 한 칸에서 (확정 계층, 불릿 문자, 항목명)을 뜯는다.

    이 표는 '요구내용 및 산출근거' 표와 달리 글머리표·항목명·산출식이 한 칸에
    붙어 있다.  `․ 운영용품(50천원×5종×2회)`
    """
    cleaned = (text or "").strip()
    if not cleaned:
        return None, "", ""
    found = OUTLINE_LEVEL1.match(cleaned)
    if found:
        return 1, "", found.group(2).strip()
    found = OUTLINE_LEVEL2.match(cleaned)
    if found:
        return 2, "", found.group(2).strip()
    head = cleaned[0]
    if head in DOT_BULLETS or head in ARROW_BULLETS or is_private_use(head):
        return None, head, cleaned[1:].strip()
    return None, "", ""


def _learn_outline_bullets(tables: list[list[list[str]]]) -> dict[str, int]:
    """산출내역 표의 불릿이 몇 단인지 문서 전체를 훑어 정한다.

    한글의 깊은 단계 글머리표는 사설영역 글리프라 문자만으로는 단계를 알 수 없다.
    바로 위 줄보다 한 단 깊다고 보되, 같은 불릿이 연달아 나오면 형제이므로
    깊이를 더하지 않는다.
    """
    votes: dict[str, Counter] = defaultdict(Counter)
    for rows in tables:
        previous, tentative = 0, {}
        for cells in rows:
            depth, bullet, name = _outline_marker(cells[0] if cells else "")
            if not name:
                continue
            if depth is not None:
                previous = depth
                tentative = {mark: level for mark, level in tentative.items() if level <= depth}
            elif bullet:
                level = tentative.get(bullet) or min(previous + 1, 6)
                tentative[bullet] = level
                votes[bullet][level] += 1
                previous = level
    return {bullet: counts.most_common(1)[0][0] for bullet, counts in votes.items()}


def _parse_outline_rows(rows: list[list[str]], bullets: dict[str, int]) -> list[PlanItem]:
    """산출내역 표를 산출근거 항목으로 읽는다.

    금액은 두 곳에 있다. 자기 금액은 `500천원` 처럼 '천원'이 붙고, 그 뒤는 연도별
    예산 열이라 숫자만 있다. 상위 줄은 자기 금액 칸이 없고 요구액(A) 열이 곧 금액이다.
    """
    items: list[PlanItem] = []
    started = False
    for number, cells in enumerate(rows, start=1):
        if not started:
            started = any("산출내역" in cell.replace(" ", "") for cell in cells)
            continue
        depth, bullet, name = _outline_marker(cells[0] if cells else "")
        if not name:
            continue
        if depth is None:
            depth = bullets.get(bullet, 3)
        item = PlanItem(depth=depth, name=name, row=number)
        found = FORMULA_TAIL.match(name)
        if found and found.group(1).strip():
            item.name, item.formula = found.group(1).strip(), found.group(2).strip()
        for cell in cells[1:]:
            marked = OUTLINE_AMOUNT.match(cell.strip())
            if marked:
                item.amount = float(marked.group(1).replace(",", ""))
                break
        if item.amount is None:
            for cell in cells[1:]:
                if NUMBER.match(cell.strip()):
                    item.amount = float(cell.replace(",", ""))
                    break
        items.append(item)
    for index, item in enumerate(items):
        deeper = index + 1 < len(items) and items[index + 1].depth > item.depth
        item.leaf = not deeper
    return items


def _outline_total(rows: list[list[str]]) -> Optional[float]:
    """산출내역 표 첫 데이터 행의 요구액(A). 사업계획 절 합계 확인용."""
    for cells in rows[2:]:
        for cell in cells[1:]:
            if NUMBER.match(cell):
                return float(cell.replace(",", ""))
    return None


def parse_plan(path: str) -> list[PlanProject]:
    """설명서 HWPX를 사업 목록으로 읽는다."""
    chunks = native_chunks(path)
    basis_tables = [_cells(chunk["text"]) for chunk in chunks
                    if chunk.get("type") == "table" and _is_basis_table(_cells(chunk["text"]))]
    bullets = _learn_bullets(basis_tables)
    outline_tables = [_cells(chunk["text"]) for chunk in chunks
                      if chunk.get("type") == "table" and _is_outline_table(_cells(chunk["text"]))]
    outline_bullets = _learn_outline_bullets(outline_tables)

    projects: list[PlanProject] = []
    current: Optional[PlanProject] = None
    part_title = ""
    heading = ""
    heading_number = ""
    heading_line = ""

    for chunk in chunks:
        text = chunk.get("text") or ""
        if chunk.get("type") == "text":
            lines = [line.strip() for line in text.splitlines() if line.strip()]
            accepted = ""
            for line in lines:
                found = HEADING.match(line)
                if found:
                    heading_number, heading = found.group(1), found.group(2)
                    heading_line = accepted = line
                    part_title = ""
            # 번호 없는 세부 쪽 제목은 조직 표 바로 앞 줄에만 온다. 본문 아무 줄이나
            # 받으면 사업계획의 목록 줄이 사업명이 되어 버린다.
            if lines and lines[-1] != accepted and _looks_like_title(lines[-1]):
                part_title = lines[-1]
            if current is not None:
                total = re.search(r"총사업비\s*:\s*([\d,]+)\s*천원", text)
                if total:
                    amount = float(total.group(1).replace(",", ""))
                    if current.request is None:
                        current.request = amount
                        current.funding = {
                            label: float(value.replace(",", ""))
                            for label, value in re.findall(r"([가-힣]+)\s*([\d,]+)\s*천원",
                                                          text[total.end():])
                        }
            continue

        rows = _cells(text)
        if not rows:
            continue
        if _is_organisation(rows):
            policy, unit, program, organisation = _organisation(rows)
            title = part_title
            if not title and current is not None and heading and any(
                    _squeeze_name(one.name) == _squeeze_name(heading) for one in projects):
                # 이름은 같은데 번호가 다시 붙어 나왔다. 사업 안에서 1부터 다시 매긴
                # 세부 쪽이다. (61번 아래의 '1. · 2. · 3.' 쪽들) 새 사업으로 세면
                # 같은 이름의 사업이 둘이 되고, UBIS 요구액과 나눠 비교돼 오류가 난다.
                title = heading
            # 번호가 없어도 일단 사업으로 만든다. 앞 사업에 붙일지 별개로 둘지는
            # 산출근거를 들고 왔는지 보고 다 읽은 뒤에 정한다(_fold_unnumbered).
            current = PlanProject(name=title or heading or "사업명 확인 필요",
                                  number=heading_number, order=len(projects) + 1,
                                  heading=heading_line, policy=policy, unit=unit,
                                  program=program, organisation=organisation,
                                  unnumbered=bool(title))
            part_title = ""
            projects.append(current)
            continue
        if current is None:
            continue
        if _is_basis_table(rows):
            current.items.extend(_parse_basis_rows(rows, bullets))
            current.had_own_items = True
        elif _is_outline_table(rows):
            if _is_supplement_table(rows):
                current.supplement = True
                continue
            current.outline_tables.append(rows)
            if current.outline_total is None:
                current.outline_total = _outline_total(rows)
        elif len(rows) <= 4:
            found = _year_request(rows)
            if found is None:
                continue
            if current.year_request is None:
                current.year_request = found

    # 2027 설명서처럼 '요구내용 및 산출근거' 표가 아직 비어 있으면 산출내역 표로 읽는다.
    # 항목·계층·산출식·금액이 다 있고 비목만 없다. 비목이 없다고 통째로 버리면
    # 금액 검산(목표 1)이 아무 일도 하지 못한다.
    projects = _fold_unnumbered(projects)
    for project in projects:
        # 산출내역 표는 **항상** 읽어 둔다. 예전에는 '요구내용 및 산출근거' 표가
        # 없을 때만 읽었다. 그래서 두 서식이 같은 항목에 서로 다른 금액을 들고
        # 있어도 프로그램이 아무 말을 하지 않았다. 담당자가 한쪽만 고치는 일이
        # 실제로 있으니(3,060 을 3,600 으로) 두 표를 다 들고 있다가 마주 놓고 본다.
        project.outline_items = [_parse_outline_rows(rows, outline_bullets)
                                 for rows in project.outline_tables]
        if project.items or not project.outline_items:
            continue
        for block in project.outline_items:
            project.items.extend(block)
        project.from_outline = bool(project.items)
    return projects
