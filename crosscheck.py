"""세 가지 목표를 검사한다. 값을 고치지 않고 지적만 한다.

  목표 1  설명서 안에서의 금액 오류
  목표 2A 설명서 = UBIS 세출요구 = **생성한 입력본**      (입력 전)
  목표 2B 설명서 = UBIS 세출요구 = **입력 후 다운로드본**  (입력 후)
  목표 3  비목 확정 상태

2A와 2B는 시점이 다르다. 생성본과 다운로드본은 같이 존재할 수 없으므로 따로 돌린다.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional

from bimok_resolver import normalise
from edufine_form import read_form

NUMBER_UNIT = re.compile(r"([\d,]+(?:\.\d+)?)\s*(천원|원|%)?")


def calc(formula: str) -> Optional[float]:
    """산출식을 천원 단위로 계산한다.

    '=' 뒤의 표기 합계는 산출 요소가 아니다. 다시 곱하면 검산 자체가 틀어지므로
    좌변만 계산한다. 설명서의 단가는 원 단위이고, 금액 열은 천원이다.
    """
    if not formula:
        return None
    body = formula.split("=")[0]
    parts = [part for part in re.split(r"[×xX*]", body) if part.strip()]
    if not parts:
        return None
    product = 1.0
    for index, part in enumerate(parts):
        found = NUMBER_UNIT.search(part.replace(" ", ""))
        if not found:
            return None
        value = float(found.group(1).replace(",", ""))
        unit = found.group(2)
        if unit == "%":
            value /= 100
        elif index == 0:
            if unit == "천원":
                value *= 1000
            elif unit != "원":
                return None
        product *= value
    return product / 1000


@dataclass
class Issue:
    goal: str
    severity: str        # 오류 / 확인 필요 / 안내
    project: str
    item: str = ""
    message: str = ""
    left: Optional[float] = None
    right: Optional[float] = None
    index: Optional[int] = None      # 설명서 항목 목록에서의 자리. 같은 이름이 여럿이라 이름으로는 못 짚는다.
    form: str = ""                   # 어느 서식에서 난 지적인가. "" 는 요구내용 및 산출근거 표
    block: int = 0                   # 그 서식의 몇 번째 표인가
    row: int = 0                     # 그 표에서의 줄 번호
    blocking: bool = True            # 이 오류가 입력본의 값을 틀리게 만드는가

    @property
    def gap(self) -> Optional[float]:
        if self.left is None or self.right is None:
            return None
        return self.left - self.right


SKIPPED = "건너뜀"      # 검사를 하지 못했다. 목록에는 싣지 않고 몇 건인지만 센다.


@dataclass
class Report:
    """지적 목록. **오류만 싣는다** (2.15.0).

    예전에는 '확인 필요'와 '안내'도 실었다. 담당자는 오류가 아닌 줄까지 하나씩 열어
    봐야 했고, 정작 오류가 그 사이에 묻혔다. 이제 검사하는 자리에서 추론해 오류인지
    아닌지 정하고, 오류가 아니면 싣지 않는다.

    다만 **검사를 하지 못한 것**은 버리지 않는다. 오류 0건이 '다 맞다'인지 '못 봤다'인지
    구별되지 않으면 조용한 실패가 된다. 그래서 `skipped` 에 세어 두고 화면 한 줄로 말한다.
    `dropped` 는 오류가 아니라고 판단해 뺀 것 — 테스트와 원인 추적용이다.
    """

    issues: list[Issue] = field(default_factory=list)
    skipped: list = field(default_factory=list)
    dropped: list = field(default_factory=list)

    def add(self, *args, **kwargs) -> None:
        issue = Issue(*args, **kwargs)
        if issue.severity == "오류":
            self.issues.append(issue)
        elif issue.severity == SKIPPED:
            self.skipped.append(issue)
        else:
            self.dropped.append(issue)

    def extend(self, other: "Report") -> None:
        self.issues.extend(other.issues)
        self.skipped.extend(other.skipped)
        self.dropped.extend(other.dropped)

    @property
    def errors(self) -> list[Issue]:
        return [issue for issue in self.issues if issue.severity == "오류"]

    @property
    def blocking_errors(self) -> list[Issue]:
        """입력본을 틀리게 만드는 오류만.

        머리글 총사업비처럼 **입력본에 실리지 않는 칸**의 오류는 설명서에서 고쳐야 하지만,
        그 때문에 입력본 생성을 막으면 담당자는 고칠 권한도 없는 남의 사업 한 칸 때문에
        하루를 잃는다. 알리되 막지는 않는다.
        """
        return [issue for issue in self.errors if issue.blocking]

    @property
    def reviews(self) -> list[Issue]:
        return [issue for issue in self.issues if issue.severity == "확인 필요"]

    def of(self, goal: str) -> list[Issue]:
        return [issue for issue in self.issues if issue.goal.startswith(goal)]


# ---------------------------------------------------------------- 위계 읽기

TOLERANCE = 1.0          # 천원. 반올림 차이만 눈감는다.


def children_of(items, index: int) -> list[int]:
    """바로 아래 단계 자식들의 자리."""
    depth = items[index].depth
    found = []
    for position in range(index + 1, len(items)):
        if items[position].depth <= depth:
            break
        if items[position].depth == depth + 1:
            found.append(position)
    return found


def subtree_of(items, index: int) -> list[int]:
    """이 항목 아래 전부의 자리.

    글머리표를 한 단계 잘못 읽으면 '바로 아래 자식'은 통째로 바뀌지만, 이 범위와
    그 안의 잎은 그대로다. 그래서 합계를 두 방식으로 재 보고 서로 다른 답이 나올 때
    '오류'라고 단정하지 않는다.
    """
    depth = items[index].depth
    end = index + 1
    while end < len(items) and items[end].depth > depth:
        end += 1
    return list(range(index + 1, end))


def leaf_total(items, positions) -> float:
    return sum(items[position].amount or 0 for position in positions if items[position].is_leaf)


# ---------------------------------------------------------------- 목표 1

def _check_leaves(project, report: Report, items=None, form: str = "",
                  skip=frozenset(), block: int = 0) -> None:
    """산출식 = 금액.

    items 를 주면 그 목록을 본다. 설명서는 같은 사업의 산출 내역을 두 서식으로 적어서
    ('요구내용 및 산출근거' 표와 사업계획의 '산출내역' 표) 양쪽을 같은 잣대로 검산해야
    한다. 자리(index)는 설명서 항목 목록에서의 자리이므로 다른 서식에는 붙이지 않는다.
    """
    rows = project.items if items is None else items
    own = items is None
    for index, item in enumerate(rows):
        if not item.is_leaf or id(item) in skip:
            continue
        if not item.formula:
            # 산출내역 표에는 산출식 없이 금액만 적힌 잎이 있다. 없는 식을 두고
            # '계산할 수 없다'고 지적하면 목록이 그런 줄로 가득 찬다.
            continue
        computed = calc(item.formula)
        at = index if own else None
        mark = dict(index=at, form=form, block=block, row=item.row)
        if computed is None:
            report.add("목표1", SKIPPED, project.name, item.name,
                       f"산출식을 계산할 수 없습니다: {item.formula}", **mark)
        elif item.amount is None:
            report.add("목표1", "오류", project.name, item.name, "금액이 비어 있습니다.", **mark)
        elif abs(computed - item.amount) > TOLERANCE:
            label = form or "설명서"
            where = f"{label} {item.row}행 '{item.name}'" if item.row else f"'{item.name}'"
            report.add("목표1", "오류", project.name, item.name,
                       f"{where} — 산출식 {item.formula} 을 계산하면 {computed:,.0f}천원인데 "
                       f"적힌 금액은 {item.amount:,.0f}천원입니다 "
                       f"(차이 {abs(item.amount - computed):,.0f}천원).",
                       item.amount, computed, **mark)


def _check_shape(project, report: Report) -> None:
    """합계를 따지기 전에 트리 모양부터 본다.

    단계를 잘못 읽으면 멀쩡한 금액도 틀린 것처럼 보인다. 모양이 의심스러우면
    '오류'가 아니라 '확인 필요'로 말한다. 틀렸다고 단정하는 것보다 어디를 볼지
    알려주는 편이 담당자에게 쓸모 있다.
    """
    items = project.items
    previous = 0
    for index, item in enumerate(items):
        if previous and item.depth > previous + 1:
            report.add("목표1", "확인 필요", project.name, item.name,
                       f"윗 줄은 {previous}단인데 이 줄이 {item.depth}단입니다. 글머리표를 한 단계 "
                       "잘못 읽었을 수 있어 이 사업의 합계 검사는 믿지 마세요.", index=index)
        previous = item.depth
        if item.amount is not None and not item.is_leaf and not subtree_of(items, index):
            # 산출식도 아랫줄도 없는 금액 줄은 입력본에 산출기초 없이 실리고 합계에서도 빠진다.
            report.add("목표1", "오류", project.name, item.name,
                       f"산출식도 하위 항목도 없이 금액 {item.amount:,.0f}천원만 있는 줄입니다. "
                       "산출근거 합계에서 빠지고 입력본에 산출기초가 비게 됩니다. 산출식을 적어 주세요.",
                       item.amount, 0, index=index)


def _check_rollup(project, report: Report, rows=None, form: str = "",
                  skip=frozenset(), block: int = 0) -> bool:
    """부모 금액 = 아래 합계. 깊은 곳부터 본다.

    두 가지를 지킨다.

    하나. 아래가 어긋나면 위도 따라서 어긋난다. 둘 다 지적하면 한 군데 잘못이
    열 건으로 불어나 담당자가 어디를 고쳐야 할지 잃는다. 가장 깊은 자리만 말한다.

    둘. '바로 아래 단계 합계'가 안 맞아도 '맨 아래 산출근거 합계'가 맞으면, 틀린
    것은 설명서가 아니라 우리가 나눈 단계다. 그때는 오류라고 하지 않는다.
    """
    items = project.items if rows is None else rows
    own = rows is None
    label = form or "설명서"
    broken: set[int] = {index for index, item in enumerate(items) if id(item) in skip}
    for index in range(len(items) - 1, -1, -1):
        parent = items[index]
        children = children_of(items, index)
        if parent.amount is None or not children or index in broken:
            # index in broken: 이 줄은 이미 다른 검사가 짚었다(두 서식 대조). 같은
            # 잘못을 합계 어긋남으로 한 번 더 세면 한 군데 실수가 두 건으로 불어난다.
            continue
        inside = subtree_of(items, index)
        if any(position in broken for position in inside):
            continue
        direct = sum(items[position].amount or 0 for position in children)
        if abs(direct - parent.amount) <= TOLERANCE:
            continue
        at = index if own else None
        leaves = leaf_total(items, inside)
        where = f"{label} {parent.row}행 '{parent.name}'" if parent.row else f"'{parent.name}'"
        total = f"사업 총액 {project.base:,.0f}천원" if project.base is not None else "사업 총액 미상"
        if (project.request is not None and project.base is not None
                and abs(parent.amount - project.request) <= TOLERANCE
                and abs(direct - project.base) <= TOLERANCE):
            # 머리글 총사업비와 똑같은 값이 이 줄의 요구액 칸에도 적힌 쪽이다. 두 칸이
            # 같이 잘못 적힌 것이라 뿌리는 하나다(_check_request 가 오류로 말한다).
            # 같은 잘못을 두 번 오류로 세지 않고, 이 줄도 같이 고쳐야 한다고만 알린다.
            report.add("목표1", "확인 필요", project.name, parent.name,
                       f"{where} — 이 줄에 머리글 총사업비와 같은 {project.request:,.0f}천원이 "
                       f"적혀 있습니다. 아래를 더하면 {project.base:,.0f}천원입니다. "
                       "머리글 총사업비와 이 칸을 같이 고쳐 주세요.",
                       parent.amount, direct, index=at, form=form, block=block, row=parent.row)
            continue
        if abs(leaves - parent.amount) <= TOLERANCE:
            report.add("목표1", "확인 필요", project.name, parent.name,
                       f"[{total}] {where} — 바로 아래 단계만 더하면 {direct:,.0f}천원이지만, "
                       f"맨 아래 산출근거를 모두 더하면 {leaves:,.0f}천원으로 이 항목 금액과 같습니다. "
                       "금액은 맞고 단계를 나눠 읽은 방식만 다를 수 있습니다.",
                       parent.amount, leaves, index=at, form=form, block=block, row=parent.row)
            continue
        broken.add(index)
        report.add("목표1", "오류", project.name, parent.name,
                   f"[{total}] {where} 에서 어긋납니다. 이 항목은 {parent.amount:,.0f}천원인데 "
                   f"바로 아래 단계를 더하면 {direct:,.0f}천원입니다 "
                   f"(차이 {abs(parent.amount - direct):,.0f}천원 · 맨 아래 산출근거 합계는 {leaves:,.0f}천원).",
                   parent.amount, direct, index=at, form=form, block=block, row=parent.row)
    return bool(broken) and own


def _check_total(project, report: Report, quiet: bool = False) -> None:
    """올해 요구액 = 산출근거 합계.

    설명서에는 금액이 두 곳에 적힌다. 머리의 '총사업비'와 '본예산안(A)' 표의 올해
    요구액이다. 여러 해에 걸친 사업이나 재원이 갈린 사업에서는 이 둘이 다르고,
    그건 설명서의 잘못이 아니다. 둘 중 **어느 쪽과도** 맞지 않을 때만 오류다.
    """
    base = project.base
    if base is None:
        report.add("목표1", SKIPPED, project.name, "", "올해 요구액도 총사업비도 찾지 못했습니다.")
        return
    if not project.items:
        if not project.handover:
            report.add("목표1", "안내", project.name, "",
                       "요구내용 및 산출근거 표가 없습니다. 재원배분 사업인지 확인해 주세요.", base, 0)
        return

    leaves = project.item_total
    if abs(base - leaves) <= TOLERANCE:
        return

    if quiet:
        # 안쪽 어느 줄이 어긋났는지 이미 짚었다. 그 결과인 이 차이를 또 오류라고
        # 세면 한 군데 잘못이 두 건이 된다. 숫자는 상세 화면의 합계 줄에 그대로 있다.
        return

    # 다른 증인에게 물어본다. 하나라도 산출근거 합계와 맞으면 '읽는 기준'의 문제다.
    # 총사업비는 증인이 아니다. 올해 그 사업의 사업비라 요구액(A)과 같아야 하는 값이고,
    # 어긋나면 _check_request 가 오류로 말한다. 여기서 증인으로 세우면 잘못 적힌 총사업비가
    # 잘못된 쪽에 손을 들어 주고, 요구액이 멀쩡한 사업이 '기준이 갈렸다'로 넘어간다.
    others = {"산출내역 표 합계": project.outline_total}
    agree = [name for name, value in others.items()
             if value is not None and abs(value - leaves) <= TOLERANCE]
    others["총사업비"] = project.request
    others["1단 항목 합계"] = sum(item.amount or 0 for item in project.items if item.depth == 1)
    spoken = " · ".join(f"{name} {value:,.0f}천원" for name, value in others.items()
                        if value is not None)
    gap = abs(base - leaves)

    # 표 하나에 여러 사업 몫이 함께 실린 쪽인지 본다. 1단 항목 하나가 요구액과 똑같으면
    # 그 항목이 이 사업 몫이고 나머지는 남의 사업이다. 실제로 총액배분 사업 한 쪽에
    # 세 사업의 산출근거가 함께 실려 있었다(105,000 + 893,500 + 25,000 = 1,023,500).
    mine = [item.name for item in project.items
            if item.depth == 1 and item.amount is not None
            and abs(item.amount - base) <= TOLERANCE]
    if mine:
        report.add("목표1", "확인 필요", project.name, "",
                   f"한글 설명서 올해 요구액 {base:,.0f}천원 · 이 쪽의 산출근거를 모두 더하면 "
                   f"{leaves:,.0f}천원 — {gap:,.0f}천원 차이. 다만 1단 항목 "
                   f"'{mine[0]}' 하나가 {base:,.0f}천원으로 요구액과 같습니다. "
                   "한 표에 다른 사업 몫까지 함께 실린 것 같습니다. "
                   f"이 사업 몫이 그 항목뿐인지 확인해 주세요. ({spoken})",
                   base, leaves)
        return
    if agree:
        # 두 표(요구내용 및 산출근거 · 사업계획 산출내역)가 같은 합계를 말하는데 요구액(A)만
        # 다르다. 요구액 한 칸이 틀린 것이다. 그대로 두면 UBIS·입력본과도 어긋난다.
        report.add("목표1", "오류", project.name, "",
                   f"한글 설명서 올해 요구액 {base:,.0f}천원 · 산출근거를 모두 더하면 {leaves:,.0f}천원 "
                   f"— {gap:,.0f}천원 차이. {', '.join(agree)}도 산출근거와 같으므로 요구액(A) 한 칸이 "
                   f"틀렸습니다. ({spoken})",
                   base, leaves)
        return
    report.add("목표1", "오류", project.name, "",
               f"한글 설명서 올해 요구액 {base:,.0f}천원 · 산출근거를 모두 더하면 {leaves:,.0f}천원 "
               f"— {gap:,.0f}천원 차이. 설명서 안 어느 금액과도 맞지 않습니다."
               + (f" ({spoken})" if spoken else ""), base, leaves)


def _families(items) -> dict:
    """부모 경로 -> 그 아래 바로 붙은 자식들.

    두 서식을 마주 놓고 볼 때 이름만으로는 짝을 지을 수 없다. 한 사업 안에 '공립'
    '기본' '운영용품' 이 여러 번 나온다. 위로 거슬러 올라간 이름들을 열쇠로 써야
    같은 자리끼리 비교된다.
    """
    found: dict[tuple, list] = {}
    stack: list[tuple[int, str]] = []
    for item in items:
        while stack and stack[-1][0] >= item.depth:
            stack.pop()
        found.setdefault(tuple(name for _, name in stack), []).append(item)
        stack.append((item.depth, _squeeze(item.name)))
    return found


def _check_forms(project, report: Report) -> set:
    """같은 항목에 두 서식이 서로 다른 금액을 적었는지 본다.

    설명서는 한 사업의 산출 내역을 두 번 적는다. '요구내용 및 산출근거' 표와 사업계획
    절의 '산출내역' 표다. 담당자가 한쪽만 고치면 각 표 안에서는 합이 맞으니 검산으로는
    잡히지 않는다. 두 표를 마주 놓고 봐야 한다.

    쪼갠 방식이 다를 때는 말하지 않는다. 같은 금액을 '공립(1,402명)' 한 줄로 적은 표와
    '공립(202명) + 공립카드결제(1,200명)' 두 줄로 적은 표가 실제로 있고, 위 단계 금액은
    서로 같다. 자식 수가 다르면 나눠 적은 방식이 다른 것이니 그 가족은 건너뛴다.
    """
    reported: set = set()
    if project.from_outline or not project.had_own_items:
        return reported
    basis = _families(project.items)
    order = {id(item): index for index, item in enumerate(project.items)}
    for number, rows in enumerate(project.outline_items):
        for parent, kids in _families(rows).items():
            mates = basis.get(parent)
            if mates is None or len(mates) != len(kids):
                continue
            for kid in kids:
                if kid.amount is None:
                    continue
                name = _squeeze(kid.name)
                if sum(1 for one in kids if _squeeze(one.name) == name) != 1:
                    continue
                hits = [one for one in mates
                        if _squeeze(one.name) == name and one.amount is not None]
                if len(hits) != 1 or abs(hits[0].amount - kid.amount) <= TOLERANCE:
                    continue
                mate = hits[0]
                if (project.request is not None and project.base is not None
                        and {round(kid.amount, 3), round(mate.amount, 3)}
                        == {round(project.request, 3), round(project.base, 3)}):
                    # 한쪽 표가 이 줄에 총사업비를, 다른 쪽이 올해 요구액을 적은 경우다.
                    # 금액을 잘못 적은 게 아니라 칸의 뜻이 갈린 것이라 '오류'가 아니다.
                    # 어느 칸을 고쳐야 하는지는 _check_outline 이 '확인 필요'로 말한다.
                    continue
                reported.add(id(kid))
                reported.add(id(mate))
                report.add("목표1", "오류", project.name, mate.name,
                           f"'{mate.name}' 금액이 설명서 안에서 두 가지입니다 — "
                           f"요구내용 및 산출근거 표 {mate.amount:,.0f}천원 · "
                           f"사업계획 산출내역 표 {kid.amount:,.0f}천원 "
                           f"(차이 {abs(mate.amount - kid.amount):,.0f}천원). "
                           "한쪽만 고친 것 같습니다. 두 표를 같은 금액으로 맞춰 주세요.",
                           mate.amount, kid.amount, index=order.get(id(mate)),
                           form="산출내역 표", block=number, row=kid.row)
    return reported


def _check_outline(project, report: Report, skip=frozenset()) -> None:
    """사업계획 절의 '산출내역' 표도 그 표만으로 검산한다.

    예전에는 '요구내용 및 산출근거' 표가 없을 때만 이 표를 읽었다. 그래서 앞 표가 있는
    사업에서는 이 표의 산출식이 틀려도, 항목 금액이 아래 합과 어긋나도 아무 말이 없었다.
    두 표 모두 담당자가 K-에듀파인에 옮겨 적는 근거다. 둘 다 검산한다.
    """
    if project.from_outline:
        return        # 이 표를 이미 산출근거로 읽었다. _check_leaves 가 본 목록이다
    for number, rows in enumerate(project.outline_items):
        _check_leaves(project, report, items=rows, form="산출내역 표", skip=skip, block=number)
        _check_rollup(project, report, rows=rows, form="산출내역 표", skip=skip, block=number)


def _check_request(project, report: Report) -> None:
    """머리글 총사업비 = '2027년 본예산안(A)' 표의 올해 요구액.

    총사업비는 여러 해의 합계가 아니다. **올해 그 사업의 사업비**다. 표본 69개 사업 중
    67개가 두 값이 정확히 같고, 어긋난 2개는 성격이 서로 달랐다. 그래서 '다르다'로 끝내지
    않고 **산출근거 합계에게 물어 어느 칸이 튀는지** 가린다.

      26. 특수교육교재교구지원   총사업비 3,240,000 · 요구액(A) 1,040,000
          산출근거 합계도, 지원형태 계도, 집행계획 합계도, 투자실적 요구액도 1,040,000.
          머리글 한 칸만 튄다 → 오류. 고칠 곳이 한 곳으로 정해진다.

      59. 소규모유치원…(총액배분)  총사업비 1,023,500 · 요구액(A) 105,000
          산출근거 합계도 1,023,500이다. 그 쪽 표에 다른 사업 몫까지 함께 실려 있어서,
          어긋남의 뿌리가 머리글이 아니다. 같은 어긋남을 _check_total 이 표까지 짚어
          한 번 말하므로 여기서는 잠자코 있는다. 두 건으로 세지 않는다.
    """
    _one_request(project, report, project.name, project.request, project.year_request,
                 project.item_total if project.items else None)
    # 번호 없이 이어 붙은 세부 쪽도 저마다 머리글을 들고 온다. 산출근거는 한 트리로
    # 합쳐 두므로(_fold_unnumbered) 그 쪽 머리글은 아무도 보지 않는 칸이 되어 있었다.
    # 실제로 '특수학교(유치원) 방학 중 방과후 과정 급식비 지원' 쪽이 머리글 81,900,
    # 요구액 33,400 이었는데 조용히 지나갔다. 증인은 그 쪽 요구액과 같은 1단 항목이다.
    for part in project.parts:
        if part.request is None or part.year_request is None:
            continue
        witness = next((index for index, item in enumerate(project.items)
                        if item.depth == 1 and item.amount is not None
                        and abs(item.amount - part.year_request) <= TOLERANCE), None)
        _one_request(project, report, part.name, part.request, part.year_request,
                     None if witness is None else part.year_request,
                     item=part.name, index=witness)


def _one_request(project, report: Report, label: str, request, year_request,
                 leaves, item: str = "", index=None) -> None:
    """머리글 한 칸과 요구액 한 칸을 맞춰 본다. 사업에도, 세부 쪽에도 같은 잣대를 쓴다."""
    if request is None or year_request is None:
        return
    gap = abs(request - year_request)
    if gap <= TOLERANCE:
        return
    where = f"'{label}' 쪽의 " if item else ""
    pair = (f"{where}머리글 총사업비 {request:,.0f}천원 · "
            f"'본예산안(A)' 표의 올해 요구액 {year_request:,.0f}천원 "
            f"— {gap:,.0f}천원 차이. ")
    if leaves is not None and abs(leaves - request) <= TOLERANCE:
        return        # 산출근거도 총사업비 쪽이다. _check_total 이 표까지 짚어 말한다
    if leaves is None and project.handover:
        # 산출근거를 따로 셀 수 없는 재배정 사업 쪽이다. 입력하지 않는 사업이고, 어느 칸이
        # 틀렸는지 가릴 증인도 없다. 오류라고 단정하지 않는다.
        report.add("목표1", "확인 필요", project.name, item, pair + "(재배정 사업, 증인 없음)")
        return
    if leaves is not None and abs(leaves - year_request) <= TOLERANCE:
        report.add("목표1", "오류", project.name, item,
                   pair + f"산출근거를 모두 더하면 {leaves:,.0f}천원으로 요구액과 같습니다. "
                   "총사업비는 올해 그 사업의 사업비이므로 요구액과 같아야 합니다. "
                   "머리글 총사업비 한 칸만 다릅니다 — 그 칸을 고쳐 주세요. "
                   "(입력본에는 실리지 않는 칸이라 생성은 막지 않습니다)",
                   request, year_request, index=index, blocking=False)
        return
    report.add("목표1", "오류", project.name, item,
               pair + "총사업비는 올해 그 사업의 사업비이므로 두 값은 같아야 합니다."
               + (f" 산출근거를 모두 더하면 {leaves:,.0f}천원으로 양쪽 어디와도 다릅니다."
                  if leaves is not None else " 산출근거를 이 쪽만 따로 셀 수 없어 어느 쪽이 맞는지 "
                  "가리지 못했습니다.")
               + " 설명서에서 둘 중 틀린 칸을 고쳐 주세요.",
               request, year_request, index=index, blocking=False)


def check_plan(projects) -> Report:
    """설명서 안에서 닫히는 검사. 다른 파일이 없어도 돌아간다."""
    report = Report()
    for project in projects:
        _check_leaves(project, report)
        _check_shape(project, report)
        deeper = _check_rollup(project, report)
        _check_total(project, report, quiet=deeper)
        _check_request(project, report)
        # 두 서식이 어긋난 자리를 먼저 짚는다. 그 자리는 표 안의 합계도 같이 어긋나
        # 있으니, 같은 잘못을 두 번 세지 않도록 검산에서 건너뛴다.
        settled = _check_forms(project, report)
        _check_outline(project, report, skip=settled)
    _check_duplicates(projects, report)
    return report


def _check_duplicates(projects, report: Report) -> None:
    """같은 사업 쪽이 설명서에 두 번 실렸는지 본다. (2.13.0)

    2027 설명서에 '47. 유치원 입학관리시스템 연수' 쪽이 글자 하나 다르지 않게 두 번
    들어 있었다. 금액 검사는 쪽마다 따로 하므로 둘 다 '맞다'고 지나가고, 입력 대상
    사업이었다면 **같은 산출근거가 입력본에 두 번** 실린다. 합계가 두 배가 되는데
    어느 검사도 말하지 않는다. 이름·요구액·산출근거가 모두 같으면 중복 쪽으로 본다.
    """
    seen: dict = {}
    for project in projects:
        key = (_squeeze(project.name), project.base, round(project.item_total, 3),
               len(project.items))
        if key[1] is None:
            continue
        first = seen.get(key)
        if first is None:
            seen[key] = project
            continue
        doubled = bool(project.items) and not project.handover
        report.add("목표1", "오류" if doubled else "확인 필요", project.name, "",
                   f"설명서에 같은 사업 쪽이 두 번 있습니다 — 이름·올해 요구액 {project.base:,.0f}천원"
                   + (f"·산출근거 {len(project.items)}줄" if project.items else "")
                   + "이 모두 같습니다. "
                   + ("그대로 두면 입력본에 같은 줄이 두 번 실려 합계가 두 배가 됩니다. "
                      "설명서에서 겹친 쪽을 지워 주세요." if doubled
                      else "입력본에는 실리지 않는 사업이지만, 설명서에서 겹친 쪽인지 확인해 주세요."),
                   project.base, first.base)


# ---------------------------------------------------------------- 목표 3

def check_bimok(projects, resolver) -> tuple[Report, dict]:
    """설명서 목코드 5자리 + 작년 통계목 2자리로 비목을 확정한다."""
    report = Report()
    settled: dict[tuple[str, int], object] = {}
    for project in projects:
        if project.handover:
            continue          # 이 과에서 입력하지 않는 사업. 비목을 정할 일이 없다.
        for index, item in enumerate(project.items):
            if not item.is_leaf or item.handover:
                continue
            resolution = resolver.resolve(item.code5, item.name,
                                          parent=parent_name(project.items, index),
                                          bimok_name=item.bimok_name)
            settled[(project.name, index)] = resolution
            if not resolution.settled:
                report.add("목표3", "확인 필요", project.name, item.name,
                           f"{item.bimok_name}({item.bimok_code5}) 통계목 미확정 — {resolution.source}"
                           + (f" 후보: {', '.join(resolution.candidates)}" if resolution.candidates else ""),
                           index=index)
    return report, settled


# ---------------------------------------------------------------- 목표 2

def issue_row(items, issue) -> int:
    """지적이 가리키는 줄의 자리. 확실하지 않으면 -1.

    이름으로 찾으면 안 된다. 한 사업 안에 '공립' '기본' '운영용품' 같은 이름이
    여러 번 나온다. 이름으로 칠하면 멀쩡한 줄까지 함께 붉어지고, 담당자는 그걸
    보고 프로그램이 덧셈을 못 한다고 읽는다. 모르면 아무 줄도 칠하지 않는다.
    """
    index = getattr(issue, "index", None)
    if index is not None and 0 <= index < len(items):
        return index
    if not getattr(issue, "item", ""):
        return -1
    same = [position for position, item in enumerate(items) if item.name == issue.item]
    return same[0] if len(same) == 1 else -1


def parent_name(items, index: int) -> str:
    """설명서에서 이 항목의 바로 위 단계 항목명."""
    depth = items[index].depth
    for back in range(index - 1, -1, -1):
        if items[back].depth < depth:
            return items[back].name
    return ""


def _by_name(pairs: dict) -> dict:
    return {normalise(name): amount for name, amount in pairs.items()}


SPACING = re.compile(r"[\s·․ㆍ∙・_]")


def _squeeze(name: str) -> str:
    """띄어쓰기만 다른 이름을 같은 이름으로 본다.

    설명서는 '초등교육과정전문성신장', UBIS는 '초등교육과정 전문성 신장'으로 적는다.
    같은 사업인데 '찾지 못했습니다'라고 말하면 담당자는 없는 문제를 찾아다닌다.
    """
    return SPACING.sub("", normalise(name))


def _loose_index(pairs: dict) -> dict:
    """공백을 지운 이름 → [(원래 이름, 금액)]. 여럿이면 짐작하지 않는다."""
    index: dict = {}
    for name, amount in pairs.items():
        index.setdefault(_squeeze(name), []).append((name, amount))
    return index


def _lookup(exact: dict, loose: dict, name: str):
    """정확한 이름 → 공백만 다른 이름 순. 후보가 둘 이상이면 찾지 못한 것으로 둔다."""
    key = normalise(name)
    if key in exact:
        return exact[key], ""
    same = loose.get(_squeeze(name)) or []
    if len(same) == 1:
        return same[0][1], same[0][0]
    return None, ""


AGGREGATED = ("총액배분", "재원배분", "경상운영비")


def _rolled_up(project) -> bool:
    """UBIS 가 다른 사업과 한 줄로 묶어 적는 사업인가.

    총액배분·재원배분·경상운영비 사업과 번호 없이 이어 붙은 세부 쪽은 UBIS 에 제 이름의
    줄이 없는 게 정상이다(2025·2027 표본에서 이런 사업 18개가 전부 그랬다). 이름으로 못
    찾았다고 오류라 하면 목록이 그런 줄로 찬다.
    """
    return (bool(getattr(project, "handover", "")) or getattr(project, "unnumbered", False)
            or getattr(project, "split_description", False) or bool(getattr(project, "parts", []))
            or any(word in project.name for word in AGGREGATED))


def compare_totals(goal: str, projects, ubis: dict, form_path: Optional[str], form_label: str) -> Report:
    """사업별 금액을 한글 설명서 · 엑셀(UBIS 검토조서 / 입력본) 사이에서 맞춰 본다.

    설명서는 한 사업을 한 덩어리로 적기도 하고, 총괄 한 쪽 + 세부 여러 쪽으로 나눠
    적기도 한다. UBIS는 그걸 단위과제 하나로 묶어 둔다. 어느 쪽으로 읽어야 하는지
    설명서만 봐서는 단정할 수 없으므로 후보를 모두 내고, **어느 후보와도 맞지 않을
    때만** 오류라고 말한다.

    사업 이름이 다른 것은 오류로 세지 않는다. 설명서와 UBIS의 표기가 다를 뿐
    금액이 틀린 게 아니고, 이름 지적이 목록을 채우면 정작 볼 금액 오류가 묻힌다.
    """
    report = Report()
    ubis_by_name = _by_name(ubis)
    ubis_loose = _loose_index(ubis)
    named = {name for project in projects
             for name in [_lookup(ubis_by_name, ubis_loose, project.name)[1] or
                          (project.name if _lookup(ubis_by_name, ubis_loose, project.name)[0] is not None else "")]
             if name}

    # 번호가 앞 사업보다 작아진 쪽(59 뒤의 1·2·3·4)은 한 묶음 아래 다시 매긴 세부 쪽이다.
    restarted: set[int] = set()
    highest = 0
    for project in projects:
        number = int(project.number) if str(getattr(project, "number", "")).isdigit() else None
        if number is None:
            continue
        if number < highest:
            restarted.add(id(project))
        highest = max(highest, number)

    form_amounts: dict[str, float] = {}
    if form_path:
        for row in read_form(form_path):
            if row.level == 1 or row.indent == 0:
                form_amounts[row.name.strip()] = row.amount or 0
    form_by_name = _by_name(form_amounts)
    form_loose = _loose_index(form_amounts)

    for project in projects:
        candidates = [pair for pair in project.amount_candidates if pair[1] is not None]
        if not candidates:
            report.add(goal, SKIPPED, project.name, "", "설명서에서 올해 요구액을 찾지 못했습니다.")
            continue

        found, alias = _lookup(ubis_by_name, ubis_loose, project.name) if ubis else (None, "")
        if found is None and ubis:
            # 이름으로 못 찾으면 금액으로 짝을 찾는다. 설명서 금액과 같은 UBIS 사업이 **하나뿐**
            # 이고 그 사업을 설명서의 다른 사업이 쓰지 않으면, 이름만 다르게 적힌 같은 사업이다.
            same = [name for name, value in ubis.items()
                    if any(abs(value - amount) <= TOLERANCE for _label, amount in candidates)
                    and name not in named]
            if len(same) != 1 and not _rolled_up(project) and id(project) not in restarted:
                report.add(goal, "오류", project.name, "",
                           "UBIS 세출요구에 이 사업이 없습니다. 이름도 금액도 맞는 사업을 찾지 못했습니다 — "
                           "UBIS 에 요구를 넣지 않았거나 사업명이 다릅니다.")
        if found is not None:
            _compare(report, goal, project, candidates, found, "엑셀 UBIS 검토조서")

        if form_path and not getattr(project, "handover", ""):
            in_form, form_alias = _lookup(form_by_name, form_loose, project.name)
            if in_form is None:
                report.add(goal, "오류", project.name, "",
                           f"{form_label}에 이 사업이 없습니다. 입력 대상 사업인데 빠졌습니다.")
            else:
                _compare(report, goal, project, candidates, in_form, f"엑셀 {form_label}")

    if form_path:
        known = {normalise(project.name) for project in projects}
        loose_known = {_squeeze(project.name) for project in projects}
        for name in form_amounts:
            if normalise(name) not in known and _squeeze(name) not in loose_known:
                report.add(goal, "오류", name, "",
                           f"{form_label}에만 있는 사업입니다. 설명서에 없는 사업이 들어갔습니다.")
    return report


def _compare(report: Report, goal: str, project, candidates: list, other: float, other_label: str) -> None:
    """설명서 금액과 엑셀 금액을 맞춘다.

    차액이 있어도 늘 오류는 아니다. 설명서가 한 사업을 여러 쪽으로 나눠 적었거나
    번호 없는 쪽이 섞여 있으면, UBIS 가 그것들을 한 사업으로 묶어 두었을 수 있다.
    그건 금액이 틀린 게 아니라 사업을 나눈 기준이 다른 것이라 오류로 세지 않는다.
    """
    if any(abs(value - other) <= TOLERANCE for _label, value in candidates):
        return
    label, closest = min(candidates, key=lambda pair: abs(pair[1] - other))
    split = (getattr(project, "unnumbered", False) or getattr(project, "split_description", False)
             or bool(getattr(project, "parts", [])))
    detail = ""
    if len(candidates) > 1:
        detail = ("  (설명서 " + " · ".join(f"{name} {value:,.0f}" for name, value in candidates)
                  + " 천원)")
    if split:
        detail += ("  설명서는 이 사업 언저리를 여러 쪽으로 나눠 적었습니다. UBIS 가 그것들을 "
                   "한 사업으로 묶어 두었다면 금액이 틀린 게 아니라 사업을 나눈 기준이 다른 것입니다.")
    report.add(goal, "확인 필요" if split else "오류", project.name, "",
               f"한글 설명서 {closest:,.0f}천원({label}) · {other_label} {other:,.0f}천원 "
               f"— {abs(closest - other):,.0f}천원 차이.{detail}", closest, other)


def check_generated(projects, ubis: dict, generated_path: str) -> Report:
    """목표 2A — 입력 전. 생성한 입력본과 맞춘다."""
    return compare_totals("목표2A", projects, ubis, generated_path, "생성한 입력본")


def check_downloaded(projects, ubis: dict, downloaded_path: str) -> Report:
    """목표 2B — 입력 후. K-에듀파인에서 내려받은 파일과 맞춘다."""
    report = compare_totals("목표2B", projects, ubis, downloaded_path, "입력 후 다운로드본")
    for row in read_form(downloaded_path):
        if row.level is not None and row.level != row.indent // 2 + 1:
            report.add("목표2B", "오류", row.name, "",
                       f"{row.row}행 레벨이 들여쓰기와 맞지 않습니다.", row.level, row.indent // 2 + 1)
        if row.basis and row.amount is not None:
            computed = calc(row.basis)
            if computed is not None and abs(computed - row.amount) > 1:
                report.add("목표2B", "오류", row.name, "",
                           f"{row.row}행 산출기초 계산값과 금액이 다릅니다.", row.amount, computed)
        if row.basis and not row.code:
            report.add("목표2B", "오류", row.name, "", f"{row.row}행 원가통계비목이 비어 있습니다.")
    return report


def _same_amount(one, other) -> bool:
    """입력본은 천원 정수다. 1천원 오타도 잡아야 하므로 반올림 차이만 눈감는다."""
    if one is None or other is None:
        return one is None and other is None
    return abs(one - other) < 0.5


def diff_forms(generated_path: str, downloaded_path: str) -> Report:
    """생성한 입력본과 다운로드본을 행 단위로 맞춰 본다. 서식이 같으므로 1:1 비교."""
    report = Report()
    made = read_form(generated_path)
    got = read_form(downloaded_path)
    remaining = list(got)
    for row in made:
        if not row.basis:
            continue
        # 생성본에서 비목이 빈칸이던 줄은 담당자가 K-에듀파인에서 직접 골라 넣는다.
        # 그 줄의 비목까지 같아야 한다고 보면, 미확정 행이 전부 '그대로 들어가지 않았다'
        # (오류)와 '생성본에 없던 줄'(확인 필요)로 두 번씩 쏟아진다. 2027 설명서에서는
        # 600건이 넘는다. 비목이 채워졌는지는 check_downloaded 가 따로 본다.
        match = next((other for other in remaining
                      if normalise(other.name) == normalise(row.name)
                      and other.basis == row.basis
                      and _same_amount(other.amount, row.amount)
                      and (not row.code or other.code == row.code)), None)
        if match is None:
            report.add("목표2B", "오류", row.name, "",
                       f"생성본 {row.row}행이 다운로드본에 그대로 들어가지 않았습니다. "
                       f"(비목 {row.code or '없음'}, 산출기초 {row.basis}, 금액 {row.amount})")
        else:
            remaining.remove(match)
    for leftover in remaining:
        if leftover.basis:
            report.add("목표2B", "오류", leftover.name, "",
                       f"다운로드본 {leftover.row}행은 생성본에 없던 줄입니다. 설명서에 없는 산출근거가 "
                       "입력됐습니다.")
    return report
