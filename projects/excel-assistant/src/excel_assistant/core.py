from __future__ import annotations

from dataclasses import dataclass
from difflib import SequenceMatcher
import re
from typing import Any, Iterable


class RequestError(ValueError):
    """Raised when a request cannot be converted to a safe task plan."""


ALIASES = {
    "이름": ("성명", "성함", "담당자", "직원명", "교사명", "고객명", "유아명", "학부모명"),
    "금액": ("합계", "비용", "가격", "매출액", "예산액", "지출액", "지급액", "지급금액", "강의비", "강의료", "수당"),
    "날짜": ("일자", "등록일", "처리일", "지급일", "시작일", "종료일"),
    "부서": ("소속", "팀", "부서명", "소속부서", "소속기관"),
    "상태": ("진행상태", "처리상태", "결과", "완료여부"),
    "코드": ("번호", "id", "식별자", "일련번호", "연번", "순번"),
    "수량": ("개수", "건수", "인원", "인원수", "유아수"),
}


@dataclass
class TaskPlan:
    operation: str
    group_column: str | None = None
    group_columns: tuple[str, ...] = ()
    value_column: str | None = None
    aggregation: str = "sum"
    filter_column: str | None = None
    filter_value: Any = None
    filter_operator: str = "eq"
    sort_column: str | None = None
    component_columns: tuple[str, ...] = ()
    total_column: str | None = None
    descending: bool = False
    include_statistics: bool = False
    confidence: float = 0.0
    confidence_margin: float = 1.0
    ambiguous_columns: tuple[str, ...] = ()
    explanation: str = ""

    def selected_columns(self) -> tuple[str, ...]:
        """Return every input column required to execute this plan."""
        values = (
            *(self.group_columns or (() if self.group_column is None else (self.group_column,))),
            self.value_column, self.filter_column, self.sort_column,
            *self.component_columns, self.total_column,
        )
        return tuple(dict.fromkeys(value for value in values if value))

    @property
    def operation_label(self) -> str:
        if self.operation == "group_sum":
            return {"sum": "그룹별 합계표", "average": "그룹별 평균표", "count": "그룹별 개수표"}.get(self.aggregation, "그룹별 요약표")
        return {
            "filter": "조건 필터표", "sort": "정렬표", "duplicates": "중복 항목표", "sum_check": "합계 검증표",
        }.get(self.operation, self.operation)


def _clean(text: str) -> str:
    return re.sub(r"[^0-9a-zA-Z가-힣]+", "", str(text)).lower()


def _alias_family(value: str) -> set[str]:
    normalized = _clean(value)
    for canonical, aliases in ALIASES.items():
        family = {_clean(canonical), *(_clean(item) for item in aliases)}
        if normalized in family:
            return family
    return {normalized}


def _best_header(
    phrase: str,
    headers: Iterable[str],
    profiles: dict[str, dict[str, Any]] | None = None,
    expected_kind: str | None = None,
) -> tuple[str | None, float]:
    target = _clean(phrase)
    if not target:
        return None, 0.0
    target_family = _alias_family(target)
    best: tuple[str | None, float] = (None, 0.0)
    for header in headers:
        normalized = _clean(header)
        score = SequenceMatcher(None, target, normalized).ratio()
        if target == normalized:
            score = 1.0
        elif target in normalized or normalized in target:
            score = max(score, 0.92)
        if target_family & _alias_family(normalized):
            score = max(score, 0.88)
        if expected_kind and profiles:
            kind = profiles.get(str(header), {}).get("kind")
            score += 0.05 if kind == expected_kind else -0.12
        if score > best[1]:
            best = (str(header), min(score, 1.0))
    return best


def _annotate_ambiguity(
    plan: TaskPlan, request: str, headers: list[str], profiles: dict[str, dict[str, Any]] | None,
) -> TaskPlan:
    """Flag alias matches where another header is nearly as plausible."""
    request_clean = _clean(request)
    margins: list[float] = []
    ambiguous: list[str] = []
    for selected in plan.selected_columns():
        family = _alias_family(selected)
        mentioned = [word for word in family if word and word in request_clean]
        if not mentioned:
            continue
        target = max(mentioned, key=len)
        scored: list[tuple[str, float]] = []
        expected = profiles.get(selected, {}).get("kind") if profiles else None
        for header in headers:
            normalized = _clean(header)
            score = SequenceMatcher(None, target, normalized).ratio()
            if target == normalized:
                score = 1.0
            elif target in normalized or normalized in target:
                score = max(score, 0.92)
            if _alias_family(target) & _alias_family(normalized):
                score = max(score, 0.88)
            if expected and profiles:
                score += 0.05 if profiles.get(header, {}).get("kind") == expected else -0.12
            scored.append((header, min(score, 1.0)))
        ranked = sorted(scored, key=lambda item: item[1], reverse=True)
        selected_score = next((score for header, score in ranked if header == selected), 0.0)
        alternative = max((score for header, score in ranked if header != selected), default=0.0)
        margin = max(selected_score - alternative, 0.0)
        margins.append(margin)
        if margin < 0.08 and alternative >= 0.75:
            ambiguous.extend(header for header, score in ranked if header != selected and selected_score - score < 0.08)
    plan.confidence_margin = min(margins, default=1.0)
    plan.ambiguous_columns = tuple(dict.fromkeys(ambiguous))
    if plan.ambiguous_columns:
        plan.confidence = min(plan.confidence, 0.64)
    return plan


def _mentioned_header(
    request: str,
    headers: list[str],
    profiles: dict[str, dict[str, Any]] | None = None,
    expected_kind: str | None = None,
    exclude: set[str] | None = None,
) -> tuple[str | None, float]:
    exclude = exclude or set()
    request_clean = _clean(request)
    candidates: list[tuple[str, float]] = []
    for header in headers:
        if header in exclude:
            continue
        family = _alias_family(header)
        score = 0.0
        if _clean(header) and _clean(header) in request_clean:
            score = 1.0
        elif any(alias and alias in request_clean for alias in family):
            score = 0.88
        if expected_kind and profiles:
            kind = profiles.get(header, {}).get("kind")
            score += 0.05 if kind == expected_kind else -0.12
        if score > 0:
            candidates.append((header, min(score, 1.0)))
    return max(candidates, key=lambda item: item[1], default=(None, 0.0))


def _numeric_header(
    headers: list[str], profiles: dict[str, dict[str, Any]] | None, request: str, exclude: set[str] | None = None,
) -> tuple[str | None, float]:
    exclude = exclude or set()
    numeric = [header for header in headers if header not in exclude and profiles and profiles.get(header, {}).get("kind") == "number"]
    mentioned = _mentioned_header(request, numeric, profiles, "number")
    if mentioned[0]:
        return mentioned
    for word in ("금액", "비용", "가격", "수당", "강의비", "점수", "수량", "인원수"):
        if word in request:
            matched, score = _best_header(word, numeric or headers, profiles, "number")
            if matched:
                return matched, score
    if len(numeric) == 1:
        return numeric[0], 0.9
    return None, 0.0


def _group_from_request(request: str, headers: list[str], profiles=None) -> tuple[str | None, float]:
    patterns = (r"(.+?)(?:을|를)?\s*기준으로", r"(.+?)별(?:로)?")
    for pattern in patterns:
        match = re.search(pattern, request)
        if match:
            return _best_header(match.group(1), headers, profiles, "text")
    return _mentioned_header(request, headers, profiles, "text")


def _groups_from_request(request: str, headers: list[str], profiles=None) -> tuple[tuple[str, ...], float]:
    phrase = None
    for pattern in (r"(.+?)(?:을|를)?\s*기준으로", r"(.+?)별(?:로)?"):
        match = re.search(pattern, request)
        if match:
            phrase = match.group(1)
            break
    if not phrase:
        group, score = _group_from_request(request, headers, profiles)
        return ((group,) if group else ()), score
    parts = [part.strip() for part in re.split(r"\s*(?:와|과|및|,|·)\s*", phrase) if part.strip()]
    resolved = []
    scores = []
    for part in parts:
        column, score = _best_header(part, headers, profiles, "text")
        if column and column not in resolved:
            resolved.append(column)
            scores.append(score)
    return tuple(resolved), min(scores, default=0.0)


def _sum_check_plan(request: str, headers: list[str], profiles=None) -> TaskPlan | None:
    if "합계" not in request or not any(word in request for word in ("맞", "일치", "검증", "확인")):
        return None
    numeric = [header for header in headers if profiles and profiles.get(header, {}).get("kind") == "number"]
    total_candidates = [header for header in numeric if any(word in _clean(header) for word in ("합계", "총계", "총합"))]
    total = total_candidates[0] if len(total_candidates) == 1 else None
    explicitly_named = [header for header in numeric if _clean(header) in _clean(request) and header != total]
    components = explicitly_named or [header for header in numeric if header != total]
    if total and len(components) >= 2:
        return TaskPlan(
            operation="sum_check", component_columns=tuple(components), total_column=total,
            confidence=0.92 if explicitly_named else 0.78,
            explanation=f"{', '.join(components)}의 행별 합계와 '{total}' 열이 일치하는지 검증합니다.",
        )
    return None


def _parse_filter(request: str, headers: list[str], profiles=None) -> TaskPlan | None:
    operators = (
        (r"(.+?)(?:이|가)?\s*([^ ]+?)\s*이상", "gte"),
        (r"(.+?)(?:이|가)?\s*([^ ]+?)\s*초과", "gt"),
        (r"(.+?)(?:이|가)?\s*([^ ]+?)\s*이하", "lte"),
        (r"(.+?)(?:이|가)?\s*([^ ]+?)\s*미만", "lt"),
        (r"(.+?)(?:이|가)\s*([^ ]+?)(?:이|가)?\s*아닌", "ne"),
        (r"(.+?)(?:에|열에)\s*([^ ]+?)(?:이|가)?\s*(?:포함|들어)", "contains"),
        (r"(.+?)(?:이|가)\s*([^ ]+?)(?:인|인 것|만)(?:\s|$)", "eq"),
    )
    for pattern, operator in operators:
        match = re.search(pattern, request)
        if not match:
            continue
        column, score = _best_header(match.group(1), headers, profiles)
        raw_value = match.group(2).strip(" ,원%")
        numeric = coerce_number(raw_value)
        value = numeric if operator in {"gte", "gt", "lte", "lt"} and numeric is not None else raw_value
        labels = {"eq": "같은", "ne": "다른", "contains": "포함하는", "gte": "이상인", "gt": "초과인", "lte": "이하인", "lt": "미만인"}
        return TaskPlan(
            operation="filter", filter_column=column, filter_value=value, filter_operator=operator,
            confidence=score, explanation=f"'{column}' 값이 '{value}' {labels[operator]} 행만 추출합니다.",
        )
    return None


def _analyze_request_impl(request: str, headers: list[str], profiles: dict[str, dict[str, Any]] | None = None) -> TaskPlan:
    text = re.sub(r"\s+", " ", request.strip())
    if not text:
        raise RequestError("원하는 작업을 입력해 주세요.")
    if not headers:
        raise RequestError("선택한 시트에서 열 제목을 찾지 못했습니다.")

    if "중복" in text:
        column, score = _mentioned_header(text.replace("중복", ""), headers, profiles)
        if column is None:
            before = text.split("중복", 1)[0]
            column, score = _best_header(before, headers, profiles)
        return TaskPlan(
            operation="duplicates", group_column=column, confidence=score,
            explanation=f"'{column}' 열의 중복값과 해당 원본 행을 별도 시트로 추출합니다.",
        )

    sum_check = _sum_check_plan(text, headers, profiles)
    if sum_check:
        return sum_check

    wants_statistics = any(word in text for word in ("최대값", "최댓값", "최소값", "최솟값"))
    aggregate_words = any(word in text for word in ("합산", "합계", "더해", "평균", "개수", "건수", "인원수", "몇 명", "몇명"))
    if wants_statistics or aggregate_words:
        groups, group_score = _groups_from_request(text, headers, profiles)
        group = groups[0] if groups else None
        aggregation = "average" if "평균" in text and not any(word in text for word in ("최대", "최소")) else "count" if any(word in text for word in ("개수", "건수", "인원수", "몇 명", "몇명")) else "sum"
        value, value_score = _numeric_header(headers, profiles, text, set(groups))
        if aggregation == "count":
            value, value_score = group, group_score
        if group is None:
            raise RequestError("그룹으로 묶을 기준 열을 정하지 못했습니다. 예: '부서별로' 또는 '이름을 기준으로'라고 입력해 주세요.")
        if value is None:
            raise RequestError("계산할 숫자 열을 정하지 못했습니다. 계산할 열 이름을 요청에 포함해 주세요.")
        descending = any(word in text for word in ("큰 순", "내림차순", "높은 순", "많은 순"))
        label = {"sum": "합산", "average": "평균 계산", "count": "개수 계산"}[aggregation]
        return TaskPlan(
            operation="group_sum", group_column=group, group_columns=groups, value_column=value, aggregation=aggregation,
            descending=descending, include_statistics=wants_statistics, confidence=min(group_score, value_score),
            explanation=f"'{group}'별로 '{value}'을 {label}하고 {'큰 순서로 정렬합니다.' if descending else '요약합니다.'}",
        )

    filter_plan = _parse_filter(text, headers, profiles)
    if filter_plan:
        return filter_plan

    if any(word in text for word in ("정렬", "큰 순", "작은 순", "오름차순", "내림차순", "가나다순")):
        column, score = _mentioned_header(text, headers, profiles)
        if column is None:
            phrase = re.split(r"(큰 순|작은 순|높은 순|낮은 순|오름차순|내림차순|가나다순)", text, maxsplit=1)[0]
            column, score = _best_header(re.sub(r"(을|를|기준으로)$", "", phrase).strip(), headers, profiles)
        descending = any(word in text for word in ("큰 순", "높은 순", "내림차순", "많은 순"))
        return TaskPlan(
            operation="sort", sort_column=column, descending=descending, confidence=score,
            explanation=f"'{column}' 열을 {'내림차순' if descending else '오름차순'}으로 정렬합니다.",
        )

    raise RequestError("지원하는 요청을 찾지 못했습니다. 합계·평균·개수, 조건 필터, 정렬 또는 중복 찾기를 구체적으로 입력해 주세요.")


def analyze_request(request: str, headers: list[str], profiles: dict[str, dict[str, Any]] | None = None) -> TaskPlan:
    plan = _analyze_request_impl(request, headers, profiles)
    return _annotate_ambiguity(plan, request, headers, profiles)


def validate_task_plan(plan: TaskPlan, headers: Iterable[str]) -> None:
    """Validate the parser/UI contract before any workbook is written."""
    available = set(headers)
    supported = {"group_sum", "filter", "sort", "duplicates", "sum_check"}
    if plan.operation not in supported:
        raise RequestError(f"지원하지 않는 작업입니다: {plan.operation}")
    missing = [column for column in plan.selected_columns() if column not in available]
    if missing:
        raise RequestError(f"원본에서 다음 열을 찾을 수 없습니다: {', '.join(missing)}")
    if plan.operation == "group_sum":
        groups = plan.group_columns or (() if plan.group_column is None else (plan.group_column,))
        if not groups or not plan.value_column:
            raise RequestError("그룹 기준 열과 계산 열을 모두 선택해 주세요.")
        if plan.aggregation not in {"sum", "average", "count"}:
            raise RequestError("지원하지 않는 계산 방식입니다.")
    elif plan.operation == "filter":
        if not plan.filter_column or plan.filter_value in (None, ""):
            raise RequestError("필터 기준 열과 조건값을 모두 입력해 주세요.")
        if plan.filter_operator not in {"eq", "ne", "contains", "gt", "gte", "lt", "lte"}:
            raise RequestError("지원하지 않는 필터 조건입니다.")
        if plan.filter_operator in {"gt", "gte", "lt", "lte"} and coerce_number(plan.filter_value) is None:
            raise RequestError("숫자 비교 조건에는 숫자 값을 입력해 주세요.")
    elif plan.operation == "sort" and not plan.sort_column:
        raise RequestError("정렬 기준 열을 선택해 주세요.")
    elif plan.operation == "duplicates" and not plan.group_column:
        raise RequestError("중복 검사 열을 선택해 주세요.")
    elif plan.operation == "sum_check":
        if len(plan.component_columns) < 2 or not plan.total_column:
            raise RequestError("더할 열을 2개 이상 선택하고 비교할 합계 열을 선택해 주세요.")
        if plan.total_column in plan.component_columns:
            raise RequestError("합계 열은 더할 열과 다르게 선택해 주세요.")


def coerce_number(value: Any) -> float | None:
    if value is None or value == "" or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    negative = text.startswith("(") and text.endswith(")")
    cleaned = re.sub(r"[^0-9.\-]", "", text)
    if cleaned in ("", "-", "."):
        return None
    try:
        number = float(cleaned)
        return -number if negative else number
    except ValueError:
        return None
