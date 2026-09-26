"""원가통계비목(7자리)을 결정한다.

설명서에는 목코드가 `210-01` 처럼 5자리까지만 적혀 있고, K-에듀파인은 7자리를 쓴다.
2026년 확정 다운로드 파일과 대조해 **앞 5자리가 그대로 일치**함을 확인했다(대조 가능한 7건 전부).

    일반수용비(210-01)  →  2100143
    강사수당(210-06)    →  2100605
    시설임차료(210-07)  →  2100702
    사업추진경비(230-02) →  2300201
    연구용역비(260-01)  →  2600101

따라서 추정해야 하는 것은 **뒤 2자리(통계목)뿐**이다. 그 2자리는 지난 연도의
확정 다운로드 파일에서 배운다. 배우지 못하면 추측하지 않고 미확정으로 남긴다.
"""

from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import Optional

from openpyxl import load_workbook

from edufine_form import FormRow, read_form

BRACKET = re.compile(r"^\[(\d{7})\]\s*(.*)$")
MARKER = re.compile(r"^\s*(?:\(?[0-9]+\)?[.)]|\(?[가-힣]\)?[.)])\s*")


def normalise(name: str) -> str:
    """'(1) 운영용품' '가) 운영용품' → '운영용품'"""
    cleaned = name.strip()
    previous = None
    while cleaned != previous:
        previous = cleaned
        cleaned = MARKER.sub("", cleaned).strip()
    return cleaned


def _is_request_form(path: str) -> bool:
    """세출예산요구내역 서식인지. 첫 행에 '레벨'이 있으면 그쪽이다."""
    book = load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = book.active
        head = "".join(str(sheet.cell(1, column).value or "") for column in range(1, 8))
    finally:
        book.close()          # read_only 로 연 통합문서는 닫아야 윈도우에서 파일이 풀린다
    return "레벨" in head


def _read_budget_sheet(path: str):
    """예산현액 조회 서식. 사업항목의 들여쓰기가 계층이고, 비목은 '[2100143] 일반수용비'."""
    sheet = load_workbook(path, data_only=True).active
    columns = {}
    for column in range(1, sheet.max_column + 1):
        label = str(sheet.cell(1, column).value or "").replace(" ", "")
        for key in ("사업항목", "원가통계비목", "산출기초"):
            if key in label and key not in columns:
                columns[key] = column
    if "사업항목" not in columns or "원가통계비목" not in columns:
        raise ValueError("K-에듀파인 세출예산요구내역도 예산현액도 아닙니다.")

    rows = []
    for number in range(2, sheet.max_row + 1):
        name = sheet.cell(number, columns["사업항목"]).value
        if not isinstance(name, str) or not name.strip():
            continue
        found = BRACKET.match(str(sheet.cell(number, columns["원가통계비목"]).value or "").strip())
        rows.append(FormRow(
            row=number,
            level=None,
            sequence=None,
            name=name.strip(),
            indent=len(name) - len(name.lstrip()),
            code=found.group(1) if found else "",
            bimok_name=found.group(2).strip() if found else "",
            detail="",
            basis=str(sheet.cell(number, columns.get("산출기초", 0) or 1).value or "").strip()
            if "산출기초" in columns else "",
        ))
    return rows


@dataclass
class Resolution:
    code: Optional[str]
    detail: Optional[str]
    source: str                       # 확정 근거
    candidates: list[str] = field(default_factory=list)

    @property
    def settled(self) -> bool:
        return bool(self.code)


class BimokResolver:
    """엑셀 확정본에서 통계목 2자리와 상세내역을 배운다.

    비목 7자리는 두 파일이 반씩 만든다.

        한글 설명서            엑셀 확정본
        일반수용비(210-01)  +  ...43        →  2100143
           앞 5자리             뒤 2자리

    앞 5자리는 한글에 적혀 있다. 뒤 2자리는 한글에 없으므로 엑셀에서 배운다.
    그래서 '기록이 없다'는 말은 늘 **엑셀** 이야기다. 담당자가 화면만 보고
    어느 파일을 더 넣어야 하는지 알 수 있게, 문구에 '한글' '엑셀'을 적는다.
    """

    def __init__(self) -> None:
        # 부모 항목까지 보는 열쇠. 목코드와 항목명만으로는 갈리는 것이 부모로 갈린다.
        #   강사수당 > 기본  → 2100605      심사수당 > 기본  → 2100607
        # 작년 5부 110개 조합에서 부모까지 보면 전부 단일 결론이었다.
        self.by_parent: dict[tuple[str, str, str], Counter] = defaultdict(Counter)
        self.by_code_and_name: dict[tuple[str, str], Counter] = defaultdict(Counter)
        self.by_code: dict[str, Counter] = defaultdict(Counter)
        # (목코드 5자리, 비목명) → 7자리.
        #   210-06 강사수당 → 2100605 · 기타운영수당 → 2100607 · 위원회수당 → 2100612
        #   210-01 일반수용비 → 2100143 · 교육경비 → 2100122
        # 한글 설명서가 이미 비목명을 들고 있다. 앞 5자리만 쓰고 이름을 버리면
        # 같은 목코드 안에서 갈리는 것을 '후보 셋 중 못 고르겠다'고 미루게 된다.
        self.by_bimok_name: dict[tuple[str, str], Counter] = defaultdict(Counter)
        # 목코드 없이 이름만으로 찾는 열쇠. 산출내역 표에는 비목이 아예 없다.
        self.by_parent_only: dict[tuple[str, str], Counter] = defaultdict(Counter)
        self.by_name_only: dict[str, Counter] = defaultdict(Counter)
        self.detail_by_code: dict[str, Counter] = defaultdict(Counter)
        self.learned_rows = 0
        self.sources: list[str] = []

    def learn(self, path: str) -> int:
        """확정된 K-에듀파인 파일 한 부를 학습한다. 여러 번 불러도 된다.

        두 서식을 받는다. `세출예산요구내역`(레벨·순번이 있는 요구 서식)과
        `예산현액`(구분·사업항목·원가통계비목만 있는 조회 서식)이다. 둘 다 항목명과
        확정된 7자리 비목을 들고 있으므로 학습 재료로는 값이 같다. 서식이 다르다고
        버리면 담당자가 가진 자료의 절반을 못 쓴다.
        """
        added = 0
        rows = read_form(path) if _is_request_form(path) else _read_budget_sheet(path)
        for index, row in enumerate(rows):
            if not row.code or not row.name:
                continue
            code5 = row.code[:5]
            name = normalise(row.name)
            parent = ""
            for back in range(index - 1, -1, -1):
                if rows[back].indent < row.indent:
                    parent = normalise(rows[back].name)
                    break
            self.by_parent[(code5, parent, name)][row.code] += 1
            self.by_code_and_name[(code5, name)][row.code] += 1
            self.by_code[code5][row.code] += 1
            if row.bimok_name:
                self.by_bimok_name[(code5, normalise(row.bimok_name))][row.code] += 1
            self.by_parent_only[(parent, name)][row.code] += 1
            self.by_name_only[name][row.code] += 1
            if row.detail:
                self.detail_by_code[row.code][row.detail] += 1
            added += 1
        self.learned_rows += added
        self.sources.append(path)
        return added

    def resolve(self, code5: str, name: str, parent: str = "", bimok_name: str = "") -> Resolution:
        """목코드 5자리와 항목명으로 7자리를 정한다.

        좁혀 가는 순서는 좁은 것부터다. 부모 항목 → 비목명 → 항목명 → 목코드.
        하나로 좁혀지지 않으면 후보만 돌려주고 확정하지 않는다. 다수결로 덮으면
        그것이 오류인지 정당한 예외인지 판별할 기회가 사라진다.
        """
        digits = code5.replace("-", "")
        if not digits:
            # 산출내역 표에서 읽은 산출근거에는 목코드가 없다. 이름만으로 찾는다.
            # 부모까지 같은 기록을 먼저 보고, 없으면 항목명만 본다.
            context = normalise(parent)
            if context:
                counts = self.by_parent_only.get((context, normalise(name)))
                if counts and len(counts) == 1:
                    code = next(iter(counts))
                    return Resolution(code, self._detail(code),
                                      f"엑셀 확정본에서 '{context} > {normalise(name)}' 로 "
                                      f"{counts[code]}회 (한글에 목코드가 없어 이름으로 찾음)")
            only = self.by_name_only.get(normalise(name))
            if only and len(only) == 1:
                code = next(iter(only))
                return Resolution(code, self._detail(code),
                                  f"엑셀 확정본에서 항목명 '{normalise(name)}' 은 모두 {code}")
            if only:
                return Resolution(None, None,
                                  f"한글 설명서에 목코드가 없고, 엑셀 확정본에서 항목명 '{normalise(name)}' 의 "
                                  "코드가 둘 이상입니다. 직접 골라 주세요.", sorted(only))
            return Resolution(None, None,
                              f"한글 설명서에 목코드가 없고, 엑셀 확정본에서도 항목명 '{normalise(name)}' 을 "
                              "찾지 못했습니다. 직접 골라 주세요.")

        # 0. 설명서에 적힌 비목명이 가장 좁은 열쇠다. 목-세목이 같아도 비목명이
        #    다르면 7자리가 다르다. 표본의 (목코드, 비목명) 조합은 전부 단일 코드였다.
        if bimok_name:
            named = self.by_bimok_name.get((digits, normalise(bimok_name)))
            if named and len(named) == 1:
                code = next(iter(named))
                return Resolution(code, self._detail(code),
                                  f"엑셀 확정본에서 비목명 '{normalise(bimok_name)}' 은 모두 {code}")

        # 1. 부모 항목까지 같은 기록. '강사수당 > 기본' 과 '심사수당 > 기본' 을 가른다.
        # 2. 설명서의 비목명을 부모 자리에 놓고 한 번 더. 작년 부모가 비목명과 같은 경우가 많다.
        for context, label in ((normalise(parent), "부모 항목"), (normalise(bimok_name), "비목명")):
            if not context:
                continue
            counts = self.by_parent.get((digits, context, normalise(name)))
            if counts and len(counts) == 1:
                code = next(iter(counts))
                return Resolution(code, self._detail(code),
                                  f"엑셀 확정본에서 '{context} > {normalise(name)}' 로 {counts[code]}회 ({label})")

        key = (digits, normalise(name))
        exact = self.by_code_and_name.get(key)
        if exact and len(exact) == 1:
            code = next(iter(exact))
            return Resolution(code, self._detail(code), f"엑셀 확정본에서 같은 목코드·항목명으로 {exact[code]}회")
        if exact and len(exact) > 1:
            return Resolution(None, None,
                              f"엑셀 확정본에서 '{normalise(name)}' 의 코드가 둘 이상입니다. 상위 항목이 달랐습니다. 직접 골라 주세요.",
                              sorted(exact))

        same_code = self.by_code.get(digits)
        if same_code and len(same_code) == 1:
            code = next(iter(same_code))
            return Resolution(code, self._detail(code), f"엑셀 확정본에서 목코드 {code5}는 모두 {code}")
        if same_code and len(same_code) > 1:
            return Resolution(None, None, f"한글의 목코드 {code5} 는 읽었는데, 엑셀 확정본에서 뒤 두 자리가 여럿입니다. 직접 골라 주세요.", sorted(same_code))

        return Resolution(None, None, f"한글에 적힌 목코드 {code5} 가 엑셀 확정본에 한 줄도 없습니다. 뒤 두 자리를 배울 곳이 없으니, 이 목코드가 든 확정본을 더 넣거나 직접 골라 주세요.")

    def _detail(self, code: str) -> Optional[str]:
        counts = self.detail_by_code.get(code)
        if counts and len(counts) == 1:
            return next(iter(counts))
        return None

    def known_codes(self) -> list[tuple[str, str]]:
        """엑셀 확정본에 한 번이라도 나온 7자리와 그 비목명. 담당자가 고를 목록.

        목코드가 없는 줄은 7자리 전부를 사람이 정해야 한다. 빈칸에 숫자를 외워 치게
        하면 오타가 그대로 입력본에 실린다. 작년에 실제로 쓴 코드 중에서 고르게 한다.
        """
        names: dict[str, str] = {}
        for (_digits, bimok_name), counts in self.by_bimok_name.items():
            for code in counts:
                names.setdefault(code, bimok_name)
        for counts in self.by_code.values():
            for code in counts:
                names.setdefault(code, "")
        # 확정본에 비목명 칸이 없으면 그 코드를 받았던 작년 항목명을 대신 보인다.
        # '2100143' 만으로는 무엇인지 모르지만 '작년: 운영용품·자료제작' 이면 고를 수 있다.
        used: dict[str, Counter] = defaultdict(Counter)
        for name, counts in self.by_name_only.items():
            for code, count in counts.items():
                used[code][name] += count
        for code, label in list(names.items()):
            if not label and used.get(code):
                top = [name for name, _count in used[code].most_common(3)]
                names[code] = "작년: " + "·".join(top)
        return sorted(names.items())

    @property
    def coverage(self) -> dict:
        return {
            "학습한 행": self.learned_rows,
            "목코드 수": len(self.by_code),
            "항목명 조합 수": len(self.by_code_and_name),
            "부모+항목명 조합 수": len(self.by_parent),
            "파일": list(self.sources),
        }
