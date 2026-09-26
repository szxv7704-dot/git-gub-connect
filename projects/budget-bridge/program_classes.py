"""'세출예산 사업별 분류'(별표 3) HWPX를 읽어 사업 분류 체계를 만든다.

이 표는 **사업 분류**다. 원가통계비목이 아니다. 두 축은 서로 다르다.

    사업 분류   이 돈이 어느 사업인가   [02]교수학습활동지원 > [01]교육과정운영 > [01]교육과정운영지원
    원가통계비목  이 돈을 무엇에 쓰는가   운영용품 → 일반수용비(210-01) → 2100143

한 사업 안에서 목코드가 1종부터 12종까지 쓰인다(표본 69개 사업 중 50개가 2종 이상).
그러므로 사업 제목으로 비목을 정할 수 없다. 이 파일은 비목과 무관하다.

할 수 있는 일은 이것이다.
  - 설명서의 정책·단위·세부사업 이름에 공식 코드를 붙인다.
  - 분류표에 없는 이름, 부모가 어긋난 분류를 짚는다.
  - '설정 예시'를 보여 사업 배치가 맞는지 담당자가 판단하게 한다.
"""

from __future__ import annotations

import html
import re
from dataclasses import dataclass, field
from typing import Optional

from hwpx_native_parser import native_chunks

MARK = re.compile(r"^\[(\d+)\]\s*(.+)$")
LEVELS = ["분야", "부문", "정책사업", "단위사업", "세부사업"]
POLICY, UNIT, PROGRAM = 2, 3, 4


def normalise(name: str) -> str:
    """띄어쓰기·가운뎃점·괄호를 지우고 비교한다. 설명서와 분류표의 표기가 조금씩 다르다."""
    return re.sub(r"[\s·・\-_()·・]", "", name or "")


@dataclass
class ProgramClass:
    level: int
    code: str
    name: str
    parents: dict = field(default_factory=dict)     # {"정책사업": (코드, 이름), …}
    note: str = ""                                  # 설정 예시

    @property
    def label(self) -> str:
        return LEVELS[self.level] if self.level < len(LEVELS) else "기타"

    @property
    def full_code(self) -> str:
        parts = [code for code, _name in self.parents.values()] + [self.code]
        return "-".join(parts)


class Classification:
    """분류표 한 부."""

    def __init__(self) -> None:
        self.entries: list[ProgramClass] = []
        self.by_level: dict[int, dict[str, ProgramClass]] = {}

    def add(self, entry: ProgramClass) -> None:
        self.entries.append(entry)
        self.by_level.setdefault(entry.level, {}).setdefault(normalise(entry.name), entry)

    def find(self, level: int, name: str) -> Optional[ProgramClass]:
        return self.by_level.get(level, {}).get(normalise(name))

    def __len__(self) -> int:
        return len(self.entries)

    @property
    def counts(self) -> dict:
        return {LEVELS[level] if level < len(LEVELS) else str(level): len(items)
                for level, items in sorted(self.by_level.items())}


def _rows(table_html: str) -> list[list[str]]:
    return [[html.unescape(re.sub(r"<[^>]+>", "", cell)).strip()
             for cell in re.findall(r"<td[^>]*>(.*?)</td>", row, re.S)]
            for row in re.findall(r"<tr>(.*?)</tr>", table_html, re.S)]


def load_classification(path: str) -> Classification:
    """분류표 HWPX를 읽는다. 표에서 `[NN]이름` 이 있는 열 번호가 곧 단계다."""
    classes = Classification()
    path_now: dict[int, tuple[str, str]] = {}
    for chunk in native_chunks(path):
        if chunk.get("type") != "table":
            continue
        rows = _rows(chunk["text"])
        if len(rows) < 5:
            continue
        for cells in rows:
            found = [(index, MARK.match(cell)) for index, cell in enumerate(cells) if MARK.match(cell)]
            if not found:
                continue
            index, hit = found[0]
            level = min(index, len(LEVELS) - 1)
            path_now[level] = (hit.group(1), hit.group(2).strip())
            for deeper in [key for key in path_now if key > level]:
                path_now.pop(deeper)
            tail = cells[-1] if len(cells) > index + 1 else ""
            classes.add(ProgramClass(
                level=level, code=hit.group(1), name=hit.group(2).strip(),
                parents={LEVELS[key]: value for key, value in path_now.items() if key < level},
                note=tail if not MARK.match(tail) else ""))
    if not classes.entries:
        raise ValueError("사업별 분류표를 읽지 못했습니다. '세출예산 사업별 분류' 문서가 맞는지 확인해 주세요.")
    return classes


def check_projects(projects, classes: Classification):
    """설명서의 정책·단위·세부사업을 분류표와 맞춰 본다."""
    from crosscheck import Report

    report = Report()
    for project in projects:
        found = {}
        for level, value in ((POLICY, project.policy), (UNIT, project.unit), (PROGRAM, project.program)):
            if not value:
                report.add("목표4", "확인 필요", project.name, LEVELS[level], "설명서에 값이 없습니다.")
                continue
            entry = classes.find(level, value)
            found[level] = entry
            if entry is None:
                report.add("목표4", "오류", project.name, LEVELS[level],
                           f"분류표에 없는 {LEVELS[level]} 이름입니다: {value}")

        program = found.get(PROGRAM)
        unit = found.get(UNIT)
        if program is not None and unit is not None:
            expected = program.parents.get("단위사업")
            if expected and normalise(expected[1]) != normalise(project.unit):
                report.add("목표4", "오류", project.name, "단위사업",
                           f"'{project.program}' 은 분류표에서 '{expected[1]}' 아래입니다. "
                           f"설명서에는 '{project.unit}' 으로 적혀 있습니다.")
    return report


def describe(project, classes: Classification) -> str:
    """상세 화면에 보여 줄 한 줄. 코드와 '설정 예시'."""
    parts = []
    for level, value in ((POLICY, project.policy), (UNIT, project.unit), (PROGRAM, project.program)):
        entry = classes.find(level, value)
        parts.append(f"[{entry.code}]{value}" if entry else f"[?]{value}")
    line = " > ".join(parts)
    program = classes.find(PROGRAM, project.program)
    if program and program.note:
        line += "\n설정 예시 — " + program.note.replace("◦", " · ").strip(" ·")
    return line
