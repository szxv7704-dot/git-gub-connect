"""비목 미확정 행을 담당자가 한 번에 처리할 수 있는 단위로 묶는다.

566행을 하나씩 고치게 하면 고문이다. 목코드로 묶으면 35줄이 된다.

다만 목코드 하나에 통계목이 여럿인 경우가 있다. 2026년 확정본에서 확인:

    210-01 → 2100143 (운영용품·자료제작·홍보영상제작)
    210-01 → 2100122 (자료수집)

그래서 목코드로만 묶어 한 값을 넣으면 자료수집 같은 항목이 틀린 코드를 조용히
받는다. 갈리는 목코드는 항목명별로 펼쳐서 따로 정하게 한다.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from bimok_resolver import normalise


@dataclass
class NameGroup:
    """같은 목코드 안에서 같은 결론을 받는 항목명 묶음."""

    code5: str                                   # 210-01
    names: list[str] = field(default_factory=list)
    rows: int = 0
    code: str = ""                               # 확정된 7자리, 미확정이면 빈 문자열
    source: str = ""
    candidates: list[str] = field(default_factory=list)
    open_rows: int = 0                           # 이 묶음에서 실제로 비목이 빈 행 수

    @property
    def settled(self) -> bool:
        return bool(self.code)

    @property
    def label(self) -> str:
        """화면에 보일 이름. '부모 > 항목명' 열쇠는 항목명으로 모아 상위 항목 수를 붙인다.

        목코드 없는 '공립' 74행은 상위 항목이 46곳이다. 열쇠를 그대로 늘어놓으면
        '(공약) 통합교육지원 프로그램 운영 > 공립 · …' 처럼 정작 항목명이 묻힌다.
        """
        leaves: dict[str, list[str]] = {}
        for key in self.names:
            parent, _, leaf = key.rpartition(" > ")
            leaves.setdefault(leaf, []).append(parent)
        if any(parents != [""] for parents in leaves.values()):
            parts = []
            for leaf, parents in list(leaves.items())[:3]:
                places = [one for one in parents if one]
                parts.append(f"{leaf} (상위 {len(places)}곳)" if len(places) > 1
                             else f"{places[0]} > {leaf}" if places else leaf)
            return " · ".join(parts) + (f"  외 {len(leaves) - 3}종" if len(leaves) > 3 else "")
        shown = " · ".join(self.names[:3])
        return shown + (f"  외 {len(self.names) - 3}종" if len(self.names) > 3 else "")

    @property
    def parents(self) -> list[str]:
        """'부모 > 항목명' 열쇠의 부모들. 목코드 있는 묶음은 빈 목록."""
        return sorted({key.rpartition(" > ")[0] for key in self.names if " > " in key})


@dataclass
class CodeGroup:
    """목코드 한 줄."""

    code5: str                                   # 210-01
    bimok_names: list[str] = field(default_factory=list)
    rows: int = 0
    subgroups: list[NameGroup] = field(default_factory=list)

    @property
    def digits(self) -> str:
        return self.code5.replace("-", "")

    @property
    def no_code(self) -> bool:
        """설명서에 목코드가 없는 묶음(산출내역 표에서 읽은 줄).

        앞 5자리를 모르므로 한 값으로 일괄 확정하면 서로 다른 비목이 같은 코드를
        조용히 받는다. 항목명마다 7자리를 따로 정해야 한다.
        """
        return not self.code5

    @property
    def split(self) -> bool:
        """항목명에 따라 결론이 갈리는가. 갈리면 펼쳐서 보여야 한다."""
        return len({group.code for group in self.subgroups}) > 1

    @property
    def settled_rows(self) -> int:
        """비목이 채워진 행 수. 빈 행을 하나하나 세서 뺀다.

        같은 항목명이라도 부모에 따라 작년 기록이 있는 줄과 없는 줄이 섞인다. 묶음
        단위로 '미확정이면 전부 미확정'으로 세면 화면 숫자(773행)가 경고 문구(663행)와
        어긋났다. 둘은 같은 행을 세야 한다.
        """
        return self.rows - sum(group.open_rows for group in self.subgroups)

    @property
    def settled(self) -> bool:
        return all(group.settled for group in self.subgroups)

    @property
    def code(self) -> str:
        """갈리지 않을 때의 단일 결론."""
        codes = {group.code for group in self.subgroups}
        return next(iter(codes)) if len(codes) == 1 else ""

    @property
    def source(self) -> str:
        return self.subgroups[0].source if len(self.subgroups) == 1 else ""

    @property
    def name_count(self) -> int:
        """항목명 가짓수. '부모 > 항목명' 열쇠는 항목명만 센다."""
        return len({key.rpartition(" > ")[2] for group in self.subgroups for key in group.names})

    @property
    def bimok_label(self) -> str:
        if self.no_code:
            return "목코드 없음"
        shown = " · ".join(self.bimok_names[:2])
        return shown + (f" 외 {len(self.bimok_names) - 2}" if len(self.bimok_names) > 2 else "")


def override_key(items, index: int) -> str:
    """담당자가 정한 값을 붙이는 열쇠. 목코드가 있으면 항목명, 없으면 '부모 > 항목명'.

    목코드가 없는 줄은 앞 5자리가 비목을 갈라 주지 않는다. '기본'이라는 이름은
    '강사수당 > 기본'(2100605)에도 '심사수당 > 기본'(2100607)에도 나온다. 이름만으로
    한 값을 넣으면 99행이 한 코드를 조용히 받는다. 부모까지 열쇠에 넣는다.
    """
    from crosscheck import parent_name

    item = items[index]
    name = normalise(item.name)
    if item.bimok_code5:
        return name
    parent = normalise(parent_name(items, index))
    return f"{parent} > {name}" if parent else name


def build_groups(session) -> list[CodeGroup]:
    """세션의 모든 산출근거 행을 목코드 → 항목명으로 묶는다.

    확정된 것도 함께 담는다. 담당자가 자동으로 정해진 값을 눈으로 확인하고
    필요하면 고칠 수 있어야 하기 때문이다.

    두 가지를 `session.unsettled` 와 똑같은 잣대로 고른다. 화면의 묶음과 경고 문구가
    같은 행을 세야 '미확정 631행인데 확정할 줄이 0종'이 되지 않는다.

      - 재배정(총액배분·재원배분) 줄은 넣지 않는다. 이 과에서 입력하지 않아 확정할
        일이 없는데, 넣으면 영원히 '미확정'으로 남아 할 일처럼 보인다.
      - 목코드가 없는 줄(산출내역 표에서 읽은 줄)은 버리지 않고 '목코드 없음' 한
        묶음으로 모은다. 예전에는 여기서 빠져 비목 화면에 한 줄도 나오지 않았고,
        2027 설명서에서는 미확정 631행 전부가 확정할 방법 없이 숨어 있었다.
    """
    buckets: dict[str, dict] = {}
    for project in session.projects:
        if project.handover:
            continue
        for index, item in enumerate(project.items):
            if not item.is_leaf or item.handover:
                continue
            bucket = buckets.setdefault(item.bimok_code5, {"bimok": [], "names": {}})
            if item.bimok_name and item.bimok_name not in bucket["bimok"]:
                bucket["bimok"].append(item.bimok_name)
            name = override_key(project.items, index)
            code, _detail, source = session.code_for(project, index)
            resolution = session.resolutions.get((project.name, index))
            entry = bucket["names"].setdefault(name, {"rows": 0, "open": 0, "code": code, "source": source,
                                                      "candidates": list(getattr(resolution, "candidates", []) or [])})
            entry["rows"] += 1
            entry["open"] += 0 if code else 1
            # 같은 항목명인데 결론이 다르면 미확정 쪽을 남겨 담당자가 보게 한다.
            if entry["code"] and not code:
                entry["code"], entry["source"] = "", source

    groups: list[CodeGroup] = []
    for code5, bucket in buckets.items():
        merged: dict[tuple[str, str], NameGroup] = {}
        for name, entry in bucket["names"].items():
            key = (entry["code"], entry["source"] if not entry["code"] else "")
            group = merged.get(key)
            if group is None:
                group = NameGroup(code5=code5, code=entry["code"], source=entry["source"],
                                  candidates=entry["candidates"])
                merged[key] = group
            group.names.append(name)
            group.rows += entry["rows"]
            group.open_rows += entry["open"]
        for group in merged.values():
            group.names.sort()
        subgroups = sorted(merged.values(), key=lambda one: (-one.rows, one.names[0] if one.names else ""))
        groups.append(CodeGroup(code5=code5, bimok_names=bucket["bimok"],
                                rows=sum(one.rows for one in subgroups), subgroups=subgroups))

    # 미확정이 많은 목코드를 위로 올린다. 할 일이 먼저 보여야 한다.
    groups.sort(key=lambda one: (one.settled, -(one.rows - one.settled_rows), -one.rows))
    return groups


def summary(groups: list[CodeGroup]) -> dict:
    return {
        "목코드": len(groups),
        "남은 목코드": sum(1 for group in groups if not group.settled),
        "행": sum(group.rows for group in groups),
        "확정 행": sum(group.settled_rows for group in groups),
        "갈리는 목코드": sum(1 for group in groups if group.split),
    }
