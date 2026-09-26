"""세 목표를 한 흐름으로 엮는다. 화면(tkinter)과 분리해 두어 테스트할 수 있다.

    설명서(hwpx) ─┐
    작년 다운로드 ─┼─▶ 비목 확정 ─▶ 입력본.xlsx ─▶ [담당자가 K-에듀파인에 입력]
    UBIS 검토조서 ─┘                                     │
                                                         ▼
                                            입력 후 다운로드본 ─▶ 목표 2B
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from bimok_resolver import BimokResolver, normalise
from crosscheck import Report, check_bimok, check_downloaded, check_generated, check_plan, diff_forms
from edufine_form import OutputRow, basis_text, prefix, write_form
from plan_parser import PlanProject, parse_plan
from ubis_review import read_adjustments, read_requests

DEFAULT_DETAIL = "본청"


@dataclass
class Session:
    plan_path: str = ""
    ubis_path: str = ""
    last_year_path: str = ""          # 서식 본. 이 파일의 모양으로 입력본을 만든다.
    extra_last_year: list = field(default_factory=list)   # 비목 학습만 하는 추가 확정본
    class_path: str = ""          # 세출예산 사업별 분류(별표 3). 선택.
    card_path: str = ""           # 단위과제카드 목록. 넣으면 내가 입력할 사업만 남긴다. 선택.

    projects: list[PlanProject] = field(default_factory=list)
    ubis: dict = field(default_factory=dict)
    resolver: BimokResolver = field(default_factory=BimokResolver)
    resolutions: dict = field(default_factory=dict)
    # 담당자가 확정한 값. 목코드 단위와, 갈리는 목코드의 항목명 단위 두 가지다.
    overrides_by_code: dict = field(default_factory=dict)          # "210-01" → 7자리
    overrides_by_name: dict = field(default_factory=dict)          # ("210-01", "운영용품") → 7자리
    plan_report: Report = field(default_factory=Report)
    bimok_report: Report = field(default_factory=Report)
    classes: object = None
    class_report: Report = field(default_factory=Report)
    cards: object = None
    card_report: Report = field(default_factory=Report)
    all_projects: list = field(default_factory=list)   # 카드로 좁히기 전 전체
    adjustments: dict = field(default_factory=dict)     # UBIS 조정액(D). 본예산 대조에서 차액 설명용

    # 추경. 설명서가 추경 서식이면 입력본 대신 점검만 한다.
    previous_ubis_path: str = ""      # 직전 차수 검토조서. 비우면 UBIS 파일 옆에서 찾는다.
    supplement: object = None         # supplement.SuppDocument
    supplement_ubis: object = None    # supplement.UbisSupp
    supplement_report: Report = field(default_factory=Report)   # UBIS 대조 + 직전 차수 연결
    writing_report: object = None     # writing_check.WritingReport — 작성요령 점검(확인 필요, 막지 않음)
    previous_used: str = ""           # 직전 차수 연결에 실제로 쓴 검토조서

    @property
    def is_supplement(self) -> bool:
        return self.supplement is not None

    # ---------------------------------------------------------------- 읽기
    def load(self, progress=None) -> None:
        def step(message: str, ratio: float) -> None:
            if progress:
                progress(message, ratio)

        if not self.plan_path:
            raise ValueError("사업설명서 HWPX를 선택해 주세요.")
        from supplement import is_supplement_file

        if is_supplement_file(self.plan_path):
            self._load_supplement(step)
            return
        if not self.last_year_path:
            raise ValueError("작년 K-에듀파인 다운로드 파일을 선택해 주세요. 서식과 비목을 여기서 배웁니다.")
        step("사업설명서를 읽는 중…", 0.05)
        self.projects = parse_plan(self.plan_path)
        if not self.projects:
            raise ValueError("설명서에서 사업을 찾지 못했습니다. 부서별사업별 설명서가 맞는지 확인해 주세요.")
        self.all_projects = list(self.projects)
        if self.card_path:
            # K-에듀파인은 자기 카드에 딸린 사업만 입력할 수 있다. 과 전체를 늘어놓으면
            # 자기 것을 찾는 데만 시간이 든다. 넣었으면 그 범위로 좁힌다.
            from program_cards import check_cards, load_cards

            self.cards = load_cards(self.card_path)
            self.card_report = check_cards(self.projects, self.cards)
            self.projects = [one for one in self.projects if self.cards.find(one.name)]
            if not self.projects:
                raise ValueError("과제카드 목록의 사업이 설명서에 하나도 없습니다. "
                                 "다른 과의 설명서이거나 다른 해의 카드 목록인지 확인해 주세요.")
            step(f"과제카드 {len(self.cards)}장 · 내가 입력할 사업 {len(self.projects)}개만 남겼습니다.", 0.4)
        leaves = sum(1 for one in self.projects for item in one.items if item.is_leaf)
        step(f"사업 {len(self.projects)}개 · 산출근거 {leaves:,}행을 읽었습니다.", 0.45)

        self.resolver = BimokResolver()
        learned = self.resolver.learn(self.last_year_path)
        for extra in self.extra_last_year:
            if extra and extra != self.last_year_path:
                learned += self.resolver.learn(extra)
        step(f"확정본 {len(self.resolver.sources)}부에서 비목 {learned:,}행을 배웠습니다.", 0.65)

        if self.ubis_path:
            self.ubis = read_requests(self.ubis_path)
            self.adjustments = read_adjustments(self.ubis_path)
            step(f"UBIS 세출요구 {len(self.ubis):,}개 사업을 읽었습니다.", 0.8)
        else:
            self.ubis = {}
            step("UBIS 검토조서가 없어 금액 대조는 건너뜁니다.", 0.8)

        if self.class_path:
            from program_classes import check_projects, load_classification

            self.classes = load_classification(self.class_path)
            self.class_report = check_projects(self.projects, self.classes)
            step(f"사업별 분류 {len(self.classes):,}항목과 맞췄습니다.", 0.85)
        else:
            self.classes, self.class_report = None, Report()

        step("금액을 검사하는 중…", 0.88)
        self.plan_report = check_plan(self.projects)
        step("작성요령대로 썼는지 보는 중…", 0.91)
        self._check_writing()
        step("비목을 맞추는 중…", 0.95)
        self.bimok_report, self.resolutions = check_bimok(self.projects, self.resolver)
        step("끝났습니다.", 1.0)

    def _check_writing(self) -> None:
        """작성요령 점검. 실패해도 불러오기는 막지 않는다 — 금액 검산이 본업이다."""
        from edufine_form import read_form
        from writing_check import WritingReport, check_writing

        rows = []
        if not self.is_supplement:
            for path in [self.last_year_path, *self.extra_last_year]:
                try:
                    rows += read_form(path) if path else []
                except Exception:  # noqa: BLE001 - 예산현액 조회 서식 등은 단가 비교에 못 쓴다
                    pass
        try:
            report = check_writing(self.plan_path, self.projects, rows)
        except Exception as error:  # noqa: BLE001
            report = WritingReport()
            report.skipped.append(str(error))
        # 과제카드로 좁혔으면 내 사업 것만 남긴다.
        names = {project.name for project in self.projects}
        report.issues = [issue for issue in report.issues if issue.project in names]
        self.writing_report = report

    def _load_supplement(self, step) -> None:
        """추경 설명서. 산출식이 증감분만 있어 입력본은 만들지 않고 점검만 한다.

        K-에듀파인 작년 파일은 쓰지 않는다(비목·서식은 입력본용이다). UBIS 검토조서는
        추경용이어야 한다 — 본예산 검토조서의 '요구액'은 사업 전체라 증감과 견줄 수 없다.
        """
        from supplement import (check_supplement, compare_previous, compare_supplement, find_previous_review,
                                is_supplement_review, parse_supplement, read_supplement_review)

        step("추경 사업별 설명서를 읽는 중…", 0.1)
        document = parse_supplement(self.plan_path)
        if not document.projects:
            raise ValueError("추경 설명서에서 사업 쪽을 찾지 못했습니다. 사업별 설명서가 맞는지 확인해 주세요.")
        self.supplement = document
        self.projects = [PlanProject(name=one.name, number=one.number, order=index, heading=one.heading,
                                     policy=one.policy, unit=one.unit, program=one.program,
                                     organisation=one.organisation, request=one.total_cost,
                                     year_request=one.after, supplement=True)
                         for index, one in enumerate(document.projects, start=1)]
        self.all_projects = list(self.projects)
        groups = sum(len(table.groups) for one in document.projects for table in one.tables)
        step(f"{document.round_label or '추경'} 사업 {len(document.projects)}개 · 증감 줄 {groups:,}개를 "
             "읽었습니다.", 0.4)
        self.plan_report = check_supplement(document)
        step("설명서 안의 증감·산출식·합계를 맞춰 봤습니다.", 0.6)
        self._check_writing()

        report = Report()
        self.ubis, self.supplement_ubis = {}, None
        if self.ubis_path:
            if not is_supplement_review(self.ubis_path):
                raise ValueError("설명서는 추경인데 3번 UBIS 검토조서는 본예산용입니다.\n\n"
                                 "추경 검토조서(서식1)를 넣어 주세요. 본예산 검토조서의 요구액은 사업 전체 금액이라 "
                                 "추경 증감과 견줄 수 없습니다.")
            self.supplement_ubis = read_supplement_review(self.ubis_path)
            self.ubis = {row.name: row.change for row in self.supplement_ubis.rows}
            report.extend(compare_supplement(document, self.supplement_ubis))
            # 총괄 표와 사업 쪽이 어긋날 때 오류인지 가리는 데 UBIS 를 쓴다. 다시 검사한다.
            self.plan_report = check_supplement(document, self.supplement_ubis)
            step(f"UBIS 추경 검토조서 {len(self.supplement_ubis.rows)}개 줄과 맞췄습니다.", 0.8)
            previous = self.previous_ubis_path or find_previous_review(self.ubis_path)
            self.previous_used = previous
            if previous:
                report.extend(compare_previous(self.ubis_path, previous))
                step("직전 차수 검토조서와 기정액을 이어 봤습니다.", 0.9)
        else:
            step("UBIS 검토조서가 없어 금액 대조를 건너뜁니다.", 0.8)
        # 추경은 입력본을 만들지 않으므로 '입력본을 막는 오류'라는 말이 맞지 않는다.
        for issue in self.plan_report.issues + report.issues:
            issue.blocking = False
        self.supplement_report = report
        self.bimok_report, self.resolutions = Report(), {}
        self.class_report, self.card_report = Report(), Report()
        step("끝났습니다.", 1.0)

    # ---------------------------------------------------------------- 비목
    def code_for(self, project: PlanProject, index: int) -> tuple[str, str, str]:
        """확정된 (비목코드, 상세내역, 근거). 미확정이면 코드가 빈 문자열이다.

        담당자가 정한 값이 작년 학습보다 앞선다. 항목명 단위가 목코드 단위보다 앞선다.
        """
        from bimok_groups import override_key

        item = project.items[index]
        by_name = self.overrides_by_name.get((item.bimok_code5, override_key(project.items, index)))
        if not by_name and not item.bimok_code5:
            by_name = self.overrides_by_name.get(("", normalise(item.name)))
        if by_name:
            return by_name, self.detail_for(by_name), "담당자 확정 (항목명)"
        by_code = self.overrides_by_code.get(item.bimok_code5)
        if by_code:
            return by_code, self.detail_for(by_code), "담당자 확정 (목코드)"
        resolution = self.resolutions.get((project.name, index))
        if resolution and resolution.settled:
            return resolution.code, resolution.detail or DEFAULT_DETAIL, resolution.source
        return "", "", resolution.source if resolution else "미확정"

    def detail_for(self, code: str) -> str:
        counts = self.resolver.detail_by_code.get(code)
        if counts and len(counts) == 1:
            return next(iter(counts))
        return DEFAULT_DETAIL

    @property
    def unsettled(self) -> list[tuple[PlanProject, int, object]]:
        if self.is_supplement:
            return []
        rows = []
        for project in self.projects:
            for index, item in enumerate(project.items):
                if not item.is_leaf:
                    continue
                if item.handover or project.handover:
                    continue          # 이 과에서 입력하지 않는 줄. 확정할 일이 없다.
                if not self.code_for(project, index)[0]:
                    rows.append((project, index, item))
        return rows

    # ---------------------------------------------------------------- 출력
    def card_of(self, project) -> str:
        """이 사업이 올라갈 단위과제카드 이름. 카드 목록이 없으면 사업명 자체를 쓴다.

        K-에듀파인은 카드 한 장을 열고 그 안의 사업들을 차례로 입력한다. 실제
        다운로드 파일도 카드 한 장 분량이고, 카드 하나에 사업이 넷 들어 있으면
        순번이 1부터 55까지 끊기지 않고 이어진다. 카드 목록이 없으면 어느 사업끼리
        한 카드인지 알 수 없으므로 사업 하나를 카드 하나로 본다.
        """
        if self.cards:
            found = self.cards.find(project.name)
            if found is not None:
                return found.name
        return project.name

    def output_rows(self, handover: bool = False) -> list[tuple[str, list[OutputRow]]]:
        """카드별로 묶은 출력 행. handover=True 면 이 과에서 입력하지 않는 사업만."""
        groups: dict[str, list[OutputRow]] = {}
        order: list[str] = []
        for project in self.projects:
            if bool(project.handover) != handover:
                continue
            card = self.card_of(project)
            if card not in groups:
                groups[card] = []
                order.append(card)
            rows = groups[card]
            # 한 카드 안에서 사업은 1. 2. 3. 으로 번호가 붙는다.
            top = sum(1 for row in rows if row.level == 1) + 1
            rows.append(OutputRow(level=1, name=f"{top}. {project.name}", amount=project.base,
                                  note=project.handover, charge=project.charge, card=card))
            counters: dict[int, int] = {}
            for index, item in enumerate(project.items):
                level = min(item.depth + 1, 6)
                counters[level] = counters.get(level, 0) + 1
                for deeper in [key for key in counters if key > level]:
                    counters.pop(deeper)
                code, detail, source = self.code_for(project, index)
                rows.append(OutputRow(
                    charge=project.charge,
                    card=card,
                    level=level,
                    name=f"{prefix(level, counters[level])} {item.name}".strip(),
                    code=code,
                    detail=detail if item.is_leaf else "",
                    basis=basis_text(item.formula, item.amount) if item.is_leaf else "",
                    amount=item.amount,
                    note=(project.handover if project.handover
                          else "" if code or not item.is_leaf else source),
                ))
        return [(card, groups[card]) for card in order]

    def project_by_name(self, name: str):
        return next((one for one in self.projects if one.name == name), None)

    def label_of(self, name: str) -> str:
        """'15. 영유아교육내실화 지원'. 없는 이름이면 그대로 돌려준다."""
        found = self.project_by_name(name)
        return found.label if found else name

    def groups(self) -> list:
        """비목 확정 화면이 쓰는 목코드 묶음."""
        from bimok_groups import build_groups

        return build_groups(self)

    def known_codes(self) -> list[tuple[str, str]]:
        """고를 수 있는 7자리 목록 (코드, 비목명). 작년 확정본 + 담당자가 이미 정한 값.

        확정본에 비목명이 없으면 설명서의 '일반수용비(210-01)' 처럼 앞 5자리가 같은
        비목명을 빌려 적는다. 5자리에 비목명이 둘 이상이면 짐작하지 않고 비워 둔다.
        """
        by_digits: dict[str, set] = {}
        for project in self.projects:
            for item in project.items:
                if item.bimok_code5 and item.bimok_name:
                    by_digits.setdefault(item.code5, set()).add(item.bimok_name)
        found = dict(self.resolver.known_codes())
        for code in list(self.overrides_by_code.values()) + list(self.overrides_by_name.values()):
            found.setdefault(code, "")
        rows = []
        for code, name in sorted(found.items()):
            if not name:
                names = by_digits.get(code[:5], set())
                name = next(iter(names)) if len(names) == 1 else ""
            rows.append((code, name))
        return rows

    def apply_remembered(self, stored: dict) -> int:
        """지난 실행에서 확정한 비목을 다시 적용한다."""
        applied = 0
        for key, code in (stored or {}).items():
            if not code:
                continue
            if "|" in key:
                code5, name = key.split("|", 1)
                self.overrides_by_name[(code5, name)] = code
            else:
                self.overrides_by_code[key] = code
            applied += 1
        return applied

    def remembered(self) -> dict:
        stored = {code5: code for code5, code in self.overrides_by_code.items() if code}
        for (code5, name), code in self.overrides_by_name.items():
            if code:
                stored[f"{code5}|{name}"] = code
        return stored

    def blocking(self) -> list[str]:
        """파일을 아예 만들지 않는 사유.

        보이지 않는 결함만 여기 넣는다. 금액이 틀리거나 분류가 틀리면 결과 파일만
        봐서는 알 수 없고, 그대로 입력되면 예산이 틀어진다.
        """
        if self.is_supplement:
            return ["추경 설명서는 입력본을 만들지 않습니다(점검 전용)"]
        reasons = []
        blocking = self.plan_report.blocking_errors
        if blocking:
            reasons.append(f"설명서 금액 오류 {len(blocking)}건")
        if self.class_report.errors:
            reasons.append(f"사업 분류 오류 {len(self.class_report.errors)}건")
        return reasons

    def warnings(self) -> list[str]:
        """확인만 받고 진행해도 되는 사유.

        비목 빈칸은 **보이는** 결함이다. 줄은 파일에 그대로 있고 칸만 비어 있어
        입력하다 반드시 마주친다. 조용히 넘어갈 수 없으므로 차단까지 하지 않는다.
        """
        rows = self.unsettled
        if not rows:
            return []
        return [f"비목 미확정 {len(rows):,}행 ({self.unsettled_kinds()}) — "
                "빈칸으로 나가며 K-에듀파인에서 직접 골라야 합니다"]

    def unsettled_kinds(self) -> str:
        """미확정 행이 몇 가지 결정으로 모이는지. '목코드 3종 · 목코드 없는 항목명 12종'.

        목코드가 있는 줄은 목코드 하나(갈리면 항목명 몇 개)로 정하고, 목코드가 없는
        줄은 항목명마다 정한다. 둘을 한 숫자로 세면 담당자가 할 일의 양을 잘못 읽는다.
        """
        coded, named = 0, 0
        for group in self.groups():
            if group.settled:
                continue
            if group.no_code:
                named += len({key.rpartition(" > ")[2] for sub in group.subgroups if not sub.settled
                              for key in sub.names})
            else:
                coded += 1
        parts = []
        if coded:
            parts.append(f"목코드 {coded}종")
        if named:
            parts.append(f"목코드 없는 항목명 {named}종")
        return " · ".join(parts) or "0종"

    def save(self, output: str, force: bool = False, allow_unsettled: bool = False) -> str:
        """입력본을 만든다. 보이지 않는 결함이 있으면 만들지 않는다.

        완성품처럼 보이는 파일을 경고와 함께 내주면 담당자는 그대로 입력한다.
        파서가 조용히 항목을 놓쳤을 때 그 줄은 파일에 아예 없고, 아무도 모른 채
        예산이 빠진다. 그런 결함은 경고가 아니라 차단으로 막는다.

        비목 빈칸은 다르다. 줄은 그대로 있고 칸만 비어 담당자가 입력하다 마주친다.
        `allow_unsettled` 로 확인만 받고 진행한다.
        """
        if self.is_supplement:
            raise ValueError("추경 설명서는 입력본을 만들지 않습니다. 산출식이 증감분만 있어 "
                             "K-에듀파인 입력본의 원천이 될 수 없습니다.")
        reasons = self.blocking()
        if reasons and not force:
            raise ValueError("다음을 해결해야 입력본을 만들 수 있습니다.\n - " + "\n - ".join(reasons))
        if self.unsettled and not (allow_unsettled or force):
            raise ValueError("비목 미확정 행이 있습니다. 확인 후 진행하려면 allow_unsettled 를 켜세요.")
        write_form(self.last_year_path, output, self.output_rows(),
                   handover=self.output_rows(handover=True))
        return output

    # ---------------------------------------------------------------- 검증
    def verify_generated(self, generated: str) -> Report:
        """목표 2A — 설명서 = UBIS = 생성한 입력본."""
        if self.is_supplement:
            return self.supplement_report
        report = check_generated(self.projects, self.ubis, generated)
        self._explain_adjustments(report)
        return report

    def verify_supplement_form(self, downloaded: str) -> Report:
        """추경 — K-에듀파인에 입력하고 내려받은 파일을 설명서 · UBIS 와 맞춘다."""
        from supplement_form import check_form, compare_form, is_supplement_form, read_supplement_form

        if not is_supplement_form(downloaded):
            raise ValueError("추경 설명서로 불러왔는데 이 파일은 본예산 서식입니다. K-에듀파인 추경 "
                             "세출예산요구내역을 내려받아 넣어 주세요. ('추경요구(②-①)' 열이 있는 파일)")
        form = read_supplement_form(downloaded)
        report = check_form(form)
        compare_form(form, self.supplement, self.supplement_ubis, report)
        return report

    def _explain_adjustments(self, report: Report) -> None:
        """설명서와 UBIS 요구액의 차이가 UBIS 조정액(D)과 똑같으면 오류가 아니다 — 목록에서 뺀다.

        2025·2026 본예산에서 이 모양의 '오류' 20건이 전부 여기에 해당했다(차액 = 조정액).
        설명서가 조정 전후 어느 금액으로 적혔는지만 확인하면 된다.
        """
        if not self.adjustments:
            return
        from crosscheck import _lookup, _by_name, _loose_index

        exact, loose = _by_name(self.adjustments), _loose_index(self.adjustments)
        for issue in report.issues:
            if issue.severity != "오류" or "UBIS" not in issue.message or issue.gap is None:
                continue
            adjust, _alias = _lookup(exact, loose, issue.project)
            if adjust and abs(abs(issue.gap) - abs(adjust)) < 0.5:
                # 차이가 조정액과 1천원도 다르지 않다. 설명서가 조정분까지 넣어 적은 것이지 틀린 게 아니다.
                issue.severity = "확인 필요"
                issue.message += f"  (= UBIS 조정액 {adjust:,.0f}천원)"
        report.dropped.extend(one for one in report.issues if one.severity != "오류")
        report.issues = [one for one in report.issues if one.severity == "오류"]

    def verify_downloaded(self, downloaded: str, generated: str = "") -> Report:
        """목표 2B — 설명서 = UBIS = 입력 후 다운로드본."""
        from supplement_form import is_supplement_form

        if self.is_supplement:
            return self.verify_supplement_form(downloaded)
        if is_supplement_form(downloaded):
            raise ValueError("이 파일은 K-에듀파인 **추경** 세출예산요구내역입니다. 추경 사업별 설명서로 불러온 뒤 "
                             "대조해 주세요.")
        report = check_downloaded(self.projects, self.ubis, downloaded)
        if generated and Path(generated).exists():
            report.extend(diff_forms(generated, downloaded))
        return report

    # ---------------------------------------------------------------- 요약
    def summary(self) -> dict:
        leaves = [item for project in self.projects for item in project.items if item.is_leaf]
        passed = [item for project in self.projects for item in project.items
                  if item.is_leaf and (item.handover or project.handover)]
        settled = len(leaves) - len(passed) - len(self.unsettled)
        return {
            "생성 시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "사업 수": len(self.projects),
            "과제카드": len(self.cards) if self.cards else 0,
            "설명서 전체 사업": len(self.all_projects),
            "산출근거 행": len(leaves),
            "입력 대상 아님": len(passed),
            "비목 확정": settled,
            "비목 미확정": len(self.unsettled),
            "목표1 오류": len(self.plan_report.errors),
            "검사 못 한 것": len(self.plan_report.skipped),
            "작년 학습 행": self.resolver.learned_rows,
            "UBIS 사업 수": len(self.ubis),
            "분류 오류": len(self.class_report.errors),
            **(self._supplement_summary() if self.is_supplement else {}),
        }

    def _supplement_summary(self) -> dict:
        document = self.supplement
        total = document.summary_total
        return {
            "점검 종류": f"추경 점검 ({document.round_label or '추경'})",
            "추경 증감 합계(총괄)": total.change if total else None,
            "증감 줄": sum(len(table.groups) for one in document.projects for table in one.tables),
            "UBIS 대조 오류": len(self.supplement_report.errors),
        }
