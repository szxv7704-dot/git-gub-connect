"""회귀 테스트. 한 번씩 실제로 사고가 났거나 날 뻔한 지점이다. 지우지 말 것.

  - 사설영역(PUA) 불릿: 두 글리프가 서로 다른 계층(U+F02FB 4단, U+F077 5단)이다.
    같은 단계로 뭉개면 상위·하위 합계가 어긋나고, 금액은 맞아 보여 알아채기 어렵다.
  - 레벨 = 들여쓰기 ÷ 2 + 1, 순번 = 4행부터 전역 일련번호.
  - 검산을 통과하지 못하면 입력본을 만들지 않는다.
"""

from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path

from bimok_resolver import BimokResolver, normalise
from converter import Session
from program_cards import Card, CardList, check_cards
from openpyxl import load_workbook
from crosscheck import (Issue, calc, check_bimok, check_plan, compare_totals, issue_row,
                        leaf_total, subtree_of)
from edufine_form import _columns, basis_text, prefix, read_form
from hwpx_native_parser import is_private_use
from plan_parser import (PlanItem, PlanPart, PlanProject, _cells, _is_basis_table,
                         _learn_bullets, _learn_outline_bullets, _parse_basis_rows,
                         _parse_outline_rows, parse_plan)
from hwpx_native_parser import native_chunks

HERE = Path(__file__).parent


def find_sample(name: str) -> Path:
    """표본 파일을 찾는다. 실제 예산자료라 배포 폴더에 복사하지 않는다."""
    for folder in SAMPLE_FOLDERS:
        if (folder / name).exists():
            return folder / name
    return HERE / name


def _sample_folders() -> list:
    folders = [HERE, HERE.parent.parent / "3 예산이 뭐예요" / "ubis꺼져",
               HERE.parent.parent / "3 예산이 뭐예요" / "9.24"]
    env = os.environ.get("UBIS_SAMPLES")
    if env:
        # 여러 폴더는 경로 구분자(윈도 ; · 리눅스 :)로 잇는다.
        folders = [Path(one) for one in env.split(os.pathsep) if one] + folders
    return folders


SAMPLE_FOLDERS = _sample_folders()


def find_sample_like(prefix: str) -> Path:
    """이름 앞부분으로 표본을 찾는다. UBIS 는 내려받은 시각을 파일 이름 끝에 붙인다."""
    for folder in SAMPLE_FOLDERS:
        if not folder.exists():
            continue
        found = sorted(one for one in folder.glob(prefix + "*") if not one.name.startswith("~$"))
        if found:
            return found[-1]
    return HERE / (prefix + "(없음)")


PLAN = find_sample("(제거)(본예산)부서별사업별 설명서(작성)_2026.9.17_17_42_49.hwpx")
LAST_YEAR = find_sample("(에듀파인)2026세출예산요구내역.xlsx")
UBIS = find_sample("(본예산세출요구)검토조서(서식1)_2026.9.17_17_24_48.xlsx")
CLASSES = find_sample("목내용세출예산 사업별 분류.hwpx")


class CalculationTests(unittest.TestCase):
    def test_unit_price_in_won(self):
        self.assertEqual(calc("50,000원×5종×2회"), 500)
        self.assertEqual(calc("7,000,000원×1개 기관"), 7000)
        self.assertEqual(calc("160,000원×2명×2시간×2회"), 1280)

    def test_thousand_won_unit_still_works(self):
        self.assertEqual(calc("50천원×5종×2회"), 500)

    def test_trailing_total_is_not_multiplied(self):
        self.assertEqual(calc("50,000원×5종×2회=500천원"), 500)

    def test_percentage(self):
        self.assertAlmostEqual(calc("1,470,000천원×3%"), 44100)


class FormRuleTests(unittest.TestCase):
    def test_prefix_by_level(self):
        self.assertEqual(prefix(1, 1), "1.")
        self.assertEqual(prefix(2, 1), "가.")
        self.assertEqual(prefix(3, 2), "2)")
        self.assertEqual(prefix(4, 2), "나)")
        self.assertEqual(prefix(5, 1), "(1)")
        self.assertEqual(prefix(6, 1), "(가)")

    def test_zero_amount_basis(self):
        self.assertEqual(basis_text("", 0), "0=")
        self.assertEqual(basis_text("50,000원×2회", 100), "50,000원×2회=")

    def test_normalise_strips_markers(self):
        self.assertEqual(normalise("(1) 운영용품"), "운영용품")
        self.assertEqual(normalise("가) 운영용품"), "운영용품")
        self.assertEqual(normalise("  2) 협의회"), "협의회")


@unittest.skipUnless(LAST_YEAR.exists(), "작년 확정본이 없으면 건너뛴다")
class LastYearFormTests(unittest.TestCase):
    def test_level_and_sequence_rule_holds(self):
        rows = read_form(LAST_YEAR)
        self.assertGreater(len(rows), 0)
        for order, row in enumerate(rows, start=1):
            self.assertEqual(row.level, row.indent // 2 + 1, f"{row.row}행 레벨")
            self.assertEqual(row.sequence, order, f"{row.row}행 순번")

    def test_merged_header_finds_request_columns(self):
        """1행이 병합돼 있어 오른쪽 셀이 비면 '요구' 산출기초·금액을 못 찾는다.

        전년도 쌍과 요구 쌍이 나란히 있어, 병합된 윗 칸을 이어 읽지 않으면 전년도
        열을 요구 열로 착각한다. 값이 채워졌는지는 표본 나름이라 묻지 않는다.
        """
        book = load_workbook(LAST_YEAR, data_only=True)
        self.addCleanup(book.close)
        sheet = book.active
        columns = _columns(sheet)
        for key in ("레벨", "순번", "사업항목", "비목", "상세내역", "산출기초", "금액"):
            self.assertIn(key, columns)
        carried = ""
        for column in range(1, columns["산출기초"] + 1):
            top = str(sheet.cell(1, column).value or "").replace(" ", "")
            if top:
                carried = top
        self.assertIn("요구", carried, "전년도 산출기초를 요구 산출기초로 잘못 잡았다")
        self.assertGreater(columns["금액"], columns["상세내역"])

    def test_resolver_learns_five_digit_prefix(self):
        resolver = BimokResolver()
        resolver.learn(str(LAST_YEAR))
        settled = resolver.resolve("210-01", "운영용품")
        self.assertEqual(settled.code, "2100143")
        self.assertTrue(settled.code.startswith("21001"))


@unittest.skipUnless(PLAN.exists(), "설명서 표본이 없으면 건너뛴다")
class PlanParsingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.projects = parse_plan(str(PLAN))
        cls.leaves = [item for project in cls.projects for item in project.items if item.is_leaf]

    def test_project_and_leaf_counts(self):
        self.assertEqual(len(self.projects), 69)
        self.assertEqual(len(self.leaves), 1118)

    def test_every_leaf_carries_a_cost_code(self):
        """설명서에 비목이 있으므로 사전이 필요 없다. 하나라도 비면 전제가 깨진 것이다."""
        self.assertEqual([item.name for item in self.leaves if not item.code5], [])

    def test_private_use_bullets_map_to_distinct_levels(self):
        chunks = native_chunks(str(PLAN))
        tables = [_cells(chunk["text"]) for chunk in chunks
                  if chunk["type"] == "table" and _is_basis_table(_cells(chunk["text"]))]
        bullets = _learn_bullets(tables)
        self.assertEqual(bullets.get(chr(0xF02FB)), 4)
        self.assertEqual(bullets.get(chr(0xF077)), 5)
        self.assertTrue(is_private_use(chr(0xF02FB)))

    def test_depths_reach_five_levels(self):
        depths = {item.depth for item in self.leaves}
        self.assertTrue({1, 2, 3, 4, 5} >= depths, depths)
        self.assertIn(5, depths, "5단 항목이 4단으로 뭉개지면 계층이 한 단 사라진다")

    def test_plan_internal_amounts_mostly_agree(self):
        """목표 1. 표본에서 남는 지적은 실제 확인 대상 1건뿐이어야 한다."""
        report = check_plan(self.projects)
        self.assertLessEqual(len(report.errors), 1, [issue.message for issue in report.errors])


@unittest.skipUnless(PLAN.exists() and LAST_YEAR.exists(), "표본이 없으면 건너뛴다")
class OutputTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.session = Session(plan_path=str(PLAN),
                              ubis_path=str(UBIS) if UBIS.exists() else "",
                              last_year_path=str(LAST_YEAR))
        cls.session.load()

    def test_invisible_faults_block_but_blank_cost_codes_only_warn(self):
        """차단과 경고를 가르는 기준은 '결과 파일만 보고 알 수 있는가' 다.

        금액이 틀리면 파일만 봐서는 알 수 없다 → 차단.
        비목이 비면 입력하다 반드시 마주친다 → 경고.
        """
        blocking = " ".join(self.session.blocking())
        self.assertNotIn("비목", blocking, "비목 빈칸은 차단 사유가 아니다")
        if self.session.unsettled:
            self.assertTrue(any("비목" in one for one in self.session.warnings()))

    def test_unsettled_codes_need_an_explicit_go_ahead(self):
        """확인 없이는 비목 빈칸인 채로 만들어지지 않는다."""
        session = Session(plan_path=str(PLAN), ubis_path="", last_year_path=str(LAST_YEAR))
        session.load()
        if not session.unsettled:
            self.skipTest("표본에 미확정이 없다")
        session.plan_report.issues = [one for one in session.plan_report.issues
                                      if one.severity != "오류"]
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "out.xlsx"
            with self.assertRaises(ValueError):
                session.save(str(target))
            session.save(str(target), allow_unsettled=True)
            self.assertTrue(target.exists())

    def test_forced_output_follows_level_and_sequence_rules(self):
        """레벨 = 들여쓰기 ÷ 2 + 1 · 순번은 **단위과제카드마다** 1부터.

        실제 다운로드 파일은 카드 한 장 분량이고 순번이 1부터 이어진다. 카드 하나에
        사업이 넷 들어 있어도 끊기지 않는다(55행짜리는 1~55). 문서 전체에 걸쳐
        번호를 매기면 실제 화면과 어긋난다.
        """
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "out.xlsx"
            self.session.save(str(target), force=True)
            rows = read_form(str(target))
            self.assertGreater(len(rows), 0)
            self.assertEqual(rows[0].sequence, 1)
            previous = None
            for row in rows:
                self.assertEqual(row.level, row.indent // 2 + 1)
                if previous is not None and row.sequence != previous + 1:
                    self.assertEqual(row.sequence, 1, f"{row.row}행: 카드가 바뀔 때만 1로 돌아간다")
                    self.assertEqual(row.level, 1, f"{row.row}행: 카드는 레벨1 행에서 시작한다")
                previous = row.sequence
            self.assertGreater(len({row.sequence for row in rows}), 1)

    def test_settled_rows_carry_code_detail_and_basis(self):
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "out.xlsx"
            self.session.save(str(target), force=True)
            filled = [row for row in read_form(str(target)) if row.code]
            self.assertTrue(filled)
            for row in filled[:50]:
                self.assertTrue(row.basis.endswith("="))
                self.assertTrue(row.detail)


class FormattingTests(unittest.TestCase):
    """금액은 사람이 자릿수를 세지 않아도 되게 적는다."""

    def test_money_has_thousand_separators_and_no_decimal(self):
        from ui_common import money
        self.assertEqual(money(3240000.0), "3,240,000")
        self.assertEqual(money(0), "0")
        self.assertEqual(money(None), "")
        self.assertEqual(money(15680), "15,680")

    def test_gap_carries_a_sign(self):
        from ui_common import signed
        self.assertEqual(signed(2200000), "+2,200,000")
        self.assertEqual(signed(-34500), "\u2212" + "34,500")
        self.assertEqual(signed(0), "0")
        self.assertEqual(signed(None), "")


class SettingsTests(unittest.TestCase):
    def test_round_trip_and_missing_file(self):
        import settings as store
        data = store.load()
        self.assertIn("recent", data)
        self.assertIn("bimok", data)
        self.assertEqual(store.bimok_key("210-01"), "210-01")
        self.assertEqual(store.bimok_key("210-01", "운영용품"), "210-01|운영용품")


@unittest.skipUnless(PLAN.exists() and LAST_YEAR.exists(), "표본이 없으면 건너뛴다")
class GroupingTests(unittest.TestCase):
    """566행을 담당자가 감당할 수 있는 단위로 묶는다."""

    @classmethod
    def setUpClass(cls):
        cls.session = Session(plan_path=str(PLAN),
                              ubis_path=str(UBIS) if UBIS.exists() else "",
                              last_year_path=str(LAST_YEAR))
        cls.session.load()

    def test_rows_collapse_into_far_fewer_code_groups(self):
        from bimok_groups import build_groups
        groups = build_groups(self.session)
        rows = sum(group.rows for group in groups)
        # 1118행 중 40행은 총액배분 재배정(620-02) 줄이다. 이 과에서 입력하지 않으므로
        # 비목 화면에 넣지 않는다(2.13.0). 넣으면 확정할 수 없는 '미확정'이 남는다.
        self.assertEqual(rows, 1078)
        self.assertLess(len(groups), 60, "묶음이 행 수만큼 많으면 묶은 의미가 없다")

    def test_last_year_proves_one_code_can_split(self):
        """210-01 은 작년 확정본에서 2100143 과 2100122 둘을 썼다.

        목코드 하나로 묶어 한 값을 넣으면 어떤 항목이 틀린 코드를 조용히 받는다.
        이 사실이 목코드 일괄 확정을 그대로 쓰면 안 되는 이유다.
        """
        counts = self.session.resolver.by_code.get("21001", {})
        if len(counts) <= 1:
            self.skipTest("이 확정본에는 210-01 이 갈리는 기록이 없다")
        self.assertEqual(self.session.resolver.resolve("210-01", "운영용품").code, "2100143")
        unknown = self.session.resolver.resolve("210-01", "한번도없던항목명")
        self.assertFalse(unknown.settled, "모르는 항목명은 추측하지 않는다")
        self.assertEqual(sorted(unknown.candidates), ["2100122", "2100143"])

    def test_split_code_group_is_expanded_not_bulk_filled(self):
        from bimok_groups import build_groups
        if len(self.session.resolver.by_code.get("21001", {})) <= 1:
            self.skipTest("이 확정본에는 210-01 이 갈리는 기록이 없다")
        groups = {group.code5: group for group in build_groups(self.session)}
        target = groups.get("210-01")
        self.assertIsNotNone(target)
        self.assertTrue(target.split, "항목명에 따라 갈리는 목코드는 펼쳐야 한다")
        self.assertEqual(target.code, "", "갈리는 목코드는 단일 값을 내놓으면 안 된다")
        self.assertTrue(any(sub.settled for sub in target.subgroups))
        self.assertTrue(any(not sub.settled for sub in target.subgroups))

    def test_override_by_code_settles_every_row_in_the_group(self):
        from bimok_groups import build_groups
        before = len(self.session.unsettled)
        target = next(group for group in build_groups(self.session) if not group.settled and not group.split)
        self.session.overrides_by_code[target.code5] = target.digits + "99"
        after = len(self.session.unsettled)
        self.assertEqual(before - after, target.rows - target.settled_rows)
        del self.session.overrides_by_code[target.code5]

    def test_name_override_beats_code_override(self):
        project = self.session.projects[0]
        index = next(i for i, item in enumerate(project.items) if item.is_leaf)
        item = project.items[index]
        self.session.overrides_by_code[item.bimok_code5] = "9999999"
        self.session.overrides_by_name[(item.bimok_code5, normalise(item.name))] = "1111111"
        self.assertEqual(self.session.code_for(project, index)[0], "1111111")
        self.session.overrides_by_code.clear()
        self.session.overrides_by_name.clear()

    def test_remembered_codes_survive_a_round_trip(self):
        self.session.overrides_by_code["210-07"] = "2100702"
        self.session.overrides_by_name[("210-01", "운영용품")] = "2100143"
        stored = self.session.remembered()
        fresh = Session()
        fresh.apply_remembered(stored)
        self.assertEqual(fresh.overrides_by_code["210-07"], "2100702")
        self.assertEqual(fresh.overrides_by_name[("210-01", "운영용품")], "2100143")
        self.session.overrides_by_code.clear()
        self.session.overrides_by_name.clear()

    def test_issue_export_writes_a_workbook(self):
        from report_export import export_issues
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "issues.xlsx"
            export_issues(str(target), self.session, {"목표1": self.session.plan_report})
            self.assertTrue(target.exists())
            from openpyxl import load_workbook
            book = load_workbook(target)
            self.assertIn("지적 목록", book.sheetnames)
            self.assertIn("비목 미확정", book.sheetnames)

    def test_progress_reports_every_stage(self):
        seen = []
        session = Session(plan_path=str(PLAN), ubis_path="", last_year_path=str(LAST_YEAR))
        session.load(progress=lambda message, ratio: seen.append((message, ratio)))
        self.assertGreaterEqual(len(seen), 5)
        self.assertEqual(seen[-1][1], 1.0)


@unittest.skipUnless(CLASSES.exists() and PLAN.exists(), "분류표가 없으면 건너뛴다")
class ClassificationTests(unittest.TestCase):
    """별표 3은 사업 분류다. 원가통계비목이 아니다."""

    @classmethod
    def setUpClass(cls):
        from program_classes import load_classification
        cls.classes = load_classification(str(CLASSES))
        cls.projects = parse_plan(str(PLAN))

    def test_five_levels_are_read(self):
        counts = self.classes.counts
        self.assertGreaterEqual(counts.get("정책사업", 0), 10)
        self.assertGreaterEqual(counts.get("단위사업", 0), 40)
        self.assertGreaterEqual(counts.get("세부사업", 0), 100)

    def test_every_project_matches_the_official_names(self):
        from program_classes import POLICY, PROGRAM, UNIT
        for project in self.projects:
            for level, value in ((POLICY, project.policy), (UNIT, project.unit), (PROGRAM, project.program)):
                self.assertIsNotNone(self.classes.find(level, value), f"{project.name} · {value}")

    def test_unknown_name_is_reported(self):
        from program_classes import check_projects
        from plan_parser import PlanProject
        broken = PlanProject(name="시험", policy="없는정책사업", unit="없는단위사업", program="없는세부사업")
        report = check_projects([broken], self.classes)
        self.assertEqual(len(report.errors), 3)

    def test_classification_carries_no_cost_codes(self):
        """이 문서로는 비목을 정할 수 없다. 근거를 테스트로 고정한다."""
        text = " ".join(entry.note for entry in self.classes.entries)
        for word in ("일반수용비", "강사수당", "시설임차료", "원가통계"):
            self.assertNotIn(word, text)

    def test_one_project_uses_many_cost_codes(self):
        """사업 제목으로 비목을 정할 수 없는 이유. 한 사업이 여러 목코드를 쓴다."""
        spread = [len({item.bimok_code5 for item in project.items if item.is_leaf})
                  for project in self.projects if project.items]
        self.assertGreater(max(spread), 5, "한 사업이 목코드 여러 종을 쓴다")
        self.assertGreater(sum(1 for one in spread if one > 1), len(spread) // 2,
                           "절반 넘는 사업이 목코드를 2종 이상 쓴다")


@unittest.skipUnless(LAST_YEAR.exists(), "작년 확정본이 없으면 건너뛴다")
class ParentContextTests(unittest.TestCase):
    """목코드와 항목명만으로는 갈리는 것이 부모 항목으로 갈린다.

        강사수당 > 기본  →  2100605
        심사수당 > 기본  →  2100607

    부모를 보지 않으면 '기본' 이 두 코드 사이에서 미확정으로 남거나,
    더 나쁘게는 한쪽으로 잘못 확정된다.
    """

    @classmethod
    def setUpClass(cls):
        cls.resolver = BimokResolver()
        cls.resolver.learn(str(LAST_YEAR))

    def test_parent_key_is_learned(self):
        self.assertTrue(self.resolver.by_parent, "부모 항목 열쇠가 비어 있으면 갈림을 풀 수 없다")
        keys = {key for key in self.resolver.by_parent if key[0] == "21006"}
        self.assertTrue(keys)

    def test_parent_settles_every_ambiguous_item_name(self):
        """항목명만으로 갈리는 조합이 부모까지 보면 하나로 좁혀지는지.

        있는 자료로 성립하는 불변식만 검사한다. 특정 코드값을 박지 않는다.
        """
        ambiguous = [(code5, name) for (code5, name), counts
                     in self.resolver.by_code_and_name.items() if len(counts) > 1]
        for code5, name in ambiguous:
            parents = [key[1] for key in self.resolver.by_parent
                       if key[0] == code5 and key[2] == name]
            self.assertTrue(parents, f"{code5}·{name} 에 부모 기록이 없다")
            for parent in parents:
                counts = self.resolver.by_parent[(code5, parent, name)]
                self.assertEqual(len(counts), 1,
                                 f"{code5} · {parent} > {name} 이 부모까지 봐도 갈린다")

    def test_parent_context_wins_over_bare_item_name(self):
        """부모가 주어지면 항목명만 보는 답보다 먼저 쓰인다."""
        for (code5, parent, name), counts in self.resolver.by_parent.items():
            if not parent or len(counts) != 1:
                continue
            expected = next(iter(counts))
            found = self.resolver.resolve(f"{code5[:3]}-{code5[3:]}", name, parent=parent)
            self.assertEqual(found.code, expected, f"{code5} · {parent} > {name}")
            self.assertIn("부모 항목", found.source)
            return
        self.skipTest("부모 기록이 없는 표본")

    def test_bimok_name_is_used_as_a_fallback_context(self):
        """설명서에는 부모 대신 비목명이 붙어 있는 줄도 있다."""
        found = self.resolver.resolve("210-06", "기본", parent="", bimok_name="강사수당")
        self.assertTrue(found.settled or found.candidates)

    def test_learning_more_files_never_loses_a_key(self):
        before = len(self.resolver.by_parent)
        self.resolver.learn(str(LAST_YEAR))
        self.assertGreaterEqual(len(self.resolver.by_parent), before)
        self.assertEqual(len(self.resolver.sources), 2, "같은 파일을 두 번 배워도 기록은 쌓인다")


@unittest.skipUnless(PLAN.exists(), "설명서 표본이 없으면 건너뛴다")
class LocatingTests(unittest.TestCase):
    """지적을 설명서에서 찾아갈 수 있어야 한다."""

    @classmethod
    def setUpClass(cls):
        cls.projects = parse_plan(str(PLAN))

    def test_every_project_carries_its_document_number(self):
        missing = [one.name for one in self.projects if not one.number]
        self.assertEqual(missing, [], "설명서에 적힌 사업 번호를 모두 읽어야 한다")

    def test_label_shows_number_then_name(self):
        target = next(one for one in self.projects if "영유아교육내실화" in one.name)
        self.assertTrue(target.label.startswith(f"{target.number}. "))
        self.assertIn(target.name, target.label)

    def test_numbers_are_unique(self):
        numbers = [one.number for one in self.projects]
        self.assertEqual(len(numbers), len(set(numbers)))

    def test_find_text_is_literal_document_text(self):
        """쪽 번호가 없으므로 문서에 그대로 있는 글자를 준다."""
        target = self.projects[0]
        self.assertEqual(target.find_text(), target.heading)
        leaf = next(item for item in target.items if item.is_leaf)
        self.assertEqual(target.find_text(leaf), leaf.formula)



class BimokScreenMatchesWarningTests(unittest.TestCase):
    """비목 화면의 묶음과 '미확정 N행' 경고는 **같은 행**을 세야 한다. (2.13.0)

    2027 설명서는 산출근거를 산출내역 표에서 읽어 줄마다 목코드가 없다. 예전에는
    그런 줄을 묶음에서 버려서, 경고는 '미확정 631행'인데 비목 화면에는 확정할 줄이
    하나도 없었다(대신 입력하지도 않는 재배정 줄 묶음 하나만 '미확정'으로 떠 있었다).
    """

    def _session(self):
        project = PlanProject(name="표본사업", request=300, year_request=300, had_own_items=True)
        project.items = [PlanItem(depth=1, name="운영", amount=300, row=1),
                         PlanItem(depth=2, name="운영용품", amount=100, formula="100,000원×1회", row=2),
                         PlanItem(depth=2, name="처음보는항목", amount=100, formula="100,000원×1회", row=3),
                         PlanItem(depth=2, name="자료제작", amount=100, formula="100,000원×1회", row=4,
                                  bimok_name="일반수용비", bimok_code5="210-01")]
        handed = PlanProject(name="넘기는사업(총액배분사업비)", request=50, year_request=50, had_own_items=True)
        handed.items = [PlanItem(depth=1, name="학교운영비", amount=50, formula="50,000원×1교",
                                 bimok_name="총액배분사업비", bimok_code5="620-02")]
        session = Session()
        session.projects = [project, handed]
        session.resolver = BimokResolver()
        session.bimok_report, session.resolutions = check_bimok(session.projects, session.resolver)
        return session

    def test_every_unsettled_row_is_on_the_screen(self):
        from bimok_groups import build_groups
        session = self._session()
        groups = build_groups(session)
        shown = sum(group.rows - group.settled_rows for group in groups)
        self.assertEqual(shown, len(session.unsettled))
        self.assertEqual(len(session.unsettled), 3)

    def test_handover_rows_are_not_offered(self):
        from bimok_groups import build_groups
        codes = {group.code5 for group in build_groups(self._session())}
        self.assertNotIn("620-02", codes, "입력하지 않는 재배정 줄을 확정하라고 내밀면 안 된다")

    def test_code_less_rows_form_one_group_settled_by_name(self):
        from bimok_groups import build_groups
        session = self._session()
        group = next(one for one in build_groups(session) if one.no_code)
        self.assertEqual(group.rows, 2)
        self.assertEqual(group.bimok_label, "목코드 없음")
        self.assertEqual(group.digits, "", "앞 5자리를 모르면 7자리 전부를 받아야 한다")
        sub = next(one for one in group.subgroups if any("운영용품" in key for key in one.names))
        self.assertEqual(sub.names, ["운영 > 운영용품"], "목코드 없는 줄은 부모까지 열쇠로 쓴다")
        for key in sub.names:
            session.overrides_by_name[("", key)] = "2100143"
        self.assertEqual(session.code_for(session.projects[0], 1)[0], "2100143")
        self.assertEqual(len(session.unsettled), 2)
        stored = session.remembered()
        fresh = Session()
        fresh.apply_remembered(stored)
        self.assertEqual(fresh.overrides_by_name[("", "운영 > 운영용품")], "2100143")

    def test_same_name_under_other_parent_is_not_overwritten(self):
        """'강사수당 > 기본' 을 정했다고 '심사수당 > 기본' 까지 같은 코드를 받으면 안 된다."""
        from bimok_groups import override_key
        project = PlanProject(name="표본", request=200, year_request=200, had_own_items=True)
        project.items = [PlanItem(depth=1, name="강사수당", amount=100, row=1),
                         PlanItem(depth=2, name="기본", amount=100, formula="100,000원×1회", row=2),
                         PlanItem(depth=1, name="심사수당", amount=100, row=3),
                         PlanItem(depth=2, name="기본", amount=100, formula="100,000원×1회", row=4)]
        session = Session()
        session.projects = [project]
        session.resolver = BimokResolver()
        session.bimok_report, session.resolutions = check_bimok(session.projects, session.resolver)
        session.overrides_by_name[("", override_key(project.items, 1))] = "2100605"
        self.assertEqual(session.code_for(project, 1)[0], "2100605")
        self.assertEqual(session.code_for(project, 3)[0], "", "다른 부모의 '기본'은 그대로 비어 있어야 한다")

    def test_warning_counts_decisions_not_just_code_groups(self):
        session = self._session()
        text = session.warnings()[0]
        self.assertIn("3행", text)
        self.assertIn("목코드 없는 항목명 2종", text)
        self.assertIn("목코드 1종", text)


class DownloadDiffTests(unittest.TestCase):
    """입력 후 대조(목표 2B)에서 비목을 K-에듀파인에서 직접 고른 줄을 오류로 몰면 안 된다."""

    def _write(self, folder, name, code, amount):
        from edufine_form import OutputRow, write_form
        template = Path(folder) / "template.xlsx"
        if not template.exists():
            from openpyxl import Workbook
            book = Workbook()
            sheet = book.active
            heads = ["레벨", "순번", "* 사업항목", "", "* 원가통계비목", "", "* 상세내역", "",
                     "2026년도 요구", ""]
            subs = ["", "", "", "", "", "", "", "", "* 산출기초", "* 금액②"]
            sheet.append(heads)
            sheet.append(subs)
            sheet.append(["", "", "총계"])
            sheet.append(["1", 1, "x"])
            book.save(template)
        target = Path(folder) / name
        rows = [OutputRow(level=1, name="1. 사업", amount=amount),
                OutputRow(level=2, name="가. 운영용품", code=code, detail="본청",
                          basis="100,000원×1회=", amount=amount)]
        write_form(str(template), str(target), [("사업", rows)])
        return str(target)

    def test_blank_code_filled_in_later_is_not_an_error(self):
        from crosscheck import diff_forms
        with tempfile.TemporaryDirectory() as folder:
            made = self._write(folder, "made.xlsx", "", 100)
            got = self._write(folder, "got.xlsx", "2100143", 100)
            report = diff_forms(made, got)
            self.assertEqual(report.issues, [], [one.message for one in report.issues])

    def test_one_thousand_won_typo_is_still_caught(self):
        from crosscheck import diff_forms
        with tempfile.TemporaryDirectory() as folder:
            made = self._write(folder, "made.xlsx", "2100143", 100)
            got = self._write(folder, "got.xlsx", "2100143", 101)
            self.assertTrue(diff_forms(made, got).errors, "1천원 오타는 잡아야 한다")


class SupplementFormTests(unittest.TestCase):
    """추경 설명서의 산출내역 표는 증감분만 적는다. 산출근거로 읽으면 안 된다. (2.13.0)"""

    def test_supplement_table_is_recognised(self):
        from plan_parser import _is_supplement_table
        supplement = [["(단위 : 천원)"],
                      ["사업명", "추가경정예 산 안(A)", "기  정예산액(B)(전년도예산액)", "비 교 증 감 내 역(A-B)"],
                      ["산출내역", "요구액"]]
        main = [["(단위 : 천원)"], ["2027년도 예산(안)", "2026년", "비교 증감"],
                ["산 출 내 역", "요구액(A)", "본예산액(B)", "최종예산액(C)", "본예산대비(A-B)"]]
        self.assertTrue(_is_supplement_table(supplement))
        self.assertFalse(_is_supplement_table(main))

    def test_supplement_plan_is_checked_not_converted(self):
        """2.13.0 까지는 추경 설명서를 받으면 이유를 말하고 멈췄다. 2.14.0 부터는 점검한다.

        입력본은 여전히 만들지 않는다. 산출식이 증감분만 있어 K-에듀파인 입력의 원천이
        될 수 없다는 사실은 그대로다. 2번 K-에듀파인 파일 없이도 불러와야 한다.
        """
        sample = find_sample("2026사업별 설명서_2026.9.21_10_16_13.hwpx")
        if not sample.exists():
            self.skipTest("추경 표본이 없다")
        session = Session(plan_path=str(sample))
        session.load()
        self.assertTrue(session.is_supplement)
        self.assertTrue(session.projects)
        self.assertEqual(session.unsettled, [], "추경에는 비목 확정할 줄이 없다")
        with tempfile.TemporaryDirectory() as folder:
            with self.assertRaises(ValueError) as caught:
                session.save(str(Path(folder) / "out.xlsx"), force=True)
        self.assertIn("추경", str(caught.exception))
        with self.assertRaises(ValueError):
            session.verify_downloaded("아무거나.xlsx")


class YearCheckTests(unittest.TestCase):
    def test_file_years_are_compared(self):
        from ui_common import year_of
        self.assertEqual(year_of("2027검토조서(서식1)_2026.9.21_9_15_28.xlsx"), 2027)
        self.assertIsNone(year_of("(본예산세출요구)검토조서(서식1)_2026.9.17_17_24_48.xlsx"),
                          "출력 날짜는 회계연도가 아니다")
        try:
            from native_ui import year_warnings
        except Exception:  # noqa: BLE001 - tkinter 없는 곳
            self.skipTest("tkinter 없음")
        found = year_warnings({"plan": "2027사업별 설명서.hwpx",
                               "ubis": "2026검토조서(서식1).xlsx",
                               "last_year": "2027(k에듀파인)세출예산요구내역.xlsx"})
        self.assertIn("ubis", found)
        self.assertIn("last_year", found)
        self.assertEqual(year_warnings({"plan": "2027사업별 설명서.hwpx",
                                        "ubis": "2027검토조서.xlsx",
                                        "last_year": "2026(k에듀파인).xlsx"}), {})



class UbisNameTests(unittest.TestCase):
    """'UBIS 검토조서 열기'가 복사하는 찾을 글자. 엑셀 찾기는 글자가 그대로여야 걸린다."""

    def test_exact_name_is_kept(self):
        from crosscheck import ubis_name
        self.assertEqual(ubis_name({"교육공무직원인건비관리": 1.0}, "교육공무직원인건비관리"), "교육공무직원인건비관리")

    def test_spacing_and_middle_dot_follow_ubis(self):
        from crosscheck import ubis_name
        ubis = {"특수학교 방과후\u2024돌봄 지원 강화": 1.0, "초등교육과정 전문성 신장": 2.0}
        self.assertEqual(ubis_name(ubis, "초등교육과정전문성신장"), "초등교육과정 전문성 신장")
        self.assertEqual(ubis_name(ubis, "특수학교 방과후·돌봄 지원 강화"), "특수학교 방과후\u2024돌봄 지원 강화")

    def test_unknown_or_ambiguous_falls_back_to_plan_name(self):
        from crosscheck import ubis_name
        self.assertEqual(ubis_name({}, "없는 사업"), "없는 사업")
        self.assertEqual(ubis_name({"가 나": 1.0, "가·나": 2.0}, "가나"), "가나")

if __name__ == "__main__":
    unittest.main()


class HierarchyTests(unittest.TestCase):
    """위계를 잘못 읽었을 때 '오류'라고 단정하지 않는지.

    실제로 있었던 일: 9,000 + 3,000 = 12,000 으로 멀쩡히 맞는 사업이 화면에서
    붉게 칠해져 나왔다. 덧셈이 틀린 게 아니라 (1) 총사업비와 올해 요구액을
    같은 값으로 봤고 (2) 같은 이름의 형제 줄을 전부 칠했기 때문이다.
    """

    @staticmethod
    def _project(rows, request=None, year_request=None, outline_total=None):
        project = PlanProject(name="표본사업", request=request,
                              year_request=year_request, outline_total=outline_total)
        for depth, name, amount, formula in rows:
            project.items.append(PlanItem(depth=depth, name=name, amount=amount, formula=formula))
        return project

    def test_direct_children_disagree_but_leaves_agree_is_a_review(self):
        """단계를 한 칸 잘못 읽어도 잎의 합이 맞으면 오류가 아니다."""
        project = self._project([
            (1, "묶음", 12000, ""),
            (3, "공립", 9000, "3,000,000원×3교"),     # 2단이어야 할 줄이 3단으로 읽혔다
            (3, "사립", 3000, "3,000,000원×1교"),
        ], year_request=12000)
        report = check_plan([project])
        self.assertEqual(report.errors, [], [issue.message for issue in report.errors])
        self.assertTrue(any("단계" in issue.message for issue in report.dropped),
                        "오류가 아니라고 판단해 뺀 기록이 남아야 한다")

    def test_real_shortfall_is_still_an_error(self):
        """느슨해지기만 하면 안 된다. 어느 방식으로 더해도 안 맞으면 오류다."""
        project = self._project([
            (1, "묶음", 12000, ""),
            (2, "공립", 9000, "3,000,000원×3교"),
            (2, "사립", 1000, "1,000,000원×1교"),
        ], year_request=10000)
        self.assertTrue(check_plan([project]).errors)

    def test_one_broken_place_is_reported_once(self):
        """아래가 어긋나면 위도 따라 어긋난다. 한 군데 잘못이 열 건이 되면 안 된다."""
        project = self._project([
            (1, "위", 10000, ""),
            (2, "가운데", 10000, ""),
            (3, "아래", 9000, "9,000,000원×1식"),
        ], year_request=10000)
        errors = check_plan([project]).errors
        self.assertEqual(len(errors), 1, [issue.message for issue in errors])
        self.assertEqual(errors[0].item, "가운데")

    def test_the_request_table_is_the_basis_not_the_headline(self):
        """맞춰야 하는 값은 '본예산안(A)' 표의 요구액이다. 머리글 총사업비가 아니다.

        총사업비는 올해 그 사업의 사업비라 요구액과 같아야 하지만, 손으로 적는 칸이라
        어긋나는 일이 있다. 그때 기준은 요구액 쪽이고(산출근거 합계가 그쪽을 편든다),
        머리글은 '고쳐 달라'고 말할 대상이다. 그것 때문에 입력본 생성을 막지는 않는다.
        """
        project = self._project([
            (1, "공립", 1040000, "20,000,000원×52개학급"),
        ], request=3240000, year_request=1040000)
        report = check_plan([project])
        self.assertEqual(project.base, 1040000)
        self.assertEqual(len(report.errors), 1, [issue.message for issue in report.errors])
        self.assertIn("머리글 총사업비 한 칸만 다릅니다", report.errors[0].message)
        self.assertEqual(report.blocking_errors, [])

    def test_no_witness_agrees_is_an_error(self):
        project = self._project([
            (1, "공립", 500, "500,000원×1식"),
        ], request=3240000, year_request=1040000)
        self.assertTrue(check_plan([project]).errors)

    def test_dangling_row_that_would_vanish_from_the_total(self):
        """산출식도 자식도 없이 금액만 있는 줄은 합계에서 조용히 빠진다."""
        project = self._project([
            (1, "묶음", 12000, ""),
            (2, "공립", 9000, "3,000,000원×3교"),
            (2, "빠진 줄", 3000, ""),
        ], year_request=12000)
        errors = check_plan([project]).errors
        self.assertTrue(any("합계에서 빠지고" in issue.message for issue in errors),
                        "산출식 없는 금액 줄은 입력본에 산출기초가 비므로 오류다 (2.15.0)")

    def test_issue_points_at_one_row_when_names_repeat(self):
        """'공립'이 두 번 나오는 사업에서 이름으로 짚으면 멀쩡한 줄까지 칠해진다."""
        items = [PlanItem(depth=1, name="이음교육", amount=12000),
                 PlanItem(depth=2, name="공립", amount=9000, formula="3,000,000원×3교"),
                 PlanItem(depth=2, name="사립", amount=3000, formula="3,000,000원×1교"),
                 PlanItem(depth=1, name="급식비", amount=33400),
                 PlanItem(depth=2, name="공립", amount=33400, formula="33,400,000원×1교")]
        self.assertEqual(issue_row(items, Issue("목표1", "오류", "표본", "공립", index=4)), 4)
        # 자리를 모르고 이름이 겹치면 아무 줄도 짚지 않는다.
        self.assertEqual(issue_row(items, Issue("목표1", "오류", "표본", "공립")), -1)
        # 이름이 하나뿐이면 옛 방식대로 찾아도 안전하다.
        self.assertEqual(issue_row(items, Issue("목표1", "오류", "표본", "사립")), 2)

    def test_every_amount_issue_carries_its_row(self):
        project = self._project([
            (1, "묶음", 12000, ""),
            (2, "공립", 9000, "3,000,000원×1교"),     # 계산값 3,000 ≠ 금액 9,000
            (2, "사립", 3000, "3,000,000원×1교"),
        ], year_request=12000)
        for issue in check_plan([project]).issues:
            if issue.item:
                self.assertIsNotNone(issue.index, issue.message)


class NameMatchingTests(unittest.TestCase):
    """설명서와 UBIS는 띄어쓰기가 자주 어긋난다. 같은 사업을 '없다'고 하면 안 된다."""

    def test_spacing_only_difference_is_the_same_project(self):
        project = PlanProject(name="초등교육과정전문성신장", year_request=1000)
        project.items.append(PlanItem(depth=1, name="운영", amount=1000, formula="1,000,000원×1식"))
        report = compare_totals("목표2A", [project], {"초등교육과정 전문성 신장": 1000}, "", "")
        self.assertEqual(report.errors, [])
        self.assertFalse([issue for issue in report.issues if "찾지 못했" in issue.message])

    def test_two_candidates_are_not_guessed(self):
        project = PlanProject(name="운영지원", year_request=1000)
        report = compare_totals("목표2A", [project], {"운영 지원": 1000, "운영지 원": 2000}, "", "")
        # 이름으로는 둘 중 하나를 못 고른다. 금액이 같은 쪽이 하나뿐이면 그 사업이다 (2.15.0).
        self.assertEqual(report.errors, [], [issue.message for issue in report.errors])

    def test_no_name_and_no_amount_match_is_an_error(self):
        """이름도 금액도 맞는 UBIS 사업이 없으면 UBIS 에 요구가 빠진 것이다 (2.15.0)."""
        project = PlanProject(name="운영지원", year_request=1000)
        report = compare_totals("목표2A", [project], {"운영 지원": 3000, "운영지 원": 2000}, "", "")
        self.assertEqual(len(report.errors), 1)
        self.assertIn("UBIS", report.errors[0].message)


@unittest.skipUnless(PLAN.exists(), "표본이 없으면 건너뛴다")
class SampleAmountTests(unittest.TestCase):
    def test_sample_has_no_blocking_amount_errors(self):
        """표본 69개 사업에서 입력본을 막는 목표1 오류는 0건이어야 한다.

        산출식·합계·두 서식 대조에서 나오는 오류는 없다. 남는 것은 머리글 총사업비
        한 칸이 다른 1건인데, 그 칸은 입력본에 실리지 않으므로 생성을 막지 않는다.
        """
        report = check_plan(parse_plan(str(PLAN)))
        self.assertEqual(report.blocking_errors, [],
                         [f"{i.project} {i.item} {i.message}" for i in report.blocking_errors])

    def test_every_project_has_this_years_request(self):
        for project in parse_plan(str(PLAN)):
            self.assertIsNotNone(project.year_request, project.name)


class SplitDescriptionTests(unittest.TestCase):
    """설명서가 한 사업을 여러 쪽으로 나눠 적을 때.

    2027 설명서에서 실제로 난 일이다.
      67번 장애영유아교육지원 — 총괄 쪽(45,400) + 세부 쪽 둘(12,000·33,400).
           세 쪽을 각각 사업으로 세면 12,000과 33,400이 UBIS 45,400과 따로 비교돼
           오류 2건이 난다. 둘 다 오류가 아니다.
      47번 유치원 입학관리시스템 연수 — 이름이 다른 별개 사업(34,500)이 번호 없이
           붙어 있고, UBIS는 둘을 한 사업(50,180)으로 묶어 둔다.
    """

    @staticmethod
    def _project(name, year_request, parts=(), items=()):
        project = PlanProject(name=name, year_request=year_request)
        for part_name, amount in parts:
            project.parts.append(PlanPart(name=part_name, year_request=amount))
        for depth, item_name, amount, formula in items:
            project.items.append(PlanItem(depth=depth, name=item_name, amount=amount, formula=formula))
        return project

    def test_overview_page_total_matches_ubis(self):
        project = self._project("장애영유아교육지원(신규, 총액배분사업비)", 45400, parts=(
            ("특수학교 유-보-초 이음교육 운영(총액배분사업비)", 12000),
            ("특수학교(유치원) 방학 중 방과후 과정 급식비 지원 (총액배분사업비)", 33400)), items=(
            (1, "특수학교 유치원 유-보-초 이음교육 운영", 12000, ""),
            (2, "공립", 9000, "3,000,000원×3교"),
            (2, "사립", 3000, "3,000,000원×1교"),
            (1, "특수학교(유치원) 방학 중 방과후 과정 급식비 지원", 33400, ""),
            (2, "공립", 33400, "33,400,000원×1교")))
        self.assertEqual(check_plan([project]).issues, [])
        report = compare_totals("목표2A", [project],
                                {"장애영유아교육지원(신규, 총액배분사업비)": 45400}, "", "")
        self.assertEqual(report.errors, [], [issue.message for issue in report.errors])

    def test_no_synthetic_amount_is_invented(self):
        """여러 쪽의 요구액을 더해 만든 숫자를 '설명서 금액'이라 부르지 않는다.

        47번은 설명서 15,680 인데, 번호 없이 붙은 **다른 사업** 34,500 을 더해
        50,180 을 만들어 화면에 '한글 설명서 50,180천원'으로 띄웠다. 설명서
        어디에도 없는 숫자였다.
        """
        project = self._project("유치원 입학관리시스템 연수(재원배분)", 15680)
        project.split_description = True
        values = [value for _label, value in project.amount_candidates]
        self.assertEqual(values, [15680])
        report = compare_totals("목표2A", [project],
                                {"유치원 입학관리시스템 연수(재원배분)": 50180}, "", "")
        self.assertNotIn(50180, [issue.left for issue in report.issues])

    def test_split_description_is_a_review_not_an_error(self):
        """설명서가 여러 쪽으로 나눠 적었으면 UBIS 와 묶는 기준이 다를 수 있다."""
        project = self._project("유치원 입학관리시스템 연수(재원배분)", 15680)
        project.split_description = True
        report = compare_totals("목표2A", [project],
                                {"유치원 입학관리시스템 연수(재원배분)": 50180}, "", "")
        self.assertEqual(report.errors, [], [issue.message for issue in report.errors])
        self.assertTrue(report.dropped)

    def test_a_plain_project_still_errors(self):
        """나눠 적은 흔적이 없으면 차액은 그대로 오류다."""
        project = self._project("유치원 입학관리시스템 연수(재원배분)", 15680)
        report = compare_totals("목표2A", [project],
                                {"유치원 입학관리시스템 연수(재원배분)": 50180}, "", "")
        self.assertEqual(len(report.errors), 1)

    def test_a_real_gap_is_still_an_error(self):
        """느슨해지기만 하면 안 된다. 어느 후보와도 안 맞으면 오류다."""
        project = self._project("학력평가지원관리", 530584)
        report = compare_totals("목표2A", [project], {"학력평가지원관리": 490584}, "", "")
        self.assertEqual(len(report.errors), 1)
        message = report.errors[0].message
        self.assertIn("한글 설명서", message)
        self.assertIn("엑셀", message)          # 어느 파일이 얼마인지 말해야 한다

    def test_total_row_of_a_basis_table_is_not_an_item(self):
        """산출근거 표 맨 끝 합계 줄은 앞 칸이 비고 금액만 남는다.

        그 금액이 항목명이 되면 '33,400'이라는 빈 자식이 생기고, 부모의 하위 합계가
        0이 되어 멀쩡한 사업이 금액 오류로 지적된다. (2027 설명서 67번)
        """
        rows = [["1.", "특수학교(유치원) 방학 중 방과후 과정 급식비 지원", "", "33,400", "천원"],
                ["", "①", "공립", "총액배분사업비(620-02)", "33,400,000원×1교", "=", "33,400", "천원"],
                ["", "", "", "", "", "", "33,400", "천원"]]
        names = [item.name for item in _parse_basis_rows(rows, {})]
        self.assertNotIn("33,400", names)


@unittest.skipUnless(PLAN.exists(), "표본이 없으면 건너뛴다")
class SampleStructureTests(unittest.TestCase):
    def test_sample_still_reads_every_project_and_leaf(self):
        """세부 쪽 인식을 넣어도 한 덩어리로 적힌 설명서가 줄어들면 안 된다."""
        projects = parse_plan(str(PLAN))
        self.assertEqual(len(projects), 69)
        self.assertEqual(sum(1 for p in projects for i in p.items if i.is_leaf), 1118)
        names = [p.name for p in projects]
        self.assertEqual(len(names), len(set(names)), "같은 이름의 사업이 둘 생기면 UBIS 대조가 갈린다")


class HandoverTests(unittest.TestCase):
    """이 과에서 K-에듀파인에 입력하지 않는 사업.

    총액배분사업비(620-02)는 예산과가 학교로 바로 재배정하고, 재원배분 사업은 다른
    과가 나눠 준다. 확정할 수 없는 비목을 '미확정'으로 세면 담당자는 끝낼 수 없는
    일을 붙들게 되고, 입력본에 섞어 두면 그대로 입력해 버린다.
    """

    @staticmethod
    def _project(name, rows):
        project = PlanProject(name=name, year_request=sum(a for *_x, a, _f in rows))
        for depth, item_name, amount, formula in rows:
            project.items.append(PlanItem(depth=depth, name=item_name, amount=amount,
                                          formula=formula, bimok_code5="620-02" if formula else ""))
        return project

    def test_total_allocation_project_is_not_entered_here(self):
        project = self._project("장애영유아교육지원(신규, 총액배분사업비)",
                                [(1, "공립", 9000, "3,000,000원×3교")])
        self.assertEqual(project.handover, "총액배분 재배정")

    def test_name_is_not_the_test_the_cost_code_is(self):
        """표본의 '특수학교 교육과정 운영지원'은 이름에 총액배분이 없는데 잎이 전부 620-02다."""
        project = self._project("특수학교 교육과정 운영지원",
                                [(1, "공립", 5000, "5,000,000원×1교")])
        self.assertEqual(project.handover, "총액배분 재배정")

    def test_shared_allocation_project_is_not_entered_here(self):
        self.assertEqual(PlanProject(name="기초학력지원센터 운영(재원배분)").handover, "재원배분")

    def test_ordinary_project_is_entered(self):
        project = PlanProject(name="영유아교육내실화 지원", year_request=4000)
        project.items.append(PlanItem(depth=1, name="운영용품", amount=4000,
                                      formula="50,000원×20종×4회", bimok_code5="210-01"))
        self.assertEqual(project.handover, "")

    def test_handover_rows_are_not_counted_as_unsettled(self):
        project = self._project("특수교육교재교구지원(총액배분사업비)",
                                [(1, "공립", 5000, "5,000,000원×1교")])
        report, settled = check_bimok([project], BimokResolver())
        self.assertEqual(settled, {})
        self.assertEqual(report.issues, [], [issue.message for issue in report.issues])


class OutlineTableTests(unittest.TestCase):
    """'산출내역' 표를 산출근거로 읽는다.

    2027 설명서는 '요구내용 및 산출근거' 표(비목이 붙은 서식)가 아직 비어 있고
    '산출내역' 표만 채워져 있었다. 비목이 없다고 통째로 버리면 금액 검산이
    아무 일도 하지 못한다. 항목·계층·산출식·금액은 다 들어 있다.
    """

    ROWS = [
        ["(단위 : 천원)"],
        ["2027년도 예산(안)", "2026년", "비교 증감"],
        ["산 출 내 역", "요구액(A)", "본예산액(B)", "최종예산액(C)"],
        ["1. 교육과정 실천중심 현장 지원", "30,920", "30,920", "36,920"],
        ["① 초등학교 신설교과 현장 이해", "7,000천원", "7,000", "7,000"],
        ["② 2022개정교육과정 핵심교원 연수 (본청)", "3,060", "3,060", "3,060"],
        ["․ 운영용품(50천원×5종×2회)", "500천원"],
        ["․ 강사수당(일반)", "1,640천원"],
        ["\U000f02fb초과(90천원×2명×1시간×2회)", "360천원"],
        ["\U000f02fb기본(160천원×2명×2시간×2회)", "1,280천원"],
        ["․ 강사실비경비(80천원×2명×2회)", "320천원"],
        ["․ 협의회(15천원×20명×2회)", "600천원"],
        ["③ 2022개정교육과정 이해 교사연수 (교육지원청)", "20,860천원", "20,860"],
    ]

    def setUp(self):
        self.bullets = _learn_outline_bullets([self.ROWS])
        self.items = _parse_outline_rows(self.ROWS, self.bullets)

    def test_header_rows_are_not_items(self):
        self.assertEqual(self.items[0].name, "교육과정 실천중심 현장 지원")
        self.assertEqual(len(self.items), 10)

    def test_name_and_formula_are_split(self):
        found = next(item for item in self.items if item.name == "운영용품")
        self.assertEqual(found.formula, "50천원×5종×2회")
        self.assertEqual(found.amount, 500)

    def test_parenthesis_without_multiplication_stays_in_the_name(self):
        """'(본청)' '(일반)' 은 산출식이 아니라 이름의 일부다."""
        names = [item.name for item in self.items]
        self.assertIn("2022개정교육과정 핵심교원 연수 (본청)", names)
        self.assertIn("강사수당(일반)", names)

    def test_thousand_won_cell_wins_over_the_year_columns(self):
        """자기 금액은 '7,000천원', 그 뒤는 연도별 예산 열이다."""
        found = next(item for item in self.items if "신설교과" in item.name)
        self.assertEqual(found.amount, 7000)

    def test_a_row_without_a_formula_can_still_be_a_leaf(self):
        found = next(item for item in self.items if "신설교과" in item.name)
        self.assertEqual(found.formula, "")
        self.assertTrue(found.is_leaf)
        parent = next(item for item in self.items if item.name == "강사수당(일반)")
        self.assertFalse(parent.is_leaf)

    def test_children_close_on_their_parent(self):
        """② 3,060 = 운영용품 500 + (초과 360 + 기본 1,280) + 강사실비 320 + 협의회 600."""
        parent = next(i for i, item in enumerate(self.items) if "핵심교원" in item.name)
        self.assertEqual(leaf_total(self.items, subtree_of(self.items, parent)),
                         self.items[parent].amount)

    def test_leaf_total_closes_on_the_request(self):
        """맨 아래 산출근거를 모두 더하면 1단 금액이자 올해 요구액이 된다."""
        project = PlanProject(name="초등교육과정 운영", year_request=30920)
        project.items.extend(self.items)
        self.assertEqual(project.item_total, 30920)
        self.assertEqual(check_plan([project]).issues, [])

    def test_thousand_won_unit_prices_calculate(self):
        """산출내역의 단가는 천원이다. 요구내용 표는 원이다. calc 가 둘 다 받는다."""
        self.assertEqual(calc("50천원×5종×2회"), 500)
        self.assertEqual(calc("50,000원×5종×2회"), 500)

    def test_bullet_levels_are_learned_per_glyph(self):
        depths = {item.name: item.depth for item in self.items}
        self.assertEqual(depths["교육과정 실천중심 현장 지원"], 1)
        self.assertEqual(depths["운영용품"], 3)
        self.assertEqual(depths["초과"], 4)


@unittest.skipUnless(PLAN.exists(), "표본이 없으면 건너뛴다")
class OutlineDoesNotDisturbTheOldFormatTests(unittest.TestCase):
    def test_sample_still_uses_the_cost_code_tables(self):
        """'요구내용 및 산출근거' 표가 있으면 그쪽을 쓴다. 산출내역은 예비다."""
        projects = parse_plan(str(PLAN))
        self.assertEqual(sum(1 for p in projects if p.from_outline), 0)
        self.assertEqual(sum(1 for p in projects for i in p.items if i.is_leaf), 1118)


class CardScopeTests(unittest.TestCase):
    """단위과제카드 목록으로 '내가 입력할 사업'만 남긴다.

    K-에듀파인은 자기 카드에 딸린 사업만 입력할 수 있다. 과 전체 설명서 69개를
    늘어놓으면 자기 것을 찾는 데만 시간이 든다.
    """

    def test_names_match_across_spacing_and_suffixes(self):
        """설명서와 카드는 띄어쓰기·붙임표·꼬리 괄호가 제멋대로 다르다."""
        cards = CardList(cards=[Card("유보초이음교육운영(특교포함)"),
                                Card("영유아교육내실화지원(특교)"),
                                Card("유보통합정책운영관리")])
        self.assertIsNotNone(cards.find("유-보-초 이음교육 운영"))
        self.assertIsNotNone(cards.find("영유아교육내실화 지원"))
        self.assertIsNotNone(cards.find("유보통합정책 운영관리"))
        self.assertIsNone(cards.find("특수교육교재교구지원"))

    def test_card_without_a_project_is_flagged(self):
        """카드에는 있는데 설명서에 없으면 요구를 빠뜨렸을 수 있다."""
        cards = CardList(cards=[Card("유보통합시스템구축(특교)")])
        report = check_cards([PlanProject(name="영유아보육사업 이관")], cards)
        # 카드 목록에는 금액이 없어 요구를 빠뜨린 건지 올해 요구가 없는 건지 가릴 수 없다.
        # 오류라고 단정하지 않는다 (2.15.0). 기록만 남긴다.
        self.assertEqual(report.errors, [])
        self.assertTrue(report.dropped)

    def test_project_without_a_card_is_only_a_notice(self):
        """카드에 없는 사업은 오류가 아니다. 다른 담당자의 사업일 뿐이다."""
        cards = CardList(cards=[Card("유보통합정책운영관리")])
        report = check_cards([PlanProject(name="특수교육교재교구지원")], cards)
        self.assertEqual(report.issues, [], "오류가 아니면 목록에 싣지 않는다")

    def test_charge_is_the_part_after_the_department(self):
        self.assertEqual(PlanProject(name="가", organisation="유초등특수교육과, 유보통합담당").charge,
                         "유보통합담당")


class SequencePerCardTests(unittest.TestCase):
    """순번은 사업(단위과제카드)마다 1부터 다시 시작한다.

    K-에듀파인은 카드 한 장씩 열어 입력하고, 실제 다운로드 파일도 카드 한 장 분량이라
    순번이 1부터 시작한다. 문서 전체에 걸쳐 번호를 매기면 실제 화면과 어긋난다.
    """

    @unittest.skipUnless(PLAN.exists() and LAST_YEAR.exists(), "표본이 없으면 건너뛴다")
    def test_sequence_restarts_at_every_level_one_row(self):
        session = Session(plan_path=str(PLAN), last_year_path=str(LAST_YEAR))
        session.load()
        with tempfile.TemporaryDirectory() as folder:
            output = str(Path(folder) / "입력본.xlsx")
            session.save(output, allow_unsettled=True)
            sheet = load_workbook(output)["입력본"]
            self.assertEqual(sheet.cell(1, 1).value, "담당")
            self.assertEqual(sheet.cell(1, 2).value, "사업(카드)")
            previous = None
            for row in range(4, sheet.max_row + 1):
                level = str(sheet.cell(row, 3).value or "")
                sequence = sheet.cell(row, 4).value
                if level == "1":
                    self.assertEqual(sequence, 1, f"{row}행: 사업이 시작하면 순번은 1이어야 한다")
                elif previous is not None:
                    self.assertEqual(sequence, previous + 1, f"{row}행")
                previous = sequence


class TwoFormsTests(unittest.TestCase):
    """설명서는 같은 사업의 산출 내역을 두 서식으로 적는다. 둘 다 검산한다.

    실제로 있었던 일: 기능 점검용으로 사업계획 '산출내역' 표의 3,060 을 3,600 으로,
    산출식 `×2시간` 을 `×1시간` 으로, 머리글 총사업비 276,650 을 276,560 으로 고친
    설명서를 넣었는데 프로그램이 한 건도 잡지 못했다. '요구내용 및 산출근거' 표만
    읽고 나머지는 앞 표가 없을 때만 쓰는 예비 취급을 했기 때문이다.
    """

    @staticmethod
    def _plan(basis, outline, request=None, year_request=None):
        project = PlanProject(name="표본사업", request=request, year_request=year_request,
                              had_own_items=bool(basis))
        for depth, name, amount, formula in basis:
            project.items.append(PlanItem(depth=depth, name=name, amount=amount, formula=formula))
        for rows in outline:
            project.outline_items.append(
                [PlanItem(depth=depth, name=name, amount=amount, formula=formula, row=row)
                 for row, (depth, name, amount, formula) in enumerate(rows, start=1)])
        return project

    SOUND_BASIS = [(1, "절", 10000, ""), (2, "가", 4000, ""),
                   (3, "가1", 4000, "4,000,000원×1건"), (2, "나", 6000, "6,000,000원×1건")]

    def test_the_two_forms_must_agree(self):
        """한쪽 표만 고친 자리를 잡는다. 표 안에서는 각각 합이 맞아 검산으로는 안 걸린다."""
        project = self._plan(self.SOUND_BASIS,
                             [[(1, "절", 10000, ""), (2, "가", 4500, ""),
                               (3, "가1", 4000, "4,000천원×1건"), (2, "나", 6000, "6,000천원×1건")]],
                             request=10000, year_request=10000)
        errors = check_plan([project]).errors
        self.assertEqual(len(errors), 1, [issue.message for issue in errors])
        self.assertIn("요구내용 및 산출근거 표 4,000천원", errors[0].message)
        self.assertIn("산출내역 표 4,500천원", errors[0].message)

    def test_the_mismatch_points_at_the_right_row(self):
        """지적이 설명서 항목 목록의 '가' 줄(1번 자리)을 가리켜야 화면이 그 줄을 칠한다."""
        project = self._plan(self.SOUND_BASIS,
                             [[(1, "절", 10000, ""), (2, "가", 4500, ""),
                               (3, "가1", 4000, "4,000천원×1건"), (2, "나", 6000, "6,000천원×1건")]],
                             request=10000, year_request=10000)
        issue = check_plan([project]).errors[0]
        self.assertEqual(issue_row(project.items, issue), 1)

    def test_a_different_split_between_the_forms_is_not_an_error(self):
        """실제 설명서: 한 표는 '공립 1,402명' 한 줄, 다른 표는 '공립 202명 + 카드결제 1,200명'.

        위 단계 금액은 같다. 자식 수가 다르면 나눠 적은 방식이 다른 것이니 말하지 않는다.
        """
        project = self._plan(
            [(1, "절", 1682400, ""), (2, "공립", 1682400, "100,000원×1,402명×12월")],
            [[(1, "절", 1682400, ""), (2, "공립", 242400, "100천원×202명×12월"),
              (2, "공립카드결제", 1440000, "100천원×1,200명×12월")]],
            request=1682400, year_request=1682400)
        report = check_plan([project])
        self.assertEqual(report.errors, [], [issue.message for issue in report.errors])

    def test_the_outline_table_formula_is_checked_too(self):
        """산출내역 표의 산출식만 고쳐 놓아도 잡아야 한다(×2시간 -> ×1시간)."""
        project = self._plan(
            [(1, "절", 4480, ""), (2, "기본", 4480, "160,000원×14명×2시간")],
            [[(1, "절", 4480, ""), (2, "기본", 4480, "160천원×14명×1시간")]],
            request=4480, year_request=4480)
        errors = check_plan([project]).errors
        self.assertEqual(len(errors), 1, [issue.message for issue in errors])
        self.assertIn("산출내역 표", errors[0].message)
        self.assertIn("2,240천원", errors[0].message)

    def test_total_cost_must_equal_this_year_request(self):
        """총사업비는 여러 해의 합계가 아니라 올해 그 사업의 사업비다. 둘은 같아야 한다.

        자릿수를 바꿔 적는 실수가 여기서 걸린다(276,650 -> 276,560).
        """
        project = self._plan(self.SOUND_BASIS, [], request=9910, year_request=10000)
        report = check_plan([project])
        errors = [issue for issue in report.errors if "총사업비" in issue.message]
        self.assertEqual(len(errors), 1, [issue.message for issue in report.errors])
        self.assertIn("9,910천원", errors[0].message)
        self.assertIn("10,000천원", errors[0].message)

    def test_a_total_cost_larger_than_the_request_is_an_error_too(self):
        """크게 적힌 쪽도 오류다. 다른 사업 몫까지 더해 적은 설명서가 실제로 있다."""
        project = self._plan(self.SOUND_BASIS, [], request=32400, year_request=10000)
        errors = [issue for issue in check_plan([project]).errors if "총사업비" in issue.message]
        self.assertEqual(len(errors), 1)
        self.assertIn("22,400천원 차이", errors[0].message)

    def test_matching_total_cost_says_nothing(self):
        """같으면 아무 말도 하지 않는다. 표본 69개 중 67개가 이쪽이다."""
        project = self._plan(self.SOUND_BASIS, [], request=10000, year_request=10000)
        report = check_plan([project])
        self.assertEqual(report.issues, [], [issue.message for issue in report.issues])

    def test_a_folded_sub_page_headline_is_checked_too(self):
        """번호 없이 이어 붙은 세부 쪽의 머리글도 본다.

        실제로 있었던 일: '특수학교(유치원) 방학 중 방과후 과정 급식비 지원' 쪽은 머리글
        총사업비 81,900천원, 요구액 33,400천원이었다. 산출근거는 앞 사업에 합쳐지므로
        그 쪽 머리글은 아무도 보지 않는 칸이 되어 있었고, 오류가 조용히 지나갔다.
        증인은 그 쪽 요구액과 금액이 같은 1단 항목이다.
        """
        project = PlanProject(name="장애영유아교육지원", request=45400, year_request=45400,
                              had_own_items=True, split_description=True)
        project.items = [
            PlanItem(depth=1, name="유-보-초 이음교육 운영", amount=12000),
            PlanItem(depth=2, name="공립", amount=12000, formula="3,000,000원×4교"),
            PlanItem(depth=1, name="방학 중 방과후 과정 급식비 지원", amount=33400),
            PlanItem(depth=2, name="공립", amount=33400, formula="33,400,000원×1교")]
        project.parts = [PlanPart(name="유-보-초 이음교육 운영", request=12000, year_request=12000),
                         PlanPart(name="방학 중 방과후 과정 급식비 지원(총액배분사업비)",
                                  request=81900, year_request=33400)]
        report = check_plan([project])
        errors = [issue for issue in report.errors if "81,900" in issue.message]
        self.assertEqual(len(errors), 1, [issue.message for issue in report.errors])
        self.assertIn("머리글 총사업비 한 칸만 다릅니다", errors[0].message)
        self.assertEqual(errors[0].index, 2, "지적이 33,400천원짜리 1단 항목을 가리켜야 한다")
        self.assertEqual(report.blocking_errors, [])

    def test_a_sound_sub_page_says_nothing(self):
        """세부 쪽 머리글이 그 쪽 요구액과 같으면 아무 말도 하지 않는다."""
        project = PlanProject(name="장애영유아교육지원", request=45400, year_request=45400,
                              had_own_items=True)
        project.items = [PlanItem(depth=1, name="가", amount=45400, formula="45,400,000원×1교")]
        project.parts = [PlanPart(name="가 쪽", request=12000, year_request=12000),
                         PlanPart(name="나 쪽", request=33400, year_request=33400)]
        self.assertEqual(check_plan([project]).issues, [])

    def test_the_outline_row_repeating_a_wrong_total_is_not_counted_twice(self):
        """머리글 총사업비와 같은 값이 산출내역 표 요구액 칸에도 적힌 쪽.

        잘못된 칸은 둘이지만 뿌리는 하나다. 총사업비 오류 한 건으로만 세고, 이 줄은
        '같이 고쳐 달라'고 확인 필요로만 알린다. 덧셈이 틀린 것이 아니다.
        """
        project = self._plan(self.SOUND_BASIS,
                             [[(1, "절", 32400, ""), (2, "가", 4000, ""),
                               (3, "가1", 4000, "4,000천원×1건"),
                               (2, "나", 6000, "6,000천원×1건")]],
                             request=32400, year_request=10000)
        report = check_plan([project])
        self.assertEqual([issue.message for issue in report.errors
                          if "어긋납니다" in issue.message], [])
        self.assertEqual(len([issue for issue in report.errors if "총사업비" in issue.message]), 1)
        self.assertTrue(any("머리글 총사업비와 같은" in issue.message
                            for issue in report.dropped),
                        [issue.message for issue in report.dropped])

    def test_the_outline_table_is_kept_even_when_the_basis_table_exists(self):
        """예비가 아니라 항상 읽어 둔다. 여기가 비면 위의 검사들이 전부 잠든다."""
        if not PLAN.exists():
            self.skipTest("설명서 표본이 없으면 건너뛴다")
        projects = parse_plan(str(PLAN))
        both = [one for one in projects if one.had_own_items and one.outline_items]
        self.assertGreater(len(both), 20, "두 서식을 다 들고 있는 사업이 이렇게 적을 수 없다")

    def test_the_real_sample_only_flags_the_one_known_headline(self):
        """멀쩡한 설명서에 새 검사가 헛짚지 않는지. 여기가 늘어나면 오류 아닌 게 오류로 뜬다.

        표본 69개 사업에서 잡히는 오류는 26번 머리글 총사업비 한 건뿐이다. 그 사업은
        요구액(A)·투자실적 요구액·지원형태 계·집행계획 합계·산출근거 합계가 모두
        1,040,000천원인데 머리글만 3,240,000천원이다. 산출식·합계·두 서식 대조에서
        나오는 오류는 0건이어야 한다.
        """
        if not PLAN.exists():
            self.skipTest("설명서 표본이 없으면 건너뛴다")
        report = check_plan(parse_plan(str(PLAN)))
        others = [issue.message for issue in report.errors if "총사업비" not in issue.message]
        self.assertEqual(others, [])
        self.assertEqual(len(report.errors), 1, [issue.message for issue in report.errors])
        self.assertEqual(report.blocking_errors, [])


class SectionOrderTests(unittest.TestCase):
    """구역이 열 개를 넘으면 이름순은 문서 순서가 아니다.

    section10 이 section2 앞에 오면 사업 차례가 뒤섞이고, UBIS 대조가 엉뚱한 사업과
    붙는다. 표본이 69구역이라 눈에 띄지 않았을 뿐이다.
    """

    def test_sections_are_read_in_document_order(self):
        import zipfile
        paragraph = ('<?xml version="1.0" encoding="UTF-8"?>'
                     '<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"'
                     ' xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">'
                     '<hp:p><hp:run><hp:t>{text}</hp:t></hp:run></hp:p></hs:sec>')
        with tempfile.TemporaryDirectory() as folder:
            target = Path(folder) / "구역시험.hwpx"
            with zipfile.ZipFile(target, "w") as archive:
                for number in range(12):
                    archive.writestr(f"Contents/section{number}.xml",
                                     paragraph.format(text=f"구역{number}"))
            texts = [chunk["text"] for chunk in native_chunks(str(target))]
        self.assertEqual("\n".join(f"구역{number}" for number in range(12)), texts[0])


class ModuleSurfaceTests(unittest.TestCase):
    """클래스·함수가 통째로 사라지는 사고를 즉시 잡는다.

    실제로 났다. 소스를 인덱스로 잘라 치환하다 `plan_parser` 의 `PlanProject`
    정의가 필드째 날아갔다. 그때는 pyflakes 가 '정의되지 않은 이름'으로 잡아
    커밋해 둔 파일로 되돌렸지만, 지워진 것이 **쓰이지 않는 필드 하나**였다면
    아무 경고 없이 빌드까지 갔을 것이다. 있어야 할 것이 있는지 여기서 못 박는다.
    """

    def _fields(self, cls):
        return set(getattr(cls, "__dataclass_fields__", {}))

    def _properties(self, cls, names):
        for name in names:
            self.assertIsInstance(getattr(cls, name, None), property,
                                  f"{cls.__name__}.{name} 속성이 사라졌다")

    def test_plan_parser_surface(self):
        import plan_parser
        for name in ("PlanItem", "PlanPart", "PlanProject", "parse_plan",
                     "_fold_unnumbered", "_parse_basis_rows", "_parse_outline_rows",
                     "_learn_bullets", "_learn_outline_bullets", "_year_request"):
            self.assertTrue(hasattr(plan_parser, name), f"plan_parser.{name} 이 사라졌다")
        self.assertLessEqual(
            {"depth", "name", "bimok_name", "bimok_code5", "formula", "amount", "row", "leaf"},
            self._fields(plan_parser.PlanItem))
        self.assertLessEqual(
            {"name", "number", "order", "heading", "policy", "unit", "program", "organisation",
             "request", "year_request", "funding", "items", "parts", "outline_total",
             "outline_tables", "outline_items", "from_outline", "unnumbered",
             "split_description", "had_own_items"},
            self._fields(plan_parser.PlanProject))
        self._properties(plan_parser.PlanProject,
                         ("item_total", "base", "amount_candidates", "charge", "handover", "label"))
        self._properties(plan_parser.PlanItem, ("is_leaf", "handover", "code5"))

    def test_converter_surface(self):
        import converter
        self.assertLessEqual(
            {"plan_path", "ubis_path", "last_year_path", "extra_last_year", "class_path",
             "card_path", "projects", "ubis", "resolver", "resolutions",
             "overrides_by_code", "overrides_by_name", "cards", "all_projects"},
            self._fields(converter.Session))
        for name in ("load", "save", "output_rows", "card_of", "code_for", "blocking",
                     "warnings", "summary", "groups", "remembered", "apply_remembered",
                     "verify_generated", "verify_downloaded"):
            self.assertTrue(callable(getattr(converter.Session, name, None)),
                            f"Session.{name} 이 사라졌다")

    def test_other_modules_surface(self):
        import bimok_resolver, crosscheck, edufine_form, program_cards, ui_common
        for name in ("read_form", "write_form", "prefix", "basis_text", "form_total",
                     "FormRow", "OutputRow", "PREFIXES"):
            self.assertTrue(hasattr(edufine_form, name), f"edufine_form.{name} 이 사라졌다")
        self.assertLessEqual({"row", "level", "sequence", "name", "indent", "code",
                              "bimok_name", "detail", "basis", "amount"},
                             self._fields(edufine_form.FormRow))
        self.assertLessEqual({"level", "name", "code", "detail", "basis", "amount",
                              "note", "charge", "card"}, self._fields(edufine_form.OutputRow))
        for name in ("check_plan", "check_bimok", "check_generated", "check_downloaded",
                     "compare_totals", "diff_forms", "calc", "issue_row", "parent_name",
                     "children_of", "subtree_of", "leaf_total", "Issue", "Report",
                     "_families", "_check_forms", "_check_outline", "_check_request"):
            self.assertTrue(hasattr(crosscheck, name), f"crosscheck.{name} 이 사라졌다")
        self.assertLessEqual({"goal", "severity", "project", "item", "message",
                              "left", "right", "index", "form", "block", "row", "blocking"},
                             self._fields(crosscheck.Issue))
        self.assertIsInstance(getattr(crosscheck.Report, "blocking_errors", None), property)
        for name in ("by_parent", "by_code_and_name", "by_code", "by_bimok_name",
                     "by_parent_only", "by_name_only", "detail_by_code"):
            self.assertTrue(hasattr(bimok_resolver.BimokResolver(), name),
                            f"BimokResolver.{name} 이 사라졌다")
        for name in ("Card", "CardList", "load_cards", "check_cards", "squeeze"):
            self.assertTrue(hasattr(program_cards, name), f"program_cards.{name} 이 사라졌다")
        for name in ("money", "signed", "gap_color", "apply_tags", "FONT",
                     "ACCENT", "ERROR", "ERROR_BG", "ERROR_INK", "OK", "OK_BG", "WARN", "WARN_BG"):
            self.assertTrue(hasattr(ui_common, name), f"ui_common.{name} 이 사라졌다")

    def test_screen_module_surface(self):
        """화면은 tkinter 가 없는 곳에서도 '글자'로는 확인할 수 있다."""
        import ast
        for name in ("budget_bridge.py", "tutorial.py"):
            tree = ast.parse((HERE / name).read_text(encoding="utf-8"))
            top = {node.name for node in ast.walk(tree)
                   if isinstance(node, (ast.FunctionDef, ast.ClassDef))}
            self.assertIn("main" if name == "budget_bridge.py" else "Tutorial", top)
        source = (HERE / "budget_bridge.py").read_text(encoding="utf-8")
        methods = {node.name for node in ast.walk(ast.parse(source))
                   if isinstance(node, ast.FunctionDef)}
        for needed in ("_page_files", "_layout_cards", "_mark_files",
                       "pick", "clear", "clear_all", "load", "save", "verify", "export",
                       "show_tutorial", "open_detail", "_show_outline", "refresh", "refresh_bimok",
                       "next_issue", "previous_issue", "edit_code", "copy_find"):
            self.assertIn(needed, methods, f"budget_bridge 에서 {needed}() 가 사라졌다")
        for needed in ("FILE_CARDS", "FLOW", "self.progress_box", "self.year_note",
                       "self.file_marks", "self.paths"):
            self.assertTrue(needed in source, f"budget_bridge 에서 {needed} 가 사라졌다")
        self.assertEqual(source.count('("plan", "1"'), 1)
        self.assertTrue('("cards", "5"' in source, "5번 과제카드 카드가 사라졌다")

        # 화면은 native_ui 로 옮겼다. 거기 있어야 할 것도 같이 못 박는다.
        native = ast.parse((HERE / "native_ui.py").read_text(encoding="utf-8"))
        classes = {node.name: {child.name for child in node.body
                               if isinstance(child, ast.FunctionDef)}
                   for node in native.body if isinstance(node, ast.ClassDef)}
        for name in ("RoundedButton", "RoundedPanel", "ProcessFlow", "FilePage"):
            self.assertIn(name, classes, f"native_ui.{name} 이 사라졌다")
        for name in ("mark_files", "bind_wheel", "open_guide", "close_guide",
                     "toggle_route", "_resized"):
            self.assertIn(name, classes["FilePage"], f"FilePage.{name}() 이 사라졌다")
        for name in ("invoke", "configure", "cget"):
            self.assertIn(name, classes["RoundedButton"], f"RoundedButton.{name}() 이 사라졌다")
        bridge = (HERE / "budget_bridge.py").read_text(encoding="utf-8")
        for wired in ("self.progress_box", "self.year_note", "self.file_view"):
            self.assertTrue(wired in bridge, f"{wired} 연결이 끊겼다")


# --------------------------------------------------------------------------- 추경 점검 (2.14.0)
# 2025·2026년 1회·2회 추경 설명서 4부 · 추경 검토조서 4부 · 본예산 2부로 확인한 규칙이다.
# 표본이 있는 PC 에서는 실제 파일로, 없으면 아래 합성 표로 같은 규칙을 검사한다.

from supplement import (SuppDocument, SuppProject, calc as supp_calc, check_supplement, compare_previous,
                        compare_supplement, match_key, parse_change_table, parse_supplement,
                        read_supplement_review, round_of, _formula_ok)

PUA4, PUA5 = "\U000F02FB", ""


def change_table_rows(line_amount: str = "△2,590천원", group_change: str = "△2,590",
                      total_change: str = "△13,597") -> list:
    """2026 제2회 추경 '부서 운영' 표를 줄인 것. 이 모양이 네 부 모두 같다."""
    return [
        ["(단위 : 천원)"],
        ["사업명", "추가경정예 산 안(A)", "기  정예산액(B)(전년도예산액)", "비 교 증 감 내 역(A-B)", "비 고"],
        ["산출내역", "요구액"],
        ["합  계", "233,731", "247,328(180,642)", "", total_change, ""],
        ["부서기본운영경비", "44,120", "46,710(35,487)", "① 일반운영비", group_change, ""],
        ["․ 개인당(△740천원×7명×6/12개월)", line_amount],
        ["44,371", "47,048(35,415)", "② 특근매식비   (△9천원×7명×85회×6/12개월)", "△2,677천원", "△2,677", ""],
        ["6,480", "6,690(5,780)", "③ 부서운영경비", "△210", ""],
        ["․ 초과(△5천원×7명×6월)", "△210천원"],
        ["134,560", "142,680(99,760)", "④ 국내여비", "△8,120", ""],
        ["․ 관내  (△20천원×7명×33회×6/12개월)", "△2,310천원"],
        ["․ 도내  (△50천원×7명×14회×6/12개월)", "△2,450천원"],
        ["․ 도외  (△80천원×7명×12회×6/12개월)", "△3,360천원"],
        ["4,200", "4,200(4,200)", "⑤ 직책급업무수행경비", "0", ""],
    ]


def small_document(rows=None, section_cost=-13597.0) -> SuppDocument:
    table = parse_change_table(rows or change_table_rows())
    table.section, table.section_cost, table.order = "부서기본운영경비", section_cost, 1
    project = SuppProject(name="부서 운영", number="5", after=233731, before=247328, change=-13597,
                          total_cost=233731, tables=[table])
    return SuppDocument(projects=[project])


class SupplementCalcTests(unittest.TestCase):
    def test_triangle_is_minus(self):
        self.assertEqual(supp_calc("△740천원×7명×6/12개월"), -2590)

    def test_month_fraction(self):
        self.assertAlmostEqual(supp_calc("△9천원×7명×85회×6/12개월"), -2677.5)
        self.assertEqual(supp_calc("△400천원×1명×12개월"), -4800)

    def test_odd_change_is_subtracted(self):
        """2025 제1회 추경 원고료. 끝전 2천원을 뺀다."""
        self.assertEqual(supp_calc("14천원×3매×1시간×2명×3회-2천원"), 250)

    def test_won_unit(self):
        self.assertEqual(supp_calc("7,500원×740개교×3부"), 16650)

    def test_rounding_only_when_there_is_a_fraction(self):
        self.assertTrue(_formula_ok(-2677.5, -2677), "끝전 버림")
        self.assertTrue(_formula_ok(940623.25, 940624), "끝전 올림")
        self.assertFalse(_formula_ok(500, 501), "끝전이 없으면 1천원도 봐주지 않는다")

    def test_old_calculator_did_not_know_fractions(self):
        """본예산 계산기로 추경 줄을 재면 6/12 가 6 이 되어 12배로 튄다. 그래서 따로 둔다."""
        self.assertNotAlmostEqual(calc("9천원×7명×85회×6/12개월") or 0, 2677.5)


class SupplementTableTests(unittest.TestCase):
    def test_groups_lines_and_total_are_read(self):
        table = parse_change_table(change_table_rows())
        self.assertEqual(table.kind, "증감")
        self.assertEqual(len(table.groups), 5)
        self.assertEqual(table.total.change, -13597)
        first = table.groups[0]
        self.assertEqual((first.after, first.before, first.previous, first.change), (44120, 46710, 35487, -2590))
        self.assertEqual(first.business, "부서기본운영경비")
        second = table.groups[1]
        self.assertEqual(second.title, "② 특근매식비")
        self.assertEqual(second.formula, "△9천원×7명×85회×6/12개월")
        self.assertEqual(second.own, -2677)
        self.assertEqual([line.amount for line in table.groups[3].lines], [-2310, -2450, -3360])

    def test_clean_table_has_no_findings(self):
        report = check_supplement(small_document())
        self.assertEqual([one.message for one in report.issues if one.severity != "안내"], [])

    def test_formula_typo_is_caught(self):
        report = check_supplement(small_document(change_table_rows(line_amount="△2,950천원",
                                                                   group_change="△2,950",
                                                                   total_change="△13,957")))
        self.assertTrue(any("산출식" in one.message and one.severity == "오류" for one in report.issues))

    def test_group_change_must_equal_a_minus_b(self):
        rows = change_table_rows()
        rows[4][1] = "44,210"            # A 를 90 틀리게
        report = check_supplement(small_document(rows))
        self.assertTrue(any("추경안" in one.message and one.severity == "오류" for one in report.issues))

    def test_total_row_must_equal_the_sum(self):
        report = check_supplement(small_document(change_table_rows(total_change="△13,579")))
        self.assertTrue(any("합계" in one.message and one.severity == "오류" for one in report.issues))

    def test_plan_cost_must_equal_the_table(self):
        """2025 제2회 추경 '교원 처우개선비 지원' — 사업비 △327,022 · 표 △310,000. 실제로 있었다."""
        report = check_supplement(small_document(section_cost=-13000))
        self.assertTrue(any("사업비" in one.message and one.severity == "오류" for one in report.issues))

    def test_private_use_bullets_nest(self):
        rows = change_table_rows()[:4] + [
            ["역사교육", "20,000", "0(0)", "③ 역사교육 정책 및 체험처 발굴 연수", "20,000", ""],
            ["․ 운영용품(50천원×26종×2회)", "2,600천원"],
            ["․ 강사수당(일반)", "1,000천원"],
            [PUA4 + "기본(160천원×1명×2시간×2회)", "640천원"],
            [PUA4 + "초과(90천원×1명×2시간×2회)", "360천원"],
            ["․ 시설임차료(600천원×2실×2회)", "2,400천원"],
            ["․ 기타", "14,000천원"],
            [PUA4 + "강사수당(일반)", "640천원"],
            [PUA5 + "기본(160천원×1명×2시간×2회)", "640천원"],
            [PUA4 + "협의회(15천원×50명×4회)", "3,000천원"],
            [PUA4 + "잡비", "10,360천원"],
        ]
        rows[3] = ["합  계", "20,000", "0(0)", "", "20,000", ""]
        table = parse_change_table(rows)
        depths = [line.depth for line in table.groups[0].lines]
        self.assertEqual(depths, [1, 1, 2, 2, 1, 1, 2, 3, 2, 2])
        project = SuppProject(name="역사", after=20000, before=0, change=20000, tables=[table])
        report = check_supplement(SuppDocument(projects=[project]))
        self.assertEqual([one.message for one in report.issues if one.severity == "오류"], [])

    def test_unreflected_table(self):
        rows = [["(단위 : 천원)"],
                ["사업명", "추가경정예 산 안(A)", "기  정예산액(B)(전년도예산액)", "증 감(A-B)", "비 고"],
                ["합  계", "1,232,187", "1,232,187(9,172,630)", "0", ""],
                ["실내외 놀이 꿈터", "1,232,187", "1,232,187(8,261,575)", "0", ""],
                ["석교유아종합학습분원", "0", "0(911,055)", "0", ""]]
        table = parse_change_table(rows)
        self.assertEqual(table.kind, "미반영")
        self.assertEqual([group.title for group in table.groups], ["실내외 놀이 꿈터", "석교유아종합학습분원"])
        self.assertEqual(table.value("after"), 1232187)


class SupplementMatchTests(unittest.TestCase):
    def test_status_tags_are_dropped_but_funding_tags_are_kept(self):
        self.assertEqual(match_key("제지출금 등(신규)"), match_key("제지출금 등"))
        self.assertEqual(match_key("기초학력보장지원(성립전 포함)"), match_key("기초학력보장지원"))
        self.assertEqual(match_key("유치원 교육과정 운영(목변경 포함)"), match_key("유치원교육과정운영"))
        self.assertNotEqual(match_key("유치원교육과정운영(총액배분사업비)"), match_key("유치원교육과정운영"))
        self.assertEqual(match_key("신흥공원 유아숲체험원 조성(신규, 재원배분)"),
                         match_key("신흥공원 유아숲체험원 조성(재원배분)"))

    def test_round_from_file_name(self):
        self.assertEqual(round_of("2026본예산검토조서(서식1)_2026.9.24_19_14_54.xlsx"), (2026, 0))
        self.assertEqual(round_of("2026추경예산1차검토조서(서식1)_2026.9.24_19_15_24.xlsx"), (2026, 1))
        self.assertEqual(round_of("2026추경사업별 설명서_2026.9.24_19_35_45.hwpx"), (2026, 1))
        self.assertEqual(round_of("2025추경2차사업별 설명서_2026.9.24_19_45_26.hwpx"), (2025, 2))
        self.assertIsNone(round_of("2027사업별 설명서_2026.9.21_9_21_26.hwpx"))


def write_review(path: str, title: str, rows: list, supplement: bool = True) -> str:
    """UBIS 검토조서(서식1) 모양의 합성 파일. 5행이 머리글, 7행이 기관 합계."""
    from openpyxl import Workbook

    book = Workbook()
    sheet = book.active
    sheet.cell(1, 1, title)
    heads = (["기관명", "단위사업", "세부사업", "사업항목1\n(세세부사업)", "사업명", "기정액\n(A)",
              "'26년\n제2회 추경요구액\n(B)", "'26년\n최종예산액\n(C=A+B)", "조정액\n(D)", "배분액", "증감액",
              "사업항목 및 산출기초", "사유", "재원배분\n여부", "총액배분\n사업비"] if supplement else
             ["기관명", "단위사업", "세부사업", "사업항목1\n(세세부사업)", "사업명", "'25년본예산액(A)",
              "'25년최종예산액", "'26년본예산요구액(B)", "조정액(D)", "배분액", "증감액",
              "사업항목 및 산출기초", "사유", "재원배분\n여부", "총액배분\n사업비"])
    for column, head in enumerate(heads, start=1):
        sheet.cell(5, column, head)
    for number, values in enumerate(rows, start=7):
        for column, value in enumerate(values, start=1):
            if value is not None:
                sheet.cell(number, column, value)
    book.save(path)
    return path


class SupplementReviewTests(unittest.TestCase):
    def _review(self, folder, change=-13597.0, adjust=0.0):
        rows = [
            ["유초등특수교육과", None, None, None, None, 247328.0, change, 0, adjust],
            ["유초등특수교육과", "기본운영비", "기관기본운영비", "", "부서 운영", 247328.0, change, 0, adjust, 0, 0,
             "◎ 부서기본운영경비(자체비) △13,597천원"],
            [None] * 11 + ["  - (자체비)일반운영비 : △2,590천원"],
            [None] * 11 + ["  - (자체비)특근매식비 : △2,677천원"],
            [None] * 11 + ["  - (자체비)부서운영경비 : △210천원"],
            [None] * 11 + ["  - (자체비)국내여비 : △8,120천원"],
            ["유초등특수교육과", "유아교육", "유아교육운영", "유치원교육과정운영", None, 50180.0, 0, 0, 0, 0, 0,
             "◎  미편성사업", None, " 재원배분(본청) ", "  "],
            ["유초등특수교육과", "유아교육", "유아교육운영", "유치원교육과정운영", None, 999500.0, 0.0, 0, 0, 0, 0,
             "◎ 유아교육교육력제고(자체비_총액배분사업비) 0천원", None, "  ", " 총액배분사업비 "],
        ]
        return write_review(str(Path(folder) / "2026추경예산2차검토조서(서식1).xlsx"),
                            "2026년도 제2회 추경 요구사업 검토조서", rows)

    def test_review_rows_and_details(self):
        with tempfile.TemporaryDirectory() as folder:
            ubis = read_supplement_review(self._review(folder))
        self.assertEqual(ubis.total.change, -13597)
        names = [row.name for row in ubis.rows]
        self.assertIn("부서 운영", names)
        self.assertIn("유치원교육과정운영(재원배분)", names, "사업명이 빈 줄은 표시 열로 가른다")
        self.assertIn("유치원교육과정운영(총액배분사업비)", names)
        office = ubis.rows[0]
        self.assertEqual(office.pieces(), [-2590, -2677, -210, -8120])

    def test_clean_pair_has_no_findings(self):
        with tempfile.TemporaryDirectory() as folder:
            ubis = read_supplement_review(self._review(folder))
        report = compare_supplement(small_document(), ubis)
        self.assertEqual([one.message for one in report.issues if one.severity != "안내"], [])

    def test_change_mismatch_is_an_error(self):
        with tempfile.TemporaryDirectory() as folder:
            ubis = read_supplement_review(self._review(folder, change=-13579.0))
        report = compare_supplement(small_document(), ubis)
        self.assertTrue(any(one.severity == "오류" and "추경요구액" in one.message for one in report.issues))

    def test_main_review_is_refused_for_a_supplement_plan(self):
        from ubis_review import read_requests
        with tempfile.TemporaryDirectory() as folder:
            path = self._review(folder)
            with self.assertRaises(ValueError) as caught:
                read_requests(path)
        self.assertIn("추경", str(caught.exception))

    def test_previous_round_is_joined_by_program_and_item(self):
        """사업명은 차수마다 바뀐다. 세부사업·사업항목 단위로 합쳐서 잇는다."""
        with tempfile.TemporaryDirectory() as folder:
            current = self._review(folder)
            previous = write_review(str(Path(folder) / "2026추경예산1차검토조서(서식1).xlsx"),
                                    "2026년도 제1회 추경 요구사업 검토조서", [
                ["유초등특수교육과", None, None, None, None, 1, 1],
                ["유초등특수교육과", "기본운영비", "기관기본운영비", "", "부서운영", 257040.0, -9712.0],
                ["유초등특수교육과", "유아교육", "유아교육운영", "유치원교육과정운영", "", 1047500.0, -48000.0],
                ["유초등특수교육과", "유아교육", "유아교육운영", "유치원교육과정운영", "", 50180.0, 0.0],
            ])
            from supplement import find_previous_review
            self.assertEqual(Path(find_previous_review(current)).name, Path(previous).name)
            report = compare_previous(current, previous)
            self.assertEqual([one.message for one in report.issues if one.severity != "안내"], [])
            broken = write_review(previous, "2026년도 제1회 추경 요구사업 검토조서", [
                ["유초등특수교육과", None, None, None, None, 1, 1],
                ["유초등특수교육과", "기본운영비", "기관기본운영비", "", "부서운영", 257040.0, -9000.0],
                ["유초등특수교육과", "유아교육", "유아교육운영", "유치원교육과정운영", "", 1097680.0, -48000.0],
            ])
            report = compare_previous(current, broken)
        self.assertEqual(len([one for one in report.issues if one.severity == "오류"]), 1)


class AdjustmentTests(unittest.TestCase):
    """2025·2026 본예산 설명서의 금액 = UBIS 요구액 + 조정액(D). 20건이 모두 그랬다."""

    def test_difference_equal_to_adjustment_is_not_an_error(self):
        from crosscheck import Report
        session = Session()
        session.adjustments = {"초등교육과정 운영": 1080.0}
        report = Report()
        report.add("목표2A", "오류", "초등교육과정 운영", "",
                   "한글 설명서 95,730천원 · 엑셀 UBIS 검토조서 94,650천원 — 1,080천원 차이.", 95730, 94650)
        report.add("목표2A", "오류", "다른 사업", "",
                   "한글 설명서 1,000천원 · 엑셀 UBIS 검토조서 900천원 — 100천원 차이.", 1000, 900)
        session._explain_adjustments(report)
        self.assertEqual([one.project for one in report.issues], ["다른 사업"])
        self.assertIn("조정액", report.dropped[0].message)

    def test_adjustments_are_read(self):
        from ubis_review import read_adjustments
        with tempfile.TemporaryDirectory() as folder:
            path = write_review(str(Path(folder) / "2026본예산검토조서.xlsx"), "2026년도 본예산 요구사업 검토조서", [
                ["유초등특수교육과", None, None, None, None, 1, 1, 1, 1],
                ["유초등특수교육과", "교육과정운영", "교육과정운영지원", "초등교육과정운영", "초등교육과정 운영",
                 98518.0, 204413.0, 94650.0, 1080.0],
            ], supplement=False)
            self.assertEqual(read_adjustments(path), {"초등교육과정 운영": 1080.0})


SUPP_PLANS = {key: find_sample_like(prefix) for key, prefix in (
    ("2025-1", "2025추경사업별 설명서"), ("2025-2", "2025추경2차사업별 설명서"),
    ("2026-1", "2026추경사업별 설명서"), ("2026-2", "2026추경2차사업별 설명서"))}
SUPP_REVIEWS = {key: find_sample_like(prefix) for key, prefix in (
    ("2025-1", "2025추경예산1차검토조서"), ("2025-2", "2025추경예산2차검토조서"),
    ("2026-1", "2026추경예산1차검토조서"), ("2026-2", "2026추경예산2차검토조서"))}
HAVE_SUPP = all(path.exists() for path in list(SUPP_PLANS.values()) + list(SUPP_REVIEWS.values()))


@unittest.skipUnless(HAVE_SUPP, "추경 표본 8부가 없으면 건너뛴다")
class SupplementSampleTests(unittest.TestCase):
    """실제 네 차수. 숫자를 바꿔야 한다면 파서가 바뀐 것이다 — 이유를 README 에 적을 것."""

    @classmethod
    def setUpClass(cls):
        cls.documents = {key: parse_supplement(str(path)) for key, path in SUPP_PLANS.items()}

    def test_shape(self):
        expected = {"2025-1": (24, 213), "2025-2": (21, 195), "2026-1": (24, 154), "2026-2": (6, 19)}
        for key, (projects, groups) in expected.items():
            document = self.documents[key]
            self.assertEqual(len(document.projects), projects, key)
            self.assertEqual(sum(len(t.groups) for p in document.projects for t in p.tables), groups, key)

    def test_every_formula_is_computed(self):
        count = 0
        for document in self.documents.values():
            for project in document.projects:
                for table in project.tables:
                    for group in table.groups:
                        for formula in [group.formula] + [line.formula for line in group.lines]:
                            if formula:
                                count += 1
                                self.assertIsNotNone(supp_calc(formula), formula)
        self.assertEqual(count, 496)

    def test_clean_rounds_stay_clean(self):
        for key in ("2025-1", "2026-1", "2026-2"):
            report = check_supplement(self.documents[key])
            self.assertEqual([one.message for one in report.errors], [], key)

    def test_real_discrepancies_in_2025_second_round_are_found(self):
        """사람이 대조해서 확인한 실제 어긋남 세 건. 이것이 안 잡히면 검사가 잠든 것이다."""
        report = check_supplement(self.documents["2025-2"])
        messages = " / ".join(one.message for one in report.errors)
        # 총괄 표 증감·추경안 차이는 뿌리가 하나라 한 건으로 센다 (2.15.0 부터)
        self.assertEqual(len(report.errors), 2, messages)
        self.assertIn("△327,022", messages)       # 교원 처우개선비 지원 사업비 ≠ 표 △310,000
        self.assertIn("△1,781,232", messages)     # 유아체험장조성 총괄 ≠ 사업 쪽 △2,123,593

    def test_every_line_is_checked(self):
        """줄마다 금액을 1천원 틀리게 만들면 매번 잡혀야 한다. 검사가 줄을 건너뛰지 않는다는 증거."""
        document = self.documents["2026-1"]
        lines = [line for project in document.projects for table in project.tables
                 for group in table.groups for line in group.lines]
        self.assertEqual(len(lines), 130)
        before = len(check_supplement(document).errors)
        for line in lines:
            line.amount += 1
            after = len(check_supplement(document).errors)
            line.amount -= 1
            self.assertGreater(after, before, f"{line.name} 의 1천원 오타를 놓쳤다")

    def test_ubis_pairs(self):
        expected_errors = {"2025-1": 0, "2025-2": 1, "2026-1": 0, "2026-2": 0}
        for key, errors in expected_errors.items():
            ubis = read_supplement_review(str(SUPP_REVIEWS[key]))
            report = compare_supplement(self.documents[key], ubis)
            self.assertEqual(len(report.errors), errors, [one.message for one in report.errors])
        # 이음교육(총액배분): 쪽 274,000/309,000 · 총괄·UBIS 964,500/999,500. 증감은 같다.
        # 쪽이 UBIS 사업의 일부만 적은 것이라 오류가 아니다 — 싣지 않고 뺀 기록만 남는다.
        ubis = read_supplement_review(str(SUPP_REVIEWS["2026-2"]))
        report = check_supplement(self.documents["2026-2"], ubis)
        self.assertEqual(report.errors, [])
        self.assertEqual([one.project for one in report.dropped], ["이음교육 운영 지원(총액배분사업비)"])

    def test_every_round_continues_from_the_previous_one(self):
        """본예산 요구액 → 1회 기정, 1회 기정+요구 → 2회 기정. 네 차수 모두 1천원도 틀리지 않는다."""
        from supplement import find_previous_review
        for key, path in SUPP_REVIEWS.items():
            previous = find_previous_review(str(path))
            self.assertTrue(previous, f"{key} 직전 차수 검토조서를 못 찾았다")
            report = compare_previous(str(path), previous)
            self.assertEqual([one.message for one in report.issues if one.severity != "안내"], [], key)

    def test_session_runs_without_last_year_file(self):
        session = Session(plan_path=str(SUPP_PLANS["2026-2"]), ubis_path=str(SUPP_REVIEWS["2026-2"]))
        session.load()
        self.assertTrue(session.is_supplement)
        self.assertEqual(session.summary()["점검 종류"], "추경 점검 (2026년 제2회추경)")
        self.assertTrue(session.previous_used, "직전 차수를 자동으로 찾아야 한다")
        self.assertEqual(session.verify_generated("").issues, [])

    def test_supplement_review_with_main_plan_is_refused(self):
        plan = find_sample_like("2026본예산사업별 설명서")
        if not plan.exists():
            self.skipTest("본예산 표본이 없다")
        session = Session(plan_path=str(SUPP_PLANS["2026-2"]),
                          ubis_path=str(find_sample_like("2026본예산검토조서")))
        with self.assertRaises(ValueError) as caught:
            session.load()
        self.assertIn("추경", str(caught.exception))


@unittest.skipUnless(find_sample_like("2026본예산사업별 설명서").exists()
                     and find_sample_like("2026본예산검토조서").exists(), "본예산 표본이 없으면 건너뛴다")
class MainBudgetAdjustmentSampleTests(unittest.TestCase):
    def test_adjusted_projects_are_not_errors(self):
        session = Session(plan_path=str(find_sample_like("2026본예산사업별 설명서")),
                          ubis_path=str(find_sample_like("2026본예산검토조서")),
                          last_year_path=str(find_sample("2026(k에듀파인)세출예산요구내역(예산입력).xlsx")))
        if not Path(session.last_year_path).exists():
            self.skipTest("작년 확정본이 없다")
        session.load()
        report = session.verify_generated("")
        adjusted = [one for one in report.dropped if "조정액" in one.message]
        self.assertEqual(len(adjusted), 17)
        self.assertEqual(len(report.errors), 4, [one.message for one in report.errors])


# --------------------------------------------------------------------------- 오류만 싣는다 (2.15.0)

class ErrorsOnlyTests(unittest.TestCase):
    """목록에는 오류만. 확인 필요는 검사 자리에서 추론해 오류로 올리거나 빼고, 안내는 뺀다.
    검사를 못 한 것은 skipped 로 센다 — '오류 0건'이 '못 봤다'를 가리지 않게."""

    def test_report_keeps_only_errors(self):
        from crosscheck import SKIPPED, Report
        report = Report()
        report.add("목표1", "오류", "가", "", "틀림")
        report.add("목표1", "확인 필요", "나", "", "애매")
        report.add("목표1", "안내", "다", "", "참고")
        report.add("목표1", SKIPPED, "라", "", "못 봄")
        self.assertEqual([one.project for one in report.issues], ["가"])
        self.assertEqual([one.project for one in report.dropped], ["나", "다"])
        self.assertEqual([one.project for one in report.skipped], ["라"])

    def test_single_year_total_cost_mismatch_is_an_error(self):
        document = small_document()
        project = document.projects[0]
        project.total_cost, project.period = 247328, (2026, 2026)      # 기정 금액 그대로
        report = check_supplement(document)
        self.assertTrue(any("총사업비" in one.message and "고치지 않은" in one.message for one in report.errors))
        project.period = (2022, 2027)                                   # 계속비 사업이면 다를 수 있다
        self.assertFalse(any("총사업비" in one.message for one in check_supplement(document).errors))

    def test_unreflected_business_with_a_change_is_an_error(self):
        rows = [["(단위 : 천원)"],
                ["사업명", "추가경정예 산 안(A)", "기  정예산액(B)(전년도예산액)", "증 감(A-B)", "비 고"],
                ["실내외 놀이 꿈터", "1,232,187", "1,232,187(8,261,575)", "△5", ""]]
        table = parse_change_table(rows)
        table.order = 1
        project = SuppProject(name="가", tables=[table])
        report = check_supplement(SuppDocument(projects=[project]))
        self.assertTrue(any("미반영" in one.message or "반영하지 않은" in one.message for one in report.errors))

    def test_total_ab_gap_without_unreflected_table_is_not_an_error(self):
        """증감이 맞고 미반영 표가 없으면 A·B 차이는 손대지 않은 세부사업을 안 실은 것이다."""
        document = small_document()
        project = document.projects[0]
        project.after, project.before = 300000, 313597             # 증감 △13,597 그대로
        report = check_supplement(document)
        self.assertFalse([one for one in report.errors if "사업계획 표" in one.message])
        self.assertTrue(report.dropped)


def form_rows() -> list:
    """K-에듀파인 추경 다운로드의 머리글 두 줄."""
    return [
        ["레벨", "순번", "* 사업항목", None, "* 원가통계\n비목", None, "* 상세내역", None, "2026년도 기정", None,
         "2026년도 최종요구", None, "추경요구\n(②-①)", "성립전", "성립전 진행", "연계\n등록", "목적사업"],
        [None] * 8 + ["산출기초", "금액①", "* 산출기초", "* 금액②"] + [None] * 5,
    ]


class SupplementFormTests2(unittest.TestCase):
    def _write(self, folder, rows) -> str:
        from openpyxl import Workbook
        book = Workbook()
        sheet = book.active
        for values in rows:
            sheet.append(values)
        path = str(Path(folder) / "2026추경1차세출예산요구내역.xlsx")
        book.save(path)
        return path

    def _rows(self, after_amount=44120, change=-2590, basis="(740,000원×61명×12개월)-(740,000원×7명×6개월)=",
              level="2"):
        return form_rows() + [[None, None, "총계", None, None, None, None, None, None, 541680, None, 541680 - 31080,
                            -31080]] + [
            ["1", 1, "1. 부서 운영", None, None, None, None, None, None, 541680, None, 510600, -31080],
            [level, 2, "  가. 일반운영비", None, "2100101", None, "본청", None, "740,000원×61명×12개월=", 541680,
             basis, after_amount if after_amount != 44120 else 510600, change if change != -2590 else -31080],
        ]

    def test_clean_form(self):
        from supplement_form import check_form, read_supplement_form
        with tempfile.TemporaryDirectory() as folder:
            form = read_supplement_form(self._write(folder, self._rows()))
        self.assertEqual(len(form.rows), 2)
        self.assertEqual(form.rows[1].basis_after, "(740,000원×61명×12개월)-(740,000원×7명×6개월)=")
        self.assertEqual(check_form(form).issues, [])

    def test_bracketed_subtraction_is_computed(self):
        from supplement_form import basis_value
        self.assertAlmostEqual(basis_value("(184,000,000원×3개원)-(9,926,670원×3개원)="), 522219.99)
        self.assertEqual(basis_value("5,000,000원×20개원×1회+4,000,000원×19개원×1회="), 176000)
        self.assertEqual(basis_value("0="), 0)

    def test_wrong_amount_and_wrong_change_are_errors(self):
        from supplement_form import check_form, read_supplement_form
        with tempfile.TemporaryDirectory() as folder:
            form = read_supplement_form(self._write(folder, self._rows(after_amount=510700)))
        messages = " / ".join(one.message for one in check_form(form).issues)
        self.assertIn("산출기초", messages)

    def test_level_must_match_indent(self):
        from supplement_form import check_form, read_supplement_form
        with tempfile.TemporaryDirectory() as folder:
            form = read_supplement_form(self._write(folder, self._rows(level="3")))
        self.assertTrue(any("레벨" in one.message for one in check_form(form).issues))

    def test_main_form_is_refused_in_supplement_mode(self):
        from supplement_form import is_supplement_form
        self.assertFalse(is_supplement_form(str(find_sample("2026(k에듀파인)세출예산요구내역(예산입력).xlsx")))
                         if find_sample("2026(k에듀파인)세출예산요구내역(예산입력).xlsx").exists() else False)


EDUFINE_SUPP = [find_sample_like(f"2026추경1차세출예산요구내역{n}") for n in (1, 2, 3)]


@unittest.skipUnless(HAVE_SUPP and all(path.exists() for path in EDUFINE_SUPP),
                     "K-에듀파인 추경 입력본 표본이 없으면 건너뛴다")
class SupplementFormSampleTests(unittest.TestCase):
    """2026 제1회 추경을 K-에듀파인에 입력하고 내려받은 3부 (유보통합 시범 · 영유아교육 내실화 · 보육사업 이관)."""

    @classmethod
    def setUpClass(cls):
        from supplement_form import read_supplement_form
        cls.document = parse_supplement(str(SUPP_PLANS["2026-1"]))
        cls.ubis = read_supplement_review(str(SUPP_REVIEWS["2026-1"]))
        cls.forms = [read_supplement_form(str(path)) for path in EDUFINE_SUPP]

    def _errors(self, form):
        from supplement_form import check_form, compare_form
        report = check_form(form)
        compare_form(form, self.document, self.ubis, report)
        return report

    def test_shape(self):
        self.assertEqual([len(form.rows) for form in self.forms], [23, 141, 22])
        self.assertEqual([form.total.change for form in self.forms], [-362916, 521000, 257243])

    def test_entered_correctly(self):
        for form in self.forms:
            report = self._errors(form)
            self.assertEqual([one.message for one in report.issues], [], form.path)
            self.assertEqual(report.skipped, [], "산출기초를 모두 계산해야 한다")

    def test_every_one_thousand_won_typo_is_caught(self):
        """모든 금액 칸(①·②·추경요구)을 1천원 틀리게, 금액 있는 맨 아래 줄을 하나씩 지워 본다."""
        for form in self.forms:
            for row in form.rows:
                for key in ("before", "after", "change"):
                    value = getattr(row, key)
                    if value is None:
                        continue
                    setattr(row, key, value + 1)
                    caught = self._errors(form).issues
                    setattr(row, key, value)
                    self.assertTrue(caught, f"{row.row}행 {row.name} {key}")
            for index in range(len(form.rows)):
                if form.children(index) or not form.rows[index].change:
                    continue
                saved = form.rows.pop(index)
                caught = self._errors(form).issues
                form.rows.insert(index, saved)
                self.assertTrue(caught, f"{saved.row}행 {saved.name} 을 빠뜨린 것을 놓쳤다")

    def test_a_line_missing_from_the_plan_is_caught(self):
        project = next(one for one in self.document.projects if "영유아교육 내실화" in one.name)
        line = project.tables[0].groups[2].lines[0]
        line.amount += 5
        try:
            report = self._errors(self.forms[1])
        finally:
            line.amount -= 5
        self.assertTrue(any("입력본" in one.message and line.name in one.message for one in report.issues))
