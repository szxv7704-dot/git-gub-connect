import unittest

from excel_assistant.core import RequestError, TaskPlan, analyze_request, validate_task_plan


HEADERS = ["이름", "부서", "상태", "금액", "날짜"]


class AnalyzeRequestTests(unittest.TestCase):
    def test_group_sum(self):
        plan = analyze_request("이름을 기준으로 금액을 합산하고 큰 순서대로 보여줘", HEADERS)
        self.assertEqual(plan.operation, "group_sum")
        self.assertEqual(plan.group_column, "이름")
        self.assertEqual(plan.value_column, "금액")
        self.assertTrue(plan.descending)

    def test_filter(self):
        plan = analyze_request("상태가 완료인 행만 보여줘", HEADERS)
        self.assertEqual(plan.operation, "filter")
        self.assertEqual(plan.filter_column, "상태")
        self.assertEqual(plan.filter_value, "완료")

    def test_sort(self):
        plan = analyze_request("금액을 큰 순서로 정렬해줘", HEADERS)
        self.assertEqual(plan.operation, "sort")
        self.assertEqual(plan.sort_column, "금액")
        self.assertTrue(plan.descending)

    def test_duplicates(self):
        plan = analyze_request("이름 중복 항목을 찾아줘", HEADERS)
        self.assertEqual(plan.operation, "duplicates")
        self.assertEqual(plan.group_column, "이름")

    def test_compound_statistics_infers_only_numeric_column(self):
        profiles = {
            "성명": {"kind": "text"},
            "주민등록번호": {"kind": "text"},
            "강의비": {"kind": "number"},
            "입금은행": {"kind": "text"},
        }
        plan = analyze_request(
            "이름을 기준으로 큰 순서대로 보여주고 최대값과 최소값 평균값을 알려줘",
            list(profiles), profiles,
        )
        self.assertEqual(plan.operation, "group_sum")
        self.assertEqual(plan.group_column, "성명")
        self.assertEqual(plan.value_column, "강의비")
        self.assertTrue(plan.descending)
        self.assertTrue(plan.include_statistics)

    def test_group_average_with_explicit_numeric_column(self):
        profiles = {"부서명": {"kind": "text"}, "금액": {"kind": "number"}, "수량": {"kind": "number"}}
        plan = analyze_request("부서별 금액 평균을 보여줘", list(profiles), profiles)
        self.assertEqual(plan.group_column, "부서명")
        self.assertEqual(plan.value_column, "금액")
        self.assertEqual(plan.aggregation, "average")

    def test_group_count(self):
        profiles = {"부서": {"kind": "text"}, "성명": {"kind": "text"}}
        plan = analyze_request("부서별 인원수를 알려줘", list(profiles), profiles)
        self.assertEqual(plan.operation, "group_sum")
        self.assertEqual(plan.group_column, "부서")
        self.assertEqual(plan.aggregation, "count")

    def test_numeric_filter(self):
        plan = analyze_request("금액이 1000 이상인 행만 보여줘", HEADERS)
        self.assertEqual(plan.operation, "filter")
        self.assertEqual(plan.filter_column, "금액")
        self.assertEqual(plan.filter_operator, "gte")
        self.assertEqual(plan.filter_value, 1000.0)

    def test_negative_filter(self):
        plan = analyze_request("상태가 완료가 아닌 항목만 보여줘", HEADERS)
        self.assertEqual(plan.filter_operator, "ne")
        self.assertEqual(plan.filter_value, "완료")

    def test_duplicate_alias_after_keyword(self):
        plan = analyze_request("중복된 성명을 찾아줘", ["성함", "소속"])
        self.assertEqual(plan.group_column, "성함")

    def test_multi_column_grouping(self):
        profiles = {"부서": {"kind": "text"}, "세부사업": {"kind": "text"}, "금액": {"kind": "number"}}
        plan = analyze_request("부서와 세부사업별 금액 합계를 보여줘", list(profiles), profiles)
        self.assertEqual(plan.group_columns, ("부서", "세부사업"))
        self.assertEqual(plan.value_column, "금액")

    def test_site_inspired_funding_sum_check(self):
        profiles = {name: {"kind": "number"} for name in ("중앙", "광역", "기초", "기타", "합계")}
        plan = analyze_request("중앙, 광역, 기초, 기타 금액의 합계가 맞는지 확인해줘", list(profiles), profiles)
        self.assertEqual(plan.operation, "sum_check")
        self.assertEqual(plan.component_columns, ("중앙", "광역", "기초", "기타"))
        self.assertEqual(plan.total_column, "합계")

    def test_plan_validation_rejects_missing_and_invalid_columns(self):
        with self.assertRaises(RequestError):
            validate_task_plan(TaskPlan(operation="sort", sort_column="없는열"), HEADERS)
        with self.assertRaises(RequestError):
            validate_task_plan(TaskPlan(operation="filter", filter_column="금액", filter_value="천원", filter_operator="gte"), HEADERS)

    def test_plan_validation_accepts_multi_group(self):
        plan = TaskPlan(operation="group_sum", group_columns=("부서", "이름"), value_column="금액")
        validate_task_plan(plan, HEADERS)

    def test_alias_ambiguity_requires_user_confirmation(self):
        profiles = {"성명": {"kind": "text"}, "담당자": {"kind": "text"}, "금액": {"kind": "number"}}
        plan = analyze_request("이름을 기준으로 금액을 합산해줘", list(profiles), profiles)
        self.assertLess(plan.confidence_margin, 0.08)
        self.assertTrue(plan.ambiguous_columns)


if __name__ == "__main__":
    unittest.main()
