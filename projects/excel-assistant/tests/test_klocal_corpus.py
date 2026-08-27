import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from excel_assistant.core import analyze_request
from excel_assistant.workbook import execute_plan, inspect_workbook


SAMPLE = Path(__file__).parent / "external_samples" / "klocal_정선군_2026.json"


@unittest.skipUnless(SAMPLE.exists(), "K-지방직 공개 검증 자료가 없습니다.")
class KLocalCorpusTests(unittest.TestCase):
    def test_real_budget_hierarchy_and_total_reconcile(self):
        source_rows = json.loads(SAMPLE.read_text(encoding="utf-8-sig"))
        self.assertGreaterEqual(len(source_rows), 3000)
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "klocal_budget.xlsx"
            output = Path(folder) / "result.xlsx"
            workbook = Workbook(write_only=True)
            sheet = workbook.create_sheet("예산데이터")
            sheet.append(["지자체", "부서", "정책사업", "단위사업", "세부사업", "예산목", "부기명", "중앙", "광역", "기초", "기타", "합계"])
            for row in source_rows:
                sheet.append([row[1], row[3], row[4], row[5], row[6], row[7], row[8], row[18], row[19], row[20], row[21], row[9]])
            workbook.save(source)

            info = inspect_workbook(source)["예산데이터"]
            plan = analyze_request("부서와 정책사업별 합계를 합산하고 큰 순서로 보여줘", info["headers"], info["profiles"])
            headers, rows, stats = execute_plan(source, "예산데이터", info["header_row"], plan, output, info["end_row"])
            self.assertEqual(headers[:2], ["부서", "정책사업"])
            self.assertEqual(stats["source_rows"], len(source_rows))
            self.assertAlmostEqual(sum(row[-1] for row in rows), sum(float(row[9] or 0) for row in source_rows))


if __name__ == "__main__":
    unittest.main()
