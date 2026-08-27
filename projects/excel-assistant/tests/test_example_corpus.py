import unittest
import tempfile
import os
from pathlib import Path

from excel_assistant.core import TaskPlan, analyze_request
from excel_assistant.workbook import execute_plan, inspect_workbook


EXAMPLE_ROOT = Path(os.environ.get("EXCEL_ASSISTANT_EXAMPLES", r"D:\자동화프로그램\8 예시파일"))


@unittest.skipUnless(EXAMPLE_ROOT.exists(), "로컬 예시파일 폴더가 없습니다.")
class ExampleCorpusTests(unittest.TestCase):
    def test_every_example_xlsx_can_be_inspected(self):
        files = list(EXAMPLE_ROOT.rglob("*.xlsx"))
        self.assertGreaterEqual(len(files), 10)
        for path in files:
            with self.subTest(path=path.name):
                info = inspect_workbook(path)
                self.assertTrue(info)
                for sheet in info.values():
                    self.assertGreaterEqual(sheet["header_row"], 1)
                    self.assertTrue(sheet["headers"])
                    self.assertEqual(set(sheet["headers"]), set(sheet["profiles"]))

    def test_fee_workbook_compound_request(self):
        path = next(EXAMPLE_ROOT.rglob("강의비_지급내역_가상개인정보.xlsx"))
        info = inspect_workbook(path)
        sheet_name = next(iter(info))
        sheet = info[sheet_name]
        plan = analyze_request(
            "이름을 기준으로 큰 순서대로 보여주고 최대값과 최소값 평균값을 알려줘",
            sheet["headers"], sheet["profiles"],
        )
        self.assertEqual(plan.group_column, "성명")
        self.assertEqual(plan.value_column, "강의비")
        output = Path(self.id().replace(".", "_") + ".xlsx")
        try:
            _, rows, stats = execute_plan(path, sheet_name, sheet["header_row"], plan, output)
            self.assertEqual(len(rows), 10)
            self.assertEqual(stats["statistics"]["maximum"], 520000.0)
            self.assertEqual(stats["statistics"]["minimum"], 330000.0)
            self.assertEqual(stats["statistics"]["average"], 410000.0)
        finally:
            output.unlink(missing_ok=True)

    def test_supported_operations_execute_across_example_corpus(self):
        with tempfile.TemporaryDirectory() as folder:
            output_root = Path(folder)
            executed = 0
            for file_index, path in enumerate(EXAMPLE_ROOT.rglob("*.xlsx")):
                info = inspect_workbook(path)
                for sheet_index, (sheet_name, sheet) in enumerate(info.items()):
                    source_sheet = sheet.get("source_sheet", sheet_name)
                    end_row = sheet.get("end_row")
                    text_columns = [name for name, profile in sheet["profiles"].items() if profile["kind"] == "text"]
                    numeric_columns = [name for name, profile in sheet["profiles"].items() if profile["kind"] == "number"]
                    if text_columns:
                        output = output_root / f"sort_{file_index}_{sheet_index}.xlsx"
                        execute_plan(path, source_sheet, sheet["header_row"], TaskPlan(operation="sort", sort_column=text_columns[0]), output, end_row)
                        self.assertTrue(output.exists())
                        executed += 1
                    if text_columns and numeric_columns:
                        output = output_root / f"sum_{file_index}_{sheet_index}.xlsx"
                        plan = TaskPlan(operation="group_sum", group_column=text_columns[0], group_columns=(text_columns[0],), value_column=numeric_columns[0])
                        execute_plan(path, source_sheet, sheet["header_row"], plan, output, end_row)
                        self.assertTrue(output.exists())
                        executed += 1
            self.assertGreaterEqual(executed, 15)
