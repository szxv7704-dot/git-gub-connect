from __future__ import annotations

import json
import sys
from pathlib import Path

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).parents[1]))

from workflow_core import (  # noqa: E402
    WorkflowError, apply_clean_plan, merge_files, plan_cleaning, read_table, save_table,
)


SAMPLE = Path(__file__).parents[2] / "excel-assistant" / "tests" / "external_samples" / "klocal_정선군_2026.json"
HEADERS = ["지자체", "부서", "정책사업", "단위사업", "세부사업", "예산목", "부기명", "중앙", "광역", "기초", "기타", "합계"]


def write_budget(path: Path, rows: list[list[object]]) -> None:
    book = Workbook(write_only=True)
    sheet = book.create_sheet("예산데이터")
    sheet.append(HEADERS)
    for row in rows:
        sheet.append([row[1], row[3], row[4], row[5], row[6], row[7], row[8], row[18], row[19], row[20], row[21], row[9]])
    book.save(path)


def test_klocal_full_corpus_merge_group_sum_sort_and_roundtrip(tmp_path: Path) -> None:
    source_rows = json.loads(SAMPLE.read_text(encoding="utf-8-sig"))
    assert len(source_rows) >= 3000
    split = len(source_rows) // 2
    first, second = tmp_path / "정선군_1.xlsx", tmp_path / "정선군_2.xlsx"
    write_budget(first, source_rows[:split])
    write_budget(second, source_rows[split:])

    merged = merge_files([first, second], include_source=False)
    assert len(merged.rows) == len(source_rows)
    plan = plan_cleaning("부서와 정책사업별 합계를 합산하고 큰 순서로 정렬해줘", merged)
    assert not plan.warnings
    cleaned = apply_clean_plan(merged, plan)
    assert cleaned.headers == ["부서", "정책사업", "합계 합계"]
    assert sum(float(row[-1]) for row in cleaned.rows) == sum(float(row[9] or 0) for row in source_rows)
    assert all(cleaned.rows[i][-1] >= cleaned.rows[i + 1][-1] for i in range(len(cleaned.rows) - 1))

    output = save_table(cleaned, tmp_path / "정선군_검증결과.xlsx")
    assert len(read_table(output).rows) == len(cleaned.rows)


def test_downloaded_klocal_pdf_is_rejected_with_clear_format_message() -> None:
    path = Path.home() / "AppData" / "Local" / "Temp" / "klocal_jeongseon_2026.pdf"
    if not path.exists():
        return
    try:
        read_table(path)
    except WorkflowError as exc:
        assert "지원하지 않는 파일 형식" in str(exc)
    else:
        raise AssertionError("PDF가 표 파일로 잘못 처리되었습니다.")
