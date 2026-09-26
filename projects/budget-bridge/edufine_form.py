"""K-에듀파인 세출예산요구내역 양식을 읽고 쓴다.

서식 규칙은 2026년 확정 다운로드 파일 33행에서 역공학해 확인했다.

    레벨 = 사업항목 들여쓰기 칸수 ÷ 2 + 1
    순번 = 4행부터 전역 일련번호 1, 2, 3 …   (형제 내 번호가 아니다)

그리드 1단은 **사업명**이다. 정책·단위·세부사업은 화면 상단에서 이미 선택된
값이라 이 표에 행으로 들어가지 않는다.

열은 위치가 아니라 헤더 이름으로 찾는다. 연도가 바뀌어 열이 늘어도 깨지지 않는다.
"""

from __future__ import annotations

from copy import copy
from dataclasses import dataclass
from typing import Iterable, Optional

from openpyxl import load_workbook
from openpyxl.worksheet.cell_range import CellRange

# 계층 접두 기호. 1단부터 6단까지.
PREFIXES = ["{n}.", "{k}.", "{n})", "{k})", "({n})", "({k})"]
KOREAN = list("가나다라마바사아자차카타파하거너더러머버서어저처커터퍼허")
HEADER_ROWS = 3          # 1~2행 헤더, 3행 총계
FIRST_DATA_ROW = 4


@dataclass
class FormRow:
    row: int
    level: Optional[int]
    sequence: Optional[int]
    name: str
    indent: int
    code: str = ""
    bimok_name: str = ""      # '[2100607] 기타운영수당' 처럼 이름이 함께 적힌 서식용
    detail: str = ""
    basis: str = ""
    amount: Optional[float] = None


def _columns(sheet) -> dict[str, int]:
    """헤더 이름으로 열 번호를 찾는다."""
    found: dict[str, int] = {}
    carried = ""
    for column in range(1, sheet.max_column + 1):
        top = str(sheet.cell(1, column).value or "").replace("\n", "").replace(" ", "")
        sub = str(sheet.cell(2, column).value or "").replace("\n", "").replace(" ", "")
        # 1행은 '2026년도 요구'처럼 두 열에 걸쳐 병합돼 있어 오른쪽 셀이 비어 있다.
        # 직전 값을 이어서 써야 '산출기초'가 전년도 것인지 요구분인지 구분된다.
        if top:
            carried = top
        elif sub:
            top = carried
        label = (top + sub).replace("*", "")
        if "레벨" in label and "레벨" not in found:
            found["레벨"] = column
        elif "순번" in label and "순번" not in found:
            found["순번"] = column
        elif "사업항목" in label and "사업항목" not in found:
            found["사업항목"] = column
        elif "원가통계비목" in label and "비목" not in found:
            found["비목"] = column
        elif "상세내역" in label and "상세내역" not in found:
            found["상세내역"] = column
        elif "요구" in label and "산출기초" in label and "산출기초" not in found:
            found["산출기초"] = column
        elif "요구" in label and "금액" in label and "금액" not in found:
            found["금액"] = column
    missing = {"레벨", "순번", "사업항목", "비목", "상세내역", "산출기초", "금액"} - set(found)
    if missing:
        raise ValueError("K-에듀파인 세출예산요구내역 양식이 아닙니다. 찾지 못한 열: " + ", ".join(sorted(missing)))
    return found


def read_form(path: str) -> list[FormRow]:
    """K-에듀파인에서 내려받은 파일(또는 이 프로그램이 만든 파일)을 읽는다."""
    sheet = load_workbook(path, data_only=True).active
    columns = _columns(sheet)
    rows: list[FormRow] = []
    for number in range(FIRST_DATA_ROW, sheet.max_row + 1):
        name = sheet.cell(number, columns["사업항목"]).value
        if not isinstance(name, str) or not name.strip():
            continue
        level = sheet.cell(number, columns["레벨"]).value
        sequence = sheet.cell(number, columns["순번"]).value
        amount = sheet.cell(number, columns["금액"]).value
        rows.append(FormRow(
            row=number,
            level=int(level) if str(level or "").strip().isdigit() else None,
            sequence=int(sequence) if str(sequence or "").strip().replace(".0", "").isdigit() else None,
            name=name.strip(),
            indent=len(name) - len(name.lstrip()),
            code=str(sheet.cell(number, columns["비목"]).value or "").strip(),
            detail=str(sheet.cell(number, columns["상세내역"]).value or "").strip(),
            basis=str(sheet.cell(number, columns["산출기초"]).value or "").strip(),
            amount=float(amount) if isinstance(amount, (int, float)) else None,
        ))
    return rows


def form_total(path: str) -> Optional[float]:
    """3행 총계의 요구 금액."""
    sheet = load_workbook(path, data_only=True).active
    value = sheet.cell(HEADER_ROWS, _columns(sheet)["금액"]).value
    return float(value) if isinstance(value, (int, float)) else None


def prefix(level: int, order: int) -> str:
    """같은 부모 아래 order번째(1부터) 항목의 계층 기호."""
    pattern = PREFIXES[min(level, len(PREFIXES)) - 1]
    return pattern.format(n=order, k=KOREAN[(order - 1) % len(KOREAN)])


def basis_text(formula: str, amount: Optional[float]) -> str:
    """K-에듀파인 산출기초 표기. 금액 0인 행은 확정본에서 '0=' 으로 저장된다."""
    cleaned = (formula or "").strip()
    if not cleaned:
        return "0=" if (amount or 0) == 0 else ""
    return cleaned if cleaned.endswith("=") else cleaned + "="


@dataclass
class OutputRow:
    level: int
    name: str
    code: str = ""
    bimok_name: str = ""      # '[2100607] 기타운영수당' 처럼 이름이 함께 적힌 서식용
    detail: str = ""
    basis: str = ""
    amount: Optional[float] = None
    note: str = ""          # 미확정 사유 등 (대본 시트에만 쓴다)
    charge: str = ""        # 담당. 필터용
    card: str = ""          # 사업(단위과제카드). 필터용


def build_rows(project_name: str, items: Iterable) -> list[OutputRow]:
    """설명서 항목을 K-에듀파인 계층으로 바꾼다.

    그리드 1단은 사업명이고, 설명서의 1./①/․/▸ 네 단계가 2~5단으로 실린다.
    세 단계로 뭉개면 한 계층이 사라진다.
    """
    rows: list[OutputRow] = [OutputRow(level=1, name=project_name)]
    counters: dict[int, int] = {}
    for item in items:
        level = min(item.depth + 1, len(PREFIXES))
        counters[level] = counters.get(level, 0) + 1
        for deeper in list(counters):
            if deeper > level:
                counters.pop(deeper)
        rows.append(OutputRow(
            level=level,
            name=f"{prefix(level, counters[level])} {item.name}".strip(),
            amount=item.amount,
        ))
    return rows


def write_form(template: str, output: str, sheets: list[tuple[str, list[OutputRow]]],
               handover: Optional[list] = None) -> None:
    """작년 양식을 복제해 올해 입력본을 만든다.

    `입력본` 시트 맨 앞에 `담당` · `사업(카드)` 두 열을 붙이고 자동필터를 건다.
    프로그램을 과 전체가 나눠 쓰므로 담당자가 자기 것만 걸러 봐야 한다. 이 파일은
    K-에듀파인에 올리는 게 아니라 사람이 보고 손으로 치는 종이라 열을 더해도 된다.
    입력 후 대조(목표 2B)는 열을 이름으로 찾으므로 열이 늘어도 깨지지 않는다.

    **순번은 사업(단위과제카드)마다 1부터 다시 매긴다.** K-에듀파인은 카드 한 장씩
    열어 입력하고, 실제 다운로드 파일도 카드 한 장 분량이라 순번이 1부터 시작한다.
    문서 전체에 걸쳐 번호를 매기면 실제 화면과 어긋난다.
    """
    book = load_workbook(template)
    sheet = book.active
    sheet.title = "입력본"
    columns = _columns(sheet)
    styles = {column: (copy(sheet.cell(FIRST_DATA_ROW, column)._style),
                       sheet.cell(FIRST_DATA_ROW, column).number_format)
              for column in range(1, sheet.max_column + 1)}
    if sheet.max_row >= FIRST_DATA_ROW:
        sheet.delete_rows(FIRST_DATA_ROW, sheet.max_row - HEADER_ROWS)

    all_rows = [row for _, rows in sheets for row in rows]
    line = FIRST_DATA_ROW
    for _card, block in sheets:
        for sequence, row in enumerate(block, start=1):
            for column, (style, number_format) in styles.items():
                cell = sheet.cell(line, column)
                cell._style = copy(style)
                cell.number_format = number_format
            sheet.cell(line, columns["레벨"]).value = str(row.level)
            sheet.cell(line, columns["순번"]).value = sequence
            sheet.cell(line, columns["사업항목"]).value = " " * ((row.level - 1) * 2) + row.name
            sheet.cell(line, columns["비목"]).value = row.code or None
            sheet.cell(line, columns["상세내역"]).value = row.detail or None
            sheet.cell(line, columns["산출기초"]).value = row.basis or None
            sheet.cell(line, columns["금액"]).value = round(row.amount) if row.amount is not None else None
            line += 1
    sheet.cell(HEADER_ROWS, columns["금액"]).value = round(sum(
        row.amount or 0 for row in all_rows if row.level == 1))

    _add_filter_columns(sheet, all_rows)
    _handover_sheet(book, handover or [])
    _guide_sheet(book, sheets, all_rows)
    book.save(output)


def _add_filter_columns(sheet, rows: list[OutputRow]) -> None:
    """맨 앞에 담당·사업(카드) 열을 끼우고 자동필터를 건다."""
    from openpyxl.styles import Alignment, Font, PatternFill

    # openpyxl 의 insert_cols 는 값만 옮기고 **병합 범위는 옛 자리에 둔다.** 그대로 두면
    # 머리글 칸이 MergedCell 이 되어 값이 None 으로 읽히고, 만든 파일을 다시 읽을 때
    # '산출기초·금액 열을 찾지 못했다'며 터진다. 풀었다가 두 칸 밀어서 다시 병합한다.
    merged = [str(area) for area in sheet.merged_cells.ranges]
    for area in merged:
        sheet.unmerge_cells(area)
    sheet.insert_cols(1, 2)
    for area in merged:
        moved = CellRange(area)
        moved.shift(col_shift=2)
        sheet.merge_cells(str(moved))
    for column, title, width in ((1, "담당", 20), (2, "사업(카드)", 34)):
        cell = sheet.cell(1, column)
        cell.value = title
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="EDEFF2")
        cell.alignment = Alignment(vertical="center")
        sheet.column_dimensions[cell.column_letter].width = width
    sheet.cell(HEADER_ROWS, 1).value = "← 자기 담당으로 걸러 보세요"
    sheet.cell(HEADER_ROWS, 1).font = Font(size=9, color="6B7280")
    for offset, row in enumerate(rows):
        sheet.cell(FIRST_DATA_ROW + offset, 1).value = row.charge or None
        sheet.cell(FIRST_DATA_ROW + offset, 2).value = row.card or None
    last = FIRST_DATA_ROW + len(rows) - 1
    if rows:
        sheet.auto_filter.ref = f"A{HEADER_ROWS}:B{last}"
    sheet.freeze_panes = f"C{FIRST_DATA_ROW}"


def _guide_sheet(book, sheets: list, all_rows: list[OutputRow]) -> None:
    """한 줄씩 짚어 가며 입력하라고 만든 표. 여기도 담당·사업으로 걸러진다."""
    from openpyxl.styles import Font, PatternFill

    guide = book.create_sheet("대본")
    missing = sum(1 for row in all_rows if not row.code and row.basis)
    guide.append([f"전체 {len(all_rows):,}행 · 비목 미확정 {missing:,}행"
                  + ("  ← 빨간 줄은 K-에듀파인에서 비목을 직접 골라야 합니다" if missing else "  · 전부 확정")])
    guide["A1"].font = Font(bold=True, size=12, color="C0392B" if missing else "1E7A4C")
    guide.append(["담당", "사업(카드)", "순번", "레벨", "사업항목", "원가통계비목",
                  "상세내역", "산출기초", "금액(천원)", "확인", "비고"])
    for cell in guide[2]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="EDEFF2")
    unsettled = PatternFill("solid", fgColor="FFC7CE")
    for _card, block in sheets:
        for sequence, row in enumerate(block, start=1):
            guide.append([row.charge, row.card, sequence, row.level,
                          " " * ((row.level - 1) * 2) + row.name,
                          row.code or "?", row.detail or "", row.basis, row.amount, "☐", row.note])
            if not row.code and row.basis:
                for cell in guide[guide.max_row]:
                    cell.fill = unsettled
    for column, width in zip("ABCDEFGHIJK", (20, 34, 6, 6, 46, 14, 12, 34, 12, 6, 40)):
        guide.column_dimensions[column].width = width
    if all_rows:
        guide.auto_filter.ref = f"A2:K{len(all_rows) + 2}"
    guide.freeze_panes = "C3"


def _handover_sheet(book, sheets: list) -> None:
    """이 과에서 K-에듀파인에 입력하지 않는 사업.

    총액배분사업비는 예산과가 학교로 바로 재배정하고, 재원배분 사업은 다른 과가
    나눠 준다. 입력본 시트에 섞어 두면 담당자가 그대로 입력해 버린다. 그렇다고
    파일에서 통째로 없애면 '설명서에 있던 사업이 왜 없지' 하고 다시 찾게 된다.
    그래서 빼되, 뺐다는 사실과 사유를 이 시트에 남긴다.
    """
    from openpyxl.styles import Font, PatternFill

    rows = [row for _name, block in sheets for row in block]
    sheet = book.create_sheet("재배정(입력 안 함)")
    total = sum(row.amount or 0 for row in rows if row.level == 1)
    sheet.append([f"K-에듀파인에 입력하지 않는 사업 {len(sheets)}건 · {len(rows):,}행 · "
                  f"합계 {round(total):,}천원"])
    sheet["A1"].font = Font(bold=True, size=12, color="6B4A08")
    sheet.append(["아래는 입력본 시트에 들어 있지 않습니다. 예산과가 학교로 재배정하거나 "
                  "다른 과가 나눠 주는 사업입니다."])
    sheet["A2"].font = Font(size=10, color="6B7280")
    sheet.append([])
    sheet.append(["사유", "레벨", "사업항목", "원가통계비목", "산출기초", "금액(천원)"])
    for cell in sheet[4]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="EDEFF2")
    for _name, block in sheets:
        for row in block:
            sheet.append([row.note if row.level == 1 else "", row.level,
                          " " * ((row.level - 1) * 2) + row.name,
                          row.code or "", row.basis, row.amount])
    for column, width in zip("ABCDEF", (16, 6, 50, 16, 34, 13)):
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A5"
