from __future__ import annotations

import csv
import hashlib
import io
import re
import shutil
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

SUPPORTED = {".xlsx", ".xls", ".pdf", ".png", ".jpg", ".jpeg", ".hwpx"}


@dataclass
class Candidate:
    kind: str
    value: str
    file: Path
    location: str
    confidence: float
    selected: bool = True
    bbox: tuple[float, float, float, float] | None = None
    reviewed: bool = False
    manual: bool = False

    @property
    def safe_value(self) -> str:
        text = self.value
        if len(text) <= 2:
            return "*" * len(text)
        return text[0] + "*" * min(8, len(text) - 2) + text[-1]


@dataclass
class FileResult:
    source: Path
    output: Path | None = None
    status: str = "실패"
    processed: int = 0
    low_confidence: int = 0
    message: str = ""


PATTERNS = [
    # 뒷자리 첫 숫자 5~8은 외국인등록번호다. 유효성은 VALIDATORS에서 다시 확인한다.
    ("주민등록번호", re.compile(r"(?<!\d)(?:\d{6})[- ]?[1-8]\d{6}(?!\d)"), 0.99),
    ("생년월일", re.compile(r"(?:생\s*년\s*월\s*일|출\s*생\s*일)\s*[:：]?\s*((?:19|20)?\d{2}[.\-/ ]?\d{2}[.\-/ ]?\d{2}\.?)"), 0.92),
    ("이메일", re.compile(r"(?i)(?<![\w.])[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}(?![\w.])"), 0.98),
    # 휴대전화·지역번호(02, 031~064)·인터넷전화(070)·안심번호(050X)·수신자부담(080).
    ("전화번호", re.compile(r"(?<!\d)(?:01[016789]|0(?:2|3[1-3]|4[1-4]|5[1-5]|6[1-4]|50\d|70|80))[- .)]?\d{3,4}[- .]?\d{4}(?!\d)"), 0.95),
    ("전화번호", re.compile(r"(?<!\d)(?:15|16|18)\d{2}[- .]?\d{4}(?!\d)"), 0.8),
    ("카드번호", re.compile(r"(?<!\d)(?:\d[ -]?){15,16}(?!\d)"), 0.75),
    ("계좌번호", re.compile(r"(?:계\s*좌\s*번\s*호|통\s*장\s*번\s*호|계\s*좌)\s*[:：]?\s*((?:\d{2,6}[- ]?){2,4}\d{2,6})"), 0.9),
    ("IP 주소", re.compile(r"(?<!\d)(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)(?!\d)"), 0.9),
    ("여권번호", re.compile(r"(?i)(?:여\s*권\s*번\s*호|passport\s*(?:no\.?|number)?)\s*[:：]?\s*([A-Z]{1,2}\d{7,8})"), 0.9),
    ("운전면허번호", re.compile(r"(?<![\w-])(?:\d{2}|[가-힣]{2})[- ]\d{2}[- ]\d{6}[- ]\d{2}(?![\w-])"), 0.9),
    # 개인사업자 등록번호는 사업주 개인을 특정할 수 있어 체크섬이 맞을 때만 후보로 둔다.
    ("사업자등록번호", re.compile(r"(?<!\d)\d{3}[- ]\d{2}[- ]\d{5}(?!\d)"), 0.9),
    # 기관명(예: '○○시 ○○구청')을 주소로 오인하지 않도록 번지/건물 번호를 필수로 둔다.
    ("주소", re.compile(r"(?:[가-힣]+(?:특별시|광역시|특별자치시|도)\s*)?[가-힣]+(?:시|군|구)\s+[가-힣0-9·.-]+(?:로|길|동|읍|면|리)\s*\d+(?:-\d+)?"), 0.72),
    ("이름", re.compile(r"(?:성\s*명|이\s*름|신\s*청\s*인|담\s*당\s*자|대\s*표\s*자|예\s*금\s*주)\s*[:：]?\s*([가-힣]{2,6})"), 0.78),
]

# 정규식의 첫 번째 그룹만 개인정보 값인 유형(앞의 라벨은 마스킹 대상이 아니다).
GROUPED_KINDS = {"이름", "생년월일", "계좌번호", "여권번호"}

# PDF·이미지에 덧쓰는 글자는 ASCII 대체 코드를 쓴다. 기본 내장 글꼴이 한글을 그리지 못한다.
ALIAS_CODES = {
    "이름": "NAME", "주민등록번호": "RRN", "생년월일": "DOB", "전화번호": "TEL", "이메일": "EMAIL",
    "카드번호": "CARD", "계좌번호": "ACCOUNT", "IP 주소": "IP", "주소": "ADDR",
    "여권번호": "PASSPORT", "운전면허번호": "LICENSE", "사업자등록번호": "BIZNO",
    "사용자 지정": "CUSTOM",
}

TYPE_FILTERS = {
    "이름": "이름", "성명": "이름", "성함": "이름", "예금주": "이름", "신청인": "이름", "담당자": "이름", "대표자": "이름",
    "주민등록번호": "주민등록번호", "주민번호": "주민등록번호", "외국인등록번호": "주민등록번호",
    "생년월일": "생년월일", "출생일": "생년월일",
    "계좌번호": "계좌번호", "통장번호": "계좌번호", "계좌": "계좌번호",
    "전화번호": "전화번호", "연락처": "전화번호", "휴대전화": "전화번호", "휴대폰": "전화번호",
    "이메일": "이메일", "메일": "이메일",
    "카드번호": "카드번호", "카드": "카드번호",
    "IP주소": "IP 주소", "IP 주소": "IP 주소", "아이피주소": "IP 주소",
    "주소": "주소",
    "여권번호": "여권번호", "여권": "여권번호",
    "운전면허번호": "운전면허번호", "면허번호": "운전면허번호", "운전면허": "운전면허번호",
    "사업자등록번호": "사업자등록번호", "사업자번호": "사업자등록번호",
}

KOREAN_SURNAMES = set("김이박최정강조윤장임한오서신권황안송전홍유고문양손배백허남심노하곽성차주우구민진지엄채원천방공현함변염여추도소석선설마길연위표명기반왕금옥육인맹제모탁국어은편용")

# 두 글자 성은 첫 글자만으로는 걸러지지 않으므로 따로 둔다.
COMPOUND_SURNAMES = {"남궁", "황보", "제갈", "사공", "선우", "서문", "독고", "동방", "망절", "소봉", "어금", "장곡", "즙문", "강전"}

HEADER_KEYWORDS = {
    "이름": ("성명", "성함", "이름", "예금주", "신청인", "담당자", "대표자", "참가자", "참석자", "교육생", "수강생", "직원명", "교직원명"),
    "주민등록번호": ("주민등록번호", "주민번호", "외국인등록번호"),
    "생년월일": ("생년월일", "출생일"),
    "계좌번호": ("계좌번호", "통장번호", "입금계좌", "지급계좌"),
    "전화번호": ("전화번호", "연락처", "휴대전화", "휴대폰", "핸드폰", "모바일", "내선번호"),
    "이메일": ("이메일", "전자우편", "메일주소"),
    "카드번호": ("카드번호",),
    "IP 주소": ("IP주소", "아이피주소"),
    "주소": ("주소", "거주지", "소재지"),
    "여권번호": ("여권번호",),
    "운전면허번호": ("운전면허번호", "면허번호"),
    "사업자등록번호": ("사업자등록번호", "사업자번호"),
}


def _digits(value: str) -> str:
    return re.sub(r"\D", "", value)


def luhn_valid(number: str) -> bool:
    digits = [int(ch) for ch in _digits(number)]
    if len(digits) not in {15, 16}:
        return False
    total = 0
    for position, digit in enumerate(reversed(digits)):
        if position % 2:
            digit *= 2
            if digit > 9:
                digit -= 9
        total += digit
    return total % 10 == 0


def rrn_checksum_valid(number: str) -> bool:
    digits = _digits(number)
    if len(digits) != 13:
        return False
    weights = (2, 3, 4, 5, 6, 7, 8, 9, 2, 3, 4, 5)
    total = sum(int(digit) * weight for digit, weight in zip(digits, weights))
    return (11 - total % 11) % 10 == int(digits[12])


def business_number_valid(number: str) -> bool:
    digits = _digits(number)
    if len(digits) != 10:
        return False
    weights = (1, 3, 7, 1, 3, 7, 1, 3, 5)
    total = sum(int(digit) * weight for digit, weight in zip(digits, weights))
    total += int(digits[8]) * 5 // 10
    return (10 - total % 10) % 10 == int(digits[9])


def _validate_rrn(value: str, _text: str) -> float | None:
    digits = _digits(value)
    if len(digits) != 13:
        return None
    month, day = int(digits[2:4]), int(digits[4:6])
    if not (1 <= month <= 12 and 1 <= day <= 31):
        return None
    # 2020년 이후 발급된 외국인등록번호는 검증번호가 없어 체크섬이 맞지 않는다.
    return 0.99 if rrn_checksum_valid(digits) else 0.9


def _validate_card(value: str, text: str) -> float | None:
    if luhn_valid(value):
        return 0.95
    # 체크섬이 맞지 않는 15~16자리는 주문번호·수납번호일 가능성이 높다.
    return 0.6 if re.search(r"카\s*드|card", text, re.IGNORECASE) else None


def _validate_business_number(value: str, _text: str) -> float | None:
    return 0.9 if business_number_valid(value) else None


VALIDATORS = {
    "주민등록번호": _validate_rrn,
    "카드번호": _validate_card,
    "사업자등록번호": _validate_business_number,
}


def classify_header(value: object) -> set[str]:
    if not isinstance(value, str):
        return set()
    compact = re.sub(r"[\s:：()\[\]·._-]+", "", value)
    kinds = set()
    for kind, keywords in HEADER_KEYWORDS.items():
        if any(keyword in compact for keyword in keywords):
            kinds.add(kind)
    return kinds


def looks_like_name(value: object, require_common_surname: bool = True) -> bool:
    if not isinstance(value, str):
        return False
    compact = re.sub(r"\s+", "", value)
    if not re.fullmatch(r"[가-힣]{2,6}", compact):
        return False
    if not require_common_surname:
        return True
    return compact[0] in KOREAN_SURNAMES or compact[:2] in COMPOUND_SURNAMES


def looks_like_account(value: object) -> bool:
    if not isinstance(value, str):
        return False
    compact = re.sub(r"\s+", "", value)
    digits = re.sub(r"\D", "", compact)
    return 8 <= len(digits) <= 20 and bool(re.fullmatch(r"\d{2,6}(?:-\d{2,6}){1,4}", compact))


def looks_like_header_value(kind: str, value: str) -> bool:
    """명시적인 개인정보 헤더 아래 값인지 유형별로 검증한다."""
    text = value.strip()
    if not text:
        return False
    if kind == "이름":
        # 명시 헤더가 있으므로 희귀 성씨·복성·띄어 쓴 이름도 허용한다.
        return looks_like_name(text, require_common_surname=False)
    if kind == "계좌번호":
        return looks_like_account(text)
    if kind == "주소":
        return len(text) >= 4 and bool(re.search(r"[가-힣]", text))
    # 같은 유형에 정규식이 여러 개인 경우(예: 휴대전화와 대표번호) 하나라도 맞으면 값으로 본다.
    patterns = [pattern for pattern_kind, pattern, _confidence in PATTERNS if pattern_kind == kind]
    validator = VALIDATORS.get(kind)
    for pattern in patterns:
        match = pattern.search(text)
        if match and (validator is None or validator(match.group(0), text) is not None):
            return True
    return False


def custom_filter(custom_values: list[str]) -> tuple[set[str], list[str]]:
    """입력한 필드명은 유형 필터로, 그 밖의 문자열은 직접 마스킹 값으로 해석한다."""
    allowed: set[str] = set()
    literals: list[str] = []
    for raw in custom_values:
        value = raw.strip()
        if not value:
            continue
        normalized = re.sub(r"\s+", " ", value)
        compact = normalized.replace(" ", "")
        kind = TYPE_FILTERS.get(normalized) or TYPE_FILTERS.get(compact)
        if kind:
            allowed.add(kind)
        else:
            literals.append(value)
    return allowed, literals


def collect_files(paths: Iterable[str | Path]) -> list[Path]:
    found: set[Path] = set()
    for raw in paths:
        p = Path(raw)
        if p.is_file() and p.suffix.lower() in SUPPORTED and not p.name.startswith("~$"):
            found.add(p.resolve())
        elif p.is_dir():
            found.update(x.resolve() for x in p.rglob("*") if x.is_file() and x.suffix.lower() in SUPPORTED and not x.name.startswith("~$"))
    return sorted(found)


def unsupported_files(paths: Iterable[str | Path]) -> list[Path]:
    """지원하지 않아 조용히 빠지는 파일을 사용자에게 알리기 위해 따로 모은다."""
    found: set[Path] = set()
    for raw in paths:
        p = Path(raw)
        candidates = [p] if p.is_file() else (x for x in p.rglob("*") if x.is_file()) if p.is_dir() else []
        for item in candidates:
            if item.suffix.lower() not in SUPPORTED and not item.name.startswith("~$"):
                found.add(item.resolve())
    return sorted(found)


def flexible_pattern(value: str) -> re.Pattern:
    compact = re.sub(r"\s+", "", value)
    parts = [re.escape(ch) for ch in compact]
    return re.compile(r"\s*".join(parts), re.IGNORECASE)


def find_in_text(text: str, file: Path, location: str, custom_values: list[str]) -> list[Candidate]:
    out: list[Candidate] = []
    seen: set[tuple[str, str]] = set()
    allowed_kinds, literal_values = custom_filter(custom_values)
    for kind, pattern, confidence in PATTERNS:
        if custom_values and kind not in allowed_kinds:
            continue
        validator = VALIDATORS.get(kind)
        for match in pattern.finditer(text):
            value = match.group(1) if kind in GROUPED_KINDS else match.group(0)
            if kind == "이름" and not looks_like_name(value):
                continue
            score = confidence
            if validator is not None:
                score = validator(value, text)
                if score is None:
                    continue
            key = (kind, re.sub(r"\s+", "", value).casefold())
            if key not in seen:
                seen.add(key)
                out.append(Candidate(kind, value, file, location, score))
    for value in literal_values:
        value = value.strip()
        if not value:
            continue
        for match in flexible_pattern(value).finditer(text):
            actual = match.group(0)
            key = ("사용자 지정", re.sub(r"\s+", "", actual).casefold())
            if key not in seen:
                seen.add(key)
                out.append(Candidate("사용자 지정", actual, file, location, 1.0))
    return out


def scan_file(path: Path, custom_values: list[str]) -> tuple[list[Candidate], list[str]]:
    ext = path.suffix.lower()
    warnings: list[str] = []
    if ext in {".xlsx", ".xls"}:
        return _scan_excel(path, custom_values), warnings
    if ext == ".pdf":
        return _scan_pdf(path, custom_values)
    if ext in {".png", ".jpg", ".jpeg"}:
        return _scan_image(path, custom_values)
    if ext == ".hwpx":
        return _scan_hwpx(path, custom_values), warnings
    return [], ["지원하지 않는 형식"]


def _scan_excel(path: Path, custom: list[str]) -> list[Candidate]:
    sheet_cells: list[tuple[str, list[tuple[int, int, str, object]]]] = []
    if path.suffix.lower() == ".xlsx":
        from openpyxl import load_workbook
        from openpyxl.utils import get_column_letter
        book = load_workbook(path, data_only=False, read_only=True)
        for ws in book.worksheets:
            cells = []
            for row_index, row in enumerate(ws.iter_rows(), 1):
                for column_index, cell in enumerate(row, 1):
                    coordinate = f"{get_column_letter(column_index)}{row_index}"
                    cells.append((row_index, column_index, coordinate, cell.value))
            sheet_cells.append((ws.title, cells))
    else:
        import xlrd
        book = xlrd.open_workbook(path, on_demand=True)
        for sheet in book.sheets():
            cells = []
            for r in range(sheet.nrows):
                for c in range(sheet.ncols):
                    cells.append((r + 1, c + 1, f"R{r + 1}C{c + 1}", sheet.cell_value(r, c)))
            sheet_cells.append((sheet.name, cells))
    out: list[Candidate] = []
    seen_keys: set[tuple[str, str, str]] = set()
    allowed_kinds, _literal_values = custom_filter(custom)
    for sheet, cells in sheet_cells:
        headers_by_column: dict[int, list[tuple[int, str, set[str]]]] = {}
        cells_by_row: dict[int, list[tuple[int, object]]] = {}
        # classify_header는 셀마다 한 번만 호출한다(열 단위로 다시 훑으면 셀 수 × 열 수가 된다).
        header_kinds_by_cell: dict[tuple[int, int], set[str]] = {}
        values_by_column: dict[int, list[str]] = {}
        for row_index, column_index, _coordinate, value in cells:
            kinds = classify_header(value)
            header_kinds_by_cell[(row_index, column_index)] = kinds
            if isinstance(value, str) and not kinds:
                values_by_column.setdefault(column_index, []).append(value)
            if value not in (None, ""):
                cells_by_row.setdefault(row_index, []).append((column_index, value))
        header_rows = {
            row_index
            for row_index, row_values in cells_by_row.items()
            if len([value for _column, value in row_values if isinstance(value, str)]) >= 2
            and any(header_kinds_by_cell.get((row_index, column), set()) for column, _value in row_values)
        }
        explicit_header_columns = {
            column_index
            for row_index, row_values in cells_by_row.items() if row_index in header_rows
            for column_index, _value in row_values
        }
        for row_index, column_index, _coordinate, value in cells:
            kinds = header_kinds_by_cell.get((row_index, column_index), set())
            if kinds:
                headers_by_column.setdefault(column_index, []).append((row_index, str(value).strip(), kinds))

        # 헤더가 낯선 표도 값 분포로 이름 열을 추론한다. 반복 직위 열은 다양성 조건에서 제외된다.
        inferred_name_columns: set[int] = set()
        for column_index, values in values_by_column.items():
            if column_index in explicit_header_columns or len(values) < 5:
                continue
            name_values = [value for value in values if looks_like_name(value)]
            unique_names = {re.sub(r"\s+", "", value) for value in name_values}
            if len(name_values) / len(values) >= 0.65 and len(unique_names) >= 4:
                inferred_name_columns.add(column_index)

        for row_index, column_index, coordinate, value in cells:
            if not isinstance(value, str) or not value.strip():
                continue
            location = f"{sheet}!{coordinate}"
            for candidate in find_in_text(value, path, location, custom):
                key = (candidate.kind, re.sub(r"\s+", "", candidate.value).casefold(), location)
                if key not in seen_keys:
                    seen_keys.add(key)
                    out.append(candidate)
            if row_index in header_rows or header_kinds_by_cell.get((row_index, column_index)):
                continue
            preceding = [item for item in headers_by_column.get(column_index, []) if item[0] < row_index]
            header_kinds = max(preceding, key=lambda item: item[0])[2] if preceding else set()
            if column_index in inferred_name_columns:
                header_kinds = header_kinds | {"이름"}
            for kind in header_kinds:
                if custom and kind not in allowed_kinds:
                    continue
                if not looks_like_header_value(kind, value):
                    continue
                candidate = Candidate(kind, value.strip(), path, location, 0.94)
                key = (candidate.kind, re.sub(r"\s+", "", candidate.value).casefold(), location)
                if key not in seen_keys:
                    seen_keys.add(key)
                    out.append(candidate)
    if hasattr(book, "close"):
        book.close()
    elif hasattr(book, "release_resources"):
        book.release_resources()
    return out


def _scan_pdf(path: Path, custom: list[str]) -> tuple[list[Candidate], list[str]]:
    import fitz
    doc = fitz.open(path)
    out: list[Candidate] = []
    warnings: list[str] = []
    for i, page in enumerate(doc):
        text = page.get_text("text")
        has_images = bool(page.get_images(full=True))
        if text.strip():
            out.extend(find_in_text(text, path, f"{i + 1}페이지", custom))
            # PDF 내부 텍스트는 읽기 순서가 열/블록 단위로 뒤섞일 수 있다. 실제 좌표의
            # 같은 행을 다시 조립하면 '성   명 : 홍은성' 같은 필드를 복원할 수 있다.
            for line in _pdf_spatial_lines(page):
                out.extend(find_in_text(line, path, f"{i + 1}페이지", custom))
        # 텍스트가 조금이라도 있는 혼합형 PDF에도 스캔 이미지가 함께 들어갈 수 있다.
        # 따라서 모든 페이지에 OCR을 병행한다.
        try:
            allowed_kinds, _literal_values = custom_filter(custom)
            name_requested = not custom or "이름" in allowed_kinds
            data, _scale = _ocr_pdf_page(page, config="--psm 11" if name_requested else "--psm 6")
            ocr_text = " ".join(x for x in data["text"] if x.strip())
            scores = [float(x) for x in data["conf"] if str(x) not in {"-1", ""} and float(x) >= 0]
            ocr_confidence = sum(scores) / len(scores) / 100 if scores else 0
            found = find_in_text(ocr_text, path, f"{i + 1}페이지(OCR)", custom)
            if not custom or "생년월일" in allowed_kinds:
                found.extend(_ocr_birth_date_candidates(data, path, f"{i + 1}페이지(OCR)"))
            if name_requested:
                found = [candidate for candidate in found if candidate.kind != "이름"]
                extra_data, _extra_scale = _ocr_pdf_page(page, config="--psm 6")
                for name_data in (data, extra_data):
                    found.extend(_ocr_name_candidates(name_data, path, f"{i + 1}페이지(OCR)"))
            for candidate in found:
                candidate.confidence = min(candidate.confidence, ocr_confidence)
            out.extend(found)
            if has_images and ocr_confidence < 0.65:
                warnings.append(f"{i + 1}페이지: OCR 신뢰도 부족({ocr_confidence:.0%})")
        except Exception as exc:
            if has_images:
                warnings.append(f"{i + 1}페이지: OCR 실패({exc})")
    doc.close()
    # 페이지 텍스트와 좌표 재조립 결과가 같은 값을 두 번 찾아내는 경우가 많다.
    unique: dict[tuple[str, str, int], Candidate] = {}
    for candidate in out:
        page_match = re.match(r"(\d+)페이지", candidate.location)
        page_no = int(page_match.group(1)) if page_match else 0
        key = (candidate.kind, re.sub(r"\s+", "", candidate.value).casefold(), page_no)
        # 같은 페이지의 텍스트층과 OCR에서 같은 값이 잡히면 한 행만 보여준다.
        # 처리 단계에서는 어느 출처의 후보든 텍스트와 OCR 좌표를 모두 검색한다.
        if key not in unique or (
            "(OCR)" in unique[key].location and "(OCR)" not in candidate.location
        ) or candidate.confidence > unique[key].confidence:
            unique[key] = candidate
    return list(unique.values()), warnings


def _pdf_spatial_lines(page) -> list[str]:
    words = sorted(page.get_text("words"), key=lambda word: ((word[1] + word[3]) / 2, word[0]))
    lines: list[list[tuple]] = []
    for word in words:
        center = (word[1] + word[3]) / 2
        height = max(1, word[3] - word[1])
        line = next((items for items in reversed(lines[-8:]) if abs(((items[0][1] + items[0][3]) / 2) - center) <= height * 0.55), None)
        if line is None:
            line = []
            lines.append(line)
        line.append(word)
    return [" ".join(str(word[4]) for word in sorted(line, key=lambda word: word[0])) for line in lines]


def _ocr_data(path_or_image, config: str = ""):
    import pytesseract
    from pytesseract import Output
    if not shutil.which("tesseract"):
        for executable in (
            Path("C:/Program Files/Tesseract-OCR/tesseract.exe"),
            Path("C:/Program Files (x86)/Tesseract-OCR/tesseract.exe"),
        ):
            if executable.exists():
                pytesseract.pytesseract.tesseract_cmd = str(executable)
                break
    return pytesseract.image_to_data(path_or_image, lang="kor+eng", config=config, output_type=Output.DICT)


def _ocr_pdf_page(page, scale: float = 3.0, config: str = "--psm 6"):
    """PDF 페이지를 고해상도 이미지로 렌더링하고 OCR 좌표와 배율을 반환한다."""
    import fitz
    from PIL import Image
    pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
    return _ocr_data(image, config=config), scale


def _ocr_name_candidates(data: dict, path: Path, location: str) -> list[Candidate]:
    """라벨의 오른쪽·아래 셀 또는 숫자 식별자 뒤의 한국 이름을 좌표로 찾는다."""
    excluded = {
        "성명", "이름", "예금주", "담당자", "대표자", "신청인", "연락처", "휴대전화", "전화번호",
        "생년월일", "이메일", "주소", "소속기관", "부서", "직위", "호봉", "은행명", "계좌번호",
        "학위", "전공", "근무처", "기간", "서명", "성별", "국적",
    }
    out: list[Candidate] = []
    seen: set[str] = set()
    lines: dict[tuple[int, int, int], list[int]] = {}
    for i, raw in enumerate(data["text"]):
        if raw.strip():
            key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
            lines.setdefault(key, []).append(i)
    all_runs = []
    for line_key, indexes in lines.items():
        runs: list[list[int]] = []
        for i in indexes:
            token = re.sub(r"\s+", "", data["text"][i])
            if not re.fullmatch(r"[가-힣]+", token):
                continue
            if runs:
                prev = runs[-1][-1]
                gap = data["left"][i] - (data["left"][prev] + data["width"][prev])
                height = max(data["height"][i], data["height"][prev])
                if gap <= height * 1.5:
                    runs[-1].append(i)
                    continue
            runs.append([i])
        for run in runs:
            value = "".join(re.sub(r"\s+", "", data["text"][i]) for i in run)
            all_runs.append({
                "line": line_key, "indexes": run, "value": value,
                "left": min(data["left"][i] for i in run),
                "top": min(data["top"][i] for i in run),
                "right": max(data["left"][i] + data["width"][i] for i in run),
                "bottom": max(data["top"][i] + data["height"][i] for i in run),
            })
    eligible = [
        run for run in all_runs
        if run["value"] not in excluded
        and re.fullmatch(r"[가-힣]{2,4}", run["value"])
        and (run["value"][0] in KOREAN_SURNAMES or run["value"][:2] in COMPOUND_SURNAMES)
    ]
    selected_runs = []
    # 계좌번호·주민번호처럼 긴 숫자 식별자 바로 뒤의 예금주/신청인 이름.
    for line_key, indexes in lines.items():
        numeric_right = max(
            (
                data["left"][i] + data["width"][i]
                for i in indexes
                if len(re.sub(r"\D", "", data["text"][i])) >= 8
            ),
            default=None,
        )
        if numeric_right is not None:
            choices = [run for run in eligible if run["line"] == line_key and run["left"] > numeric_right]
            if choices:
                selected_runs.append(min(choices, key=lambda run: run["left"] - numeric_right))
    # 이름 계열 라벨마다 가장 가까운 오른쪽 또는 바로 아래 값 하나만 선택한다.
    name_labels = {"성명", "이름", "예금주", "담당자", "대표자", "신청인"}
    for label in (run for run in all_runs if run["value"] in name_labels):
        label_center = (label["left"] + label["right"]) / 2
        choices = []
        for run in eligible:
            run_center = (run["left"] + run["right"]) / 2
            if run["line"] == label["line"] and run["left"] > label["right"]:
                choices.append((run["left"] - label["right"], run))
                continue
            vertical_gap = run["top"] - label["bottom"]
            if 0 <= vertical_gap <= max(300, (label["bottom"] - label["top"]) * 8) and abs(run_center - label_center) <= 180:
                choices.append((vertical_gap * 3 + abs(run_center - label_center), run))
        if choices:
            selected_runs.append(min(choices, key=lambda item: item[0])[1])
    for run in selected_runs:
        value = run["value"]
        scores = [float(data["conf"][i]) for i in run["indexes"] if str(data["conf"][i]) not in {"-1", ""}]
        confidence = sum(scores) / len(scores) / 100 if scores else 0.6
        if confidence < 0.35 or value in seen:
            continue
        seen.add(value)
        out.append(Candidate("이름", value, path, location, min(0.78, confidence)))
    return out


def _ocr_birth_date_candidates(data: dict, path: Path, location: str) -> list[Candidate]:
    """표 OCR에서 생년월일 라벨과 아래 값이 서로 다른 블록으로 갈라진 경우를 복원한다."""
    valid = [i for i, raw in enumerate(data["text"]) if raw.strip()]
    out = []
    for label in (i for i in valid if re.sub(r"\s+", "", data["text"][i]) in {"생년월일", "출생일"}):
        label_center = data["left"][label] + data["width"][label] / 2
        label_bottom = data["top"][label] + data["height"][label]
        choices = []
        for i in valid:
            raw = data["text"][i].strip()
            digits = re.sub(r"\D", "", raw)
            if len(digits) not in {6, 8}:
                continue
            try:
                fmt = "%Y%m%d" if len(digits) == 8 else "%y%m%d"
                datetime.strptime(digits, fmt)
            except ValueError:
                continue
            center = data["left"][i] + data["width"][i] / 2
            vertical_gap = data["top"][i] - label_bottom
            if 0 <= vertical_gap <= data["height"][label] * 10 and abs(center - label_center) <= data["width"][label] * 1.5:
                choices.append((vertical_gap + abs(center - label_center), i))
        if choices:
            _distance, index = min(choices)
            confidence = float(data["conf"][index]) / 100 if str(data["conf"][index]) not in {"-1", ""} else 0.6
            out.append(Candidate("생년월일", data["text"][index].strip(), path, location, min(0.92, max(0.5, confidence))))
    return out


def _scan_image(path: Path, custom: list[str]) -> tuple[list[Candidate], list[str]]:
    from PIL import Image, ImageEnhance, ImageOps
    try:
        original = Image.open(path).convert("RGB")
        enlarged = original.resize((original.width * 2, original.height * 2), Image.Resampling.LANCZOS)
        gray = ImageEnhance.Contrast(ImageOps.grayscale(enlarged)).enhance(2.0)
        passes = [
            (_ocr_data(original, config="--psm 6"), 1.0),
            (_ocr_data(enlarged, config="--psm 11"), 2.0),
            (_ocr_data(gray, config="--psm 11"), 2.0),
        ]
    except Exception as exc:
        return [], [f"OCR 실행 실패: {exc}"]
    out: list[Candidate] = []
    seen_keys: set[tuple] = set()
    allowed_kinds, _literal_values = custom_filter(custom)
    confidences = []
    for data, _scale in passes:
        text = " ".join(x for x in data["text"] if x.strip())
        scores = [float(x) for x in data["conf"] if str(x) not in {"-1", ""} and float(x) >= 0]
        confidence = sum(scores) / len(scores) / 100 if scores else 0
        confidences.append(confidence)
        found = find_in_text(text, path, "이미지 전체", custom)
        found.extend(_ocr_labeled_candidates(data, path))
        found.extend(_ocr_honorific_candidates(data, path))
        if not custom or "이름" in allowed_kinds:
            found.extend(_ocr_bankbook_name_regions(data, path, _scale))
        if not custom or "계좌번호" in allowed_kinds:
            found.extend(_ocr_account_candidates(data, path, require_label=not custom))
        for candidate in found:
            if custom and candidate.kind not in allowed_kinds and candidate.kind != "사용자 지정":
                continue
            candidate.confidence = min(candidate.confidence, max(0.35, confidence))
            # 좌표 기반 후보는 값이 같아도 위치가 다르면 별개다(통장 여러 장 등).
            box = tuple(round(v) for v in candidate.bbox) if candidate.bbox else None
            key = (candidate.kind, re.sub(r"\s+", "", candidate.value).casefold(), box)
            if key not in seen_keys:
                seen_keys.add(key)
                out.append(candidate)
    warnings = ["OCR 평균 신뢰도 부족(다중 전처리 적용)"] if max(confidences, default=0) < 0.65 else []
    return out, warnings


def _ocr_bankbook_name_regions(data: dict, path: Path, scale: float) -> list[Candidate]:
    """통장 계좌번호 위의 예금주 영역을 글자 판독 실패 시에도 위치 후보로 제공한다."""
    out = []
    for i, raw in enumerate(data["text"]):
        normalized = re.sub(r"[Oo]", "0", raw.strip())
        digits = re.sub(r"\D", "", normalized)
        if len(digits) < 10 or normalized.count("-") < 2:
            continue
        x = data["left"][i] / scale
        y = data["top"][i] / scale
        width = data["width"][i] / scale
        height = max(8, data["height"][i] / scale)
        bbox = (x, max(0, y - height * 4.2), x + width * 1.1, max(0, y - height * 0.35))
        out.append(Candidate("이름", "이름 영역(OCR)", path, "이미지 예금주 추정 영역", 0.55, bbox=bbox))
    return out


def _ocr_honorific_candidates(data: dict, path: Path) -> list[Candidate]:
    """OCR 행 구분이 깨져도 '님' 왼쪽의 가장 가까운 텍스트를 이름 영역으로 잡는다."""
    out = []
    valid = [i for i, raw in enumerate(data["text"]) if raw.strip()]
    for suffix in (i for i in valid if "님" in re.sub(r"\s+", "", data["text"][i])):
        center_y = data["top"][suffix] + data["height"][suffix] / 2
        choices = [
            i for i in valid
            if data["left"][i] + data["width"][i] <= data["left"][suffix]
            and abs((data["top"][i] + data["height"][i] / 2) - center_y) <= max(data["height"][suffix], data["height"][i]) * 1.5
            and not re.search(r"\d", data["text"][i])
        ]
        if not choices:
            continue
        nearest = max(choices, key=lambda i: data["left"][i] + data["width"][i])
        confidence = max(0.35, float(data["conf"][nearest]) / 100) if str(data["conf"][nearest]) not in {"-1", ""} else 0.45
        out.append(Candidate("이름", data["text"][nearest].strip() or "이름영역", path, "이미지 전체", min(0.65, confidence)))
    return out


def _ocr_account_candidates(data: dict, path: Path, require_label: bool) -> list[Candidate]:
    """공백·하이픈으로 분리되거나 0이 O로 읽힌 계좌번호를 행 단위로 복원한다."""
    lines: dict[tuple[int, int, int], list[int]] = {}
    for i, raw in enumerate(data["text"]):
        if raw.strip():
            lines.setdefault((data["block_num"][i], data["par_num"][i], data["line_num"][i]), []).append(i)
    out = []
    for indexes in lines.values():
        line = " ".join(data["text"][i] for i in indexes)
        compact_label = re.sub(r"\s+", "", line)
        has_label = any(label in compact_label for label in ("계좌번호", "통장번호", "계좌"))
        if require_label and not has_label:
            continue
        for match in re.finditer(r"(?<![0-9A-Za-z])(?:[0-9Oo]{2,6}[- ]+){2,4}[0-9Oo]{2,6}(?![0-9A-Za-z])", line):
            value = match.group(0).strip()
            normalized = re.sub(r"[Oo]", "0", value)
            digits = re.sub(r"\D", "", normalized)
            segments = re.split(r"[- ]+", normalized)
            if not (8 <= len(digits) <= 20 and (has_label or len(segments) >= 4)):
                continue
            scores = [float(data["conf"][i]) for i in indexes if str(data["conf"][i]) not in {"-1", ""} and float(data["conf"][i]) >= 0]
            confidence = sum(scores) / len(scores) / 100 if scores else 0.5
            out.append(Candidate("계좌번호", value, path, "이미지 전체", min(0.9, confidence)))
    return out


def _ocr_image(path_or_image, scale: float = 1.0):
    """문서형 이미지의 행 구조를 유지해 이름·계좌 필드를 OCR한다."""
    from PIL import Image
    image = path_or_image if isinstance(path_or_image, Image.Image) else Image.open(path_or_image)
    if scale != 1:
        image = image.resize((round(image.width * scale), round(image.height * scale)), Image.Resampling.LANCZOS)
    return _ocr_data(image, config="--psm 6"), scale


def _ocr_labeled_candidates(data: dict, path: Path) -> list[Candidate]:
    """계좌번호가 일부 오인식돼도 라벨 오른쪽 필드 전체를 마스킹 후보로 삼는다."""
    lines: dict[tuple[int, int, int], list[int]] = {}
    for i, raw in enumerate(data["text"]):
        if raw.strip():
            key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
            lines.setdefault(key, []).append(i)
    out: list[Candidate] = []
    labels = (("이름", {"예금주", "성명", "이름"}), ("계좌번호", {"계좌번호", "통장번호"}))
    for indexes in lines.values():
        words = [data["text"][i].strip() for i in indexes]
        compact = "".join(re.sub(r"\s+", "", word) for word in words)
        for kind, names in labels:
            label = next((name for name in names if name in compact), None)
            if label is None:
                continue
            value = compact.split(label, 1)[1].strip(" :：")
            if not value:
                continue
            if kind == "이름":
                match = re.search(r"[가-힣]{2,4}", value)
                value = match.group(0) if match else value[:20]
            scores = [float(data["conf"][i]) for i in indexes if str(data["conf"][i]) not in {"-1", ""}]
            confidence = sum(scores) / len(scores) / 100 if scores else 0.5
            out.append(Candidate(kind, value, path, "이미지 전체", confidence))
        # 저화질 한글 이름을 오독하더라도 존칭 '님'의 왼쪽 값 영역은 안전하게 제거한다.
        suffix_indexes = [i for i in indexes if "님" in re.sub(r"\s+", "", data["text"][i])]
        if suffix_indexes:
            suffix = suffix_indexes[0]
            before = [i for i in indexes if data["left"][i] + data["width"][i] <= data["left"][suffix] and not re.search(r"\d", data["text"][i])]
            if before:
                nearest = max(before, key=lambda i: data["left"][i] + data["width"][i])
                raw_value = data["text"][nearest].strip() or "이름영역"
                scores = [float(data["conf"][nearest])] if str(data["conf"][nearest]) not in {"-1", ""} else [45.0]
                out.append(Candidate("이름", raw_value, path, "이미지 전체", min(0.65, scores[0] / 100)))
    return out


def _hwpx_text(path: Path) -> str:
    from lxml import etree
    chunks: list[str] = []
    with zipfile.ZipFile(path) as zf:
        names = sorted(n for n in zf.namelist() if n.startswith("Contents/section") and n.endswith(".xml"))
        for name in names:
            root = etree.fromstring(zf.read(name))
            chunks.extend(root.xpath("//*[local-name()='t']/text()"))
            chunks.append("\n")
    return " ".join(chunks)


def _scan_hwpx(path: Path, custom: list[str]) -> list[Candidate]:
    from lxml import etree
    out: list[Candidate] = []
    allowed_kinds, _literal_values = custom_filter(custom)
    with zipfile.ZipFile(path) as zf:
        names = sorted(n for n in zf.namelist() if n.startswith("Contents/section") and n.endswith(".xml"))
        for section_name in names:
            root = etree.fromstring(zf.read(section_name))
            for table_index, table in enumerate(root.xpath("//*[local-name()='tbl']"), 1):
                rows = []
                for row in table.xpath(".//*[local-name()='tr']"):
                    cells = ["".join(cell.xpath(".//*[local-name()='t']/text()")) for cell in row.xpath("./*[local-name()='tc']")]
                    if cells:
                        rows.append(cells)
                header_kinds: dict[int, set[str]] = {}
                header_row = -1
                for row_index, row in enumerate(rows):
                    kinds = {column: classify_header(value) for column, value in enumerate(row)}
                    if len(row) >= 2 and sum(bool(value) for value in kinds.values()) >= 1:
                        header_kinds = kinds
                        header_row = row_index
                        break
                for row_index, row in enumerate(rows[header_row + 1:], header_row + 1):
                    for column, value in enumerate(row):
                        for kind in header_kinds.get(column, set()):
                            if custom and kind not in allowed_kinds:
                                continue
                            if looks_like_header_value(kind, value):
                                out.append(Candidate(kind, value.strip(), path, f"{section_name} 표{table_index} 행{row_index + 1} 열{column + 1}", 0.96))
    # 표가 아닌 본문과 사용자가 입력한 실제 값도 기존 정규식으로 검색한다.
    out.extend(find_in_text(_hwpx_text(path), path, "문서 전체", custom))
    unique = {}
    for candidate in out:
        key = (candidate.kind, re.sub(r"\s+", "", candidate.value).casefold(), candidate.location)
        unique[key] = candidate
    return list(unique.values())


def replacement_for(candidate: Candidate, mode: str, aliases: dict[tuple[str, str], int], ascii_only: bool = False) -> str:
    """같은 값에는 항상 같은 가명을 준다. ascii_only는 한글 글꼴이 없는 출력용이다."""
    if mode == "delete":
        return "■" * len(candidate.value)
    if mode in {"label_ko", "label_en"}:
        label = candidate.kind if mode == "label_ko" else ALIAS_CODES.get(candidate.kind, "PII")
        return f"[{label}]"
    key = (candidate.kind, re.sub(r"\s+", "", candidate.value).casefold())
    if key not in aliases:
        aliases[key] = sum(1 for k in aliases if k[0] == candidate.kind) + 1
    label = ALIAS_CODES.get(candidate.kind, "PII") if ascii_only else candidate.kind
    return f"[{label}_{aliases[key]:04d}]"


def process_file(path: Path, candidates: list[Candidate], mode: str, output_dir: Path, aliases: dict) -> FileResult:
    chosen = [c for c in candidates if c.selected and c.file == path]
    result = FileResult(path, processed=len(chosen), low_confidence=sum(c.confidence < 0.7 for c in chosen))
    try:
        ext = path.suffix.lower()
        if ext in {".xlsx", ".xls"}:
            result.output = _process_excel(path, chosen, mode, output_dir, aliases)
        elif ext == ".pdf":
            result.output = _process_pdf(path, chosen, mode, output_dir, aliases)
        elif ext in {".png", ".jpg", ".jpeg"}:
            result.output = _process_image(path, chosen, mode, output_dir, aliases)
        elif ext == ".hwpx":
            result.output = _process_hwpx(path, chosen, mode, output_dir, aliases)
        result.status = "성공"
        result.message = "처리 완료" if chosen else "선택된 개인정보 없음(복사본 생성)"
    except Exception as exc:
        result.message = str(exc)
    return result


def _replace_text(text: str, chosen: list[Candidate], mode: str, aliases: dict) -> str:
    value = text
    for c in sorted(chosen, key=lambda x: len(x.value), reverse=True):
        value = flexible_pattern(c.value).sub(lambda _: replacement_for(c, mode, aliases), value)
    return value


def _unique_output(output_dir: Path, stem: str, suffix: str) -> Path:
    candidate = output_dir / f"{stem}_비식별{suffix}"
    i = 2
    while candidate.exists():
        candidate = output_dir / f"{stem}_비식별_{i}{suffix}"
        i += 1
    return candidate


def _safe_output_stem(path: Path) -> str:
    """원본 파일명에 포함된 이름 등 개인정보가 결과명으로 복사되지 않게 한다."""
    return f"비식별문서_{sha256(path)[:12]}"


def _process_excel(path: Path, chosen: list[Candidate], mode: str, out_dir: Path, aliases: dict) -> Path:
    from openpyxl import Workbook, load_workbook
    output = _unique_output(out_dir, _safe_output_stem(path), ".xlsx")
    if path.suffix.lower() == ".xlsx":
        book = load_workbook(path, data_only=False)
        for ws in book.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str):
                        local = [c for c in chosen if c.location == f"{ws.title}!{cell.coordinate}"]
                        cell.value = _replace_text(cell.value, local, mode, aliases)
                    # 셀 메모에는 검토자 이름과 연락처가 자주 들어가지만 셀 값 스캔에는 잡히지 않는다.
                    if cell.comment is not None:
                        cell.comment.text = _replace_text(cell.comment.text or "", chosen, mode, aliases)
                        cell.comment.author = ""
        # 작성자, 회사 등 문서 속성도 파일 전체 처리 범위에 포함한다.
        book.properties.creator = ""
        book.properties.lastModifiedBy = ""
        book.properties.title = ""
        book.properties.subject = ""
        book.properties.description = ""
        book.properties.keywords = ""
        book.properties.category = ""
        book.save(output)
    else:
        import xlrd
        old = xlrd.open_workbook(path)
        book = Workbook()
        book.remove(book.active)
        for sheet in old.sheets():
            ws = book.create_sheet(sheet.name[:31])
            for r in range(sheet.nrows):
                for cidx in range(sheet.ncols):
                    value = sheet.cell_value(r, cidx)
                    if isinstance(value, str):
                        local = [c for c in chosen if c.location == f"{sheet.name}!R{r + 1}C{cidx + 1}"]
                        value = _replace_text(value, local, mode, aliases)
                    ws.cell(r + 1, cidx + 1, value)
        book.save(output)
    _verify_excel_output(output, chosen)
    return output


def _verify_excel_output(output: Path, chosen: list[Candidate]) -> None:
    """저장된 결과 셀에 선택한 원문이 남아 있으면 성공으로 보고하지 않는다."""
    if not chosen:
        return
    from openpyxl import load_workbook
    book = load_workbook(output, data_only=False, read_only=True)
    remaining: list[str] = []
    try:
        for candidate in chosen:
            if "!" not in candidate.location:
                continue
            sheet_name, coordinate = candidate.location.rsplit("!", 1)
            if sheet_name not in book.sheetnames:
                remaining.append(candidate.location)
                continue
            if coordinate.startswith("R") and "C" in coordinate:
                match = re.fullmatch(r"R(\d+)C(\d+)", coordinate)
                if not match:
                    remaining.append(candidate.location)
                    continue
                value = book[sheet_name].cell(int(match.group(1)), int(match.group(2))).value
            else:
                value = book[sheet_name][coordinate].value
            if isinstance(value, str) and flexible_pattern(candidate.value).search(value):
                remaining.append(candidate.location)
    finally:
        book.close()
    if remaining:
        try:
            output.unlink()
        except OSError:
            pass
        preview = ", ".join(remaining[:5])
        raise RuntimeError(f"저장 후 검증 실패: 원문이 남은 셀 {preview}")


# PyMuPDF 기본 글꼴(Helvetica)은 한글을 '?'로 바꾼다. 내장 CJK 글꼴을 써야 가명이 읽힌다.
PDF_ALIAS_FONT = "korea-s"


def _add_redaction(page, rect, candidate: Candidate, mode: str, aliases: dict) -> None:
    import fitz
    if mode == "delete":
        page.add_redact_annot(rect, fill=(0, 0, 0))
        return
    text = replacement_for(candidate, mode, aliases)
    # 가명은 원문보다 길어서 칸 폭에 맞추지 않으면 줄이 잘려 읽을 수 없다.
    size = max(6, min(11, rect.height * 0.7))
    try:
        width = fitz.get_text_length(text, fontname=PDF_ALIAS_FONT, fontsize=size)
        if width > rect.width > 0:
            size = max(4, size * rect.width / width)
    except Exception:
        pass
    page.add_redact_annot(rect, text=text, fill=(1, 1, 1), text_color=(0, 0, 0), fontname=PDF_ALIAS_FONT, fontsize=size)


def _pdf_word_rects(page, value: str) -> list:
    """search_for가 놓치는, 공백·줄바꿈으로 쪼개진 값을 단어 좌표로 다시 찾는다."""
    import fitz
    target = re.sub(r"\s+", "", value).casefold()
    if not target:
        return []
    tokens = []
    compact = ""
    for word in page.get_text("words"):
        token = re.sub(r"\s+", "", str(word[4])).casefold()
        if not token:
            continue
        start = len(compact)
        compact += token
        tokens.append((start, len(compact), word))
    rects = []
    offset = 0
    while True:
        start = compact.find(target, offset)
        if start < 0:
            break
        end = start + len(target)
        hits = [word for a, b, word in tokens if a < end and b > start]
        offset = end
        if not hits:
            continue
        # 서로 다른 줄의 단어가 우연히 이어 붙은 경우까지 지우면 본문이 통째로 사라진다.
        heights = [word[3] - word[1] for word in hits]
        centers = [(word[1] + word[3]) / 2 for word in hits]
        if max(centers) - min(centers) > max(heights) * 0.6:
            continue
        rect = fitz.Rect(hits[0][:4])
        for word in hits[1:]:
            rect |= fitz.Rect(word[:4])
        rects.append(rect)
    return rects


def _process_pdf(path: Path, chosen: list[Candidate], mode: str, out_dir: Path, aliases: dict) -> Path:
    import fitz
    output = _unique_output(out_dir, _safe_output_stem(path), ".pdf")
    doc = fitz.open(path)
    matched_candidates: set[int] = set()
    for page_no, page in enumerate(doc, 1):
        page_local = [c for c in chosen if c.location in {f"{page_no}페이지", f"{page_no}페이지(OCR)"}]
        for c in page_local:
            rects = [fitz.Rect(c.bbox)] if c.manual and c.bbox else (page.search_for(c.value) or _pdf_word_rects(page, c.value))
            for rect in rects:
                _add_redaction(page, rect, c, mode, aliases)
                matched_candidates.add(id(c))
        if page_local and any(not c.manual for c in page_local):
            matches = []
            seen_rects = set()
            for config in ("--psm 6", "--psm 11"):
                data, scale = _ocr_pdf_page(page, config=config)
                for c, rect in _ocr_candidate_rects(data, [item for item in page_local if not item.manual], scale):
                    key = tuple(round(v, 1) for v in (rect.x0, rect.y0, rect.x1, rect.y1))
                    if key not in seen_rects:
                        seen_rects.add(key)
                        matches.append((c, rect))
            page_image, image_scale = _render_pdf_page_image(page, 2.0)
            expanded_seen = set()
            for c, rect in matches:
                # 주소 OCR은 행정구역 앞부분을 자주 누락하므로 셀 안의 전체 글자까지 확장한다.
                # 이름·전화번호·계좌번호는 같은 셀에 다른 필드가 있을 수 있어 정확 일치 범위만 지운다.
                if c.kind == "주소":
                    rect = _expand_to_table_cell_content(page_image, rect, image_scale)
                rect = _pad_ocr_rect(rect, page.rect)
                expanded_key = tuple(round(v, 1) for v in (rect.x0, rect.y0, rect.x1, rect.y1))
                if expanded_key in expanded_seen:
                    continue
                expanded_seen.add(expanded_key)
                _add_redaction(page, rect, c, mode, aliases)
                matched_candidates.add(id(c))
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_PIXELS)
    missing = [f"{c.kind}({c.location})" for c in chosen if id(c) not in matched_candidates]
    if missing:
        doc.close()
        raise RuntimeError(f"삭제 영역 재탐색 실패: {', '.join(missing[:5])}")
    if mode == "delete":
        # 원본의 텍스트층, 이미지 XObject, 주석, 첨부, 양식, JavaScript 및 숨은
        # 객체가 결과에 복사되지 않도록 처리된 화면만 새 PDF에 넣는다.
        sanitized = fitz.open()
        try:
            for page in doc:
                pixmap = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
                new_page = sanitized.new_page(width=page.rect.width, height=page.rect.height)
                new_page.insert_image(new_page.rect, stream=pixmap.tobytes("png"))
            sanitized.set_metadata({})
            sanitized.save(output, garbage=4, clean=True, deflate=True)
        finally:
            sanitized.close()
            doc.close()
        _verify_flattened_pdf_structure(output)
        _verify_pdf_output(output, chosen, verify_pixels=True)
    else:
        doc.set_metadata({})
        doc.del_xml_metadata()
        doc.save(output, garbage=4, clean=True, deflate=True)
        doc.close()
        _verify_pdf_output(output, chosen)
    return output


def _verify_flattened_pdf_structure(output: Path) -> None:
    """완전 삭제 PDF가 새 페이지 이미지 외의 능동·숨은 객체를 갖지 않는지 확인한다."""
    import fitz
    doc = fitz.open(output)
    problems: list[str] = []
    try:
        if doc.embfile_count():
            problems.append("첨부파일")
        if doc.get_xml_metadata():
            problems.append("XMP 메타데이터")
        for page_no, page in enumerate(doc, 1):
            if page.get_text("text").strip():
                problems.append(f"{page_no}페이지 텍스트층")
            if page.first_annot is not None or page.first_widget is not None:
                problems.append(f"{page_no}페이지 주석/양식")
            if len(page.get_images(full=True)) != 1:
                problems.append(f"{page_no}페이지 이미지 구조")
    finally:
        doc.close()
    if problems:
        _fail_verification(output, sorted(set(problems)))


def _render_pdf_page_image(page, scale: float):
    from PIL import Image
    import fitz
    pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    return Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples).convert("L"), scale


def _pad_ocr_rect(rect, page_rect):
    """OCR 상자가 한글 획보다 작게 잡혀 글자 가장자리가 남는 현상을 방지한다."""
    import fitz
    height = max(2, rect.height)
    padded = fitz.Rect(rect.x0 - height * 0.15, rect.y0 - height * 0.28, rect.x1 + height * 0.15, rect.y1 + height * 0.28)
    return padded & page_rect


def _expand_to_table_cell_content(image, rect, scale: float):
    """OCR가 주소 일부만 읽어도 같은 표 셀의 실제 글자 전체 경계로 확장한다."""
    import fitz
    x0, y0, x1, y1 = [round(value * scale) for value in (rect.x0, rect.y0, rect.x1, rect.y1)]
    cx, cy = (x0 + x1) // 2, (y0 + y1) // 2
    pixels = image.load()
    width, height = image.size

    def dark(x, y):
        return pixels[x, y] < 150

    horizontal_span = range(max(0, cx - 100), min(width, cx + 101))
    horizontals = [y for y in range(max(0, cy - 180), min(height, cy + 181)) if sum(dark(x, y) for x in horizontal_span) >= len(horizontal_span) * 0.60]
    tops = [y for y in horizontals if y < y0 - 2]
    bottoms = [y for y in horizontals if y > y1 + 2]
    if not (tops and bottoms):
        return rect
    top, bottom = max(tops), min(bottoms)
    vertical_span = range(top + 2, bottom - 1)
    if len(vertical_span) < 5:
        return rect
    verticals = [x for x in range(max(0, cx - 1200), min(width, cx + 1201)) if sum(dark(x, y) for y in vertical_span) >= len(vertical_span) * 0.72]
    lefts = [x for x in verticals if x < x0 - 2]
    rights = [x for x in verticals if x > x1 + 2]
    if not (lefts and rights):
        return rect
    left, right = max(lefts), min(rights)
    if right - left < 20 or bottom - top < 12 or right - left > 900 or bottom - top > 300:
        return rect

    margin = 5
    crop_box = (left + margin, top + margin, right - margin, bottom - margin)
    crop = image.crop(crop_box).point(lambda value: 255 if value < 150 else 0)
    ink = crop.getbbox()
    if not ink:
        return rect
    ix0, iy0 = crop_box[0] + ink[0], crop_box[1] + ink[1]
    ix1, iy1 = crop_box[0] + ink[2], crop_box[1] + ink[3]
    if ix1 < x0 or ix0 > x1 or iy1 < y0 or iy0 > y1:
        return rect
    padding = 2
    return fitz.Rect((ix0 - padding) / scale, (iy0 - padding) / scale, (ix1 + padding) / scale, (iy1 + padding) / scale)


def _residual_values(text: str, chosen: list[Candidate]) -> list[str]:
    """저장된 결과에 원문 값이 그대로 남아 있는지 유형 이름만으로 돌려준다."""
    compact = re.sub(r"\s+", "", text).casefold()
    remaining = []
    for candidate in chosen:
        target = re.sub(r"\s+", "", candidate.value).casefold()
        # 좌표 전용 후보는 실제 문자열이 아니라 영역 표시이므로 검증 대상이 아니다.
        if len(target) < 2 or candidate.bbox is not None:
            continue
        if target in compact:
            remaining.append(f"{candidate.kind}({candidate.location})")
    return sorted(set(remaining))


def _fail_verification(output: Path, remaining: list[str]) -> None:
    try:
        output.unlink()
    except OSError:
        pass
    raise RuntimeError(f"저장 후 검증 실패: 원문이 남은 항목 {', '.join(remaining[:5])}")


def _verify_pdf_output(output: Path, chosen: list[Candidate], verify_pixels: bool = False) -> None:
    if not chosen:
        return
    import fitz
    doc = fitz.open(output)
    try:
        remaining: list[str] = []
        # 사용자가 다른 페이지의 같은 값을 일부러 남겼을 수 있으므로 페이지별로만 확인한다.
        for page_no, page in enumerate(doc, 1):
            local = [c for c in chosen if c.location in {f"{page_no}페이지", f"{page_no}페이지(OCR)"}]
            if local:
                remaining.extend(_residual_values(page.get_text("text"), local))
                if verify_pixels:
                    for config in ("--psm 6", "--psm 11"):
                        data, _scale = _ocr_pdf_page(page, scale=3.0, config=config)
                        ocr_text = " ".join(str(value) for value in data["text"] if str(value).strip())
                        remaining.extend(_residual_values(ocr_text, local))
    finally:
        doc.close()
    if remaining:
        _fail_verification(output, sorted(set(remaining)))


def _ocr_candidate_rects(data: dict, candidates: list[Candidate], scale: float):
    """공백을 무시한 OCR 문자열 일치 범위를 PDF 좌표 사각형으로 변환한다."""
    import fitz
    tokens = []
    compact = ""
    for i, raw in enumerate(data["text"]):
        token = re.sub(r"\s+", "", raw).casefold()
        if not token:
            continue
        start = len(compact)
        compact += token
        tokens.append((start, len(compact), i))
    results = []
    for candidate in candidates:
        target = re.sub(r"\s+", "", candidate.value).casefold()
        if not target:
            continue
        offset = 0
        while True:
            start = compact.find(target, offset)
            if start < 0:
                break
            end = start + len(target)
            indexes = [i for a, b, i in tokens if a < end and b > start]
            if indexes:
                left = min(data["left"][i] for i in indexes) / scale
                top = min(data["top"][i] for i in indexes) / scale
                right = max(data["left"][i] + data["width"][i] for i in indexes) / scale
                bottom = max(data["top"][i] + data["height"][i] for i in indexes) / scale
                results.append((candidate, fitz.Rect(left, top, right, bottom)))
            offset = end
    return results


# 기본 내장 비트맵 글꼴은 한글을 그리지 못한다. 한글 글꼴이 없을 때만 ASCII 코드로 대체한다.
ALIAS_FONT_PATHS = (
    "C:/Windows/Fonts/malgun.ttf", "C:/Windows/Fonts/gulim.ttc", "C:/Windows/Fonts/batang.ttc",
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
)
_alias_font_cache: dict[int, tuple[object, bool]] = {}


def _alias_font(size: int):
    """(글꼴, ASCII 대체 필요 여부)를 돌려준다."""
    if size not in _alias_font_cache:
        from PIL import ImageFont
        loaded = None
        for font_path in ALIAS_FONT_PATHS:
            try:
                loaded = (ImageFont.truetype(font_path, size), False)
                break
            except OSError:
                continue
        if loaded is None:
            try:
                loaded = (ImageFont.truetype("arial.ttf", size), True)
            except OSError:
                loaded = (ImageFont.load_default(), True)
        _alias_font_cache[size] = loaded
    return _alias_font_cache[size]


def _draw_mask(draw, box: tuple[int, int, int, int], candidate: Candidate, mode: str, aliases: dict) -> None:
    draw.rectangle(box, fill="black" if mode == "delete" else "white")
    if mode == "delete":
        return
    font, ascii_only = _alias_font(max(9, min(28, round((box[3] - box[1]) * 0.7))))
    draw.text((box[0], box[1]), replacement_for(candidate, mode, aliases, ascii_only=(ascii_only and mode != "label_ko")), fill="black", font=font)


def _pad_image_box(box: tuple[int, int, int, int], size: tuple[int, int]) -> tuple[int, int, int, int]:
    """OCR 상자 밖으로 나온 획과 JPEG 번짐까지 포함하는 안전 여백을 둔다."""
    left, top, right, bottom = box
    height = max(2, bottom - top)
    pad_x = max(2, round(height * 0.18))
    pad_y = max(2, round(height * 0.30))
    return (
        max(0, left - pad_x), max(0, top - pad_y),
        min(size[0], right + pad_x), min(size[1], bottom + pad_y),
    )


def _process_image(path: Path, chosen: list[Candidate], mode: str, out_dir: Path, aliases: dict) -> Path:
    from PIL import Image, ImageDraw, ImageEnhance, ImageOps
    img = Image.open(path).convert("RGB")
    draw = ImageDraw.Draw(img)
    enlarged = img.resize((img.width * 2, img.height * 2), Image.Resampling.LANCZOS)
    passes = [
        (_ocr_data(img, config="--psm 6"), 1.0),
        (_ocr_data(enlarged, config="--psm 11"), 2.0),
        (_ocr_data(ImageEnhance.Contrast(ImageOps.grayscale(enlarged)).enhance(2.0), config="--psm 11"), 2.0),
    ]
    selected_kinds = {c.kind for c in chosen}
    drawn: set[tuple[int, int, int, int]] = set()
    for candidate in chosen:
        if candidate.bbox is None:
            continue
        box = _pad_image_box(tuple(round(v) for v in candidate.bbox), img.size)
        if box not in drawn:
            drawn.add(box)
            _draw_mask(draw, box, candidate, mode, aliases)
    for data, scale in passes:
        boxes: list[tuple[str, tuple[int, int, int, int]]] = [("이름", box) for box in _ocr_honorific_boxes(data, scale)]
        boxes.extend(_ocr_labeled_line_boxes(data, selected_kinds, scale))
        for kind, box in boxes:
            match = next((c for c in chosen if c.kind == kind), None)
            box = _pad_image_box(box, img.size)
            if match and box not in drawn:
                drawn.add(box)
                _draw_mask(draw, box, match, mode, aliases)
        for candidate, rect in _ocr_candidate_rects(data, chosen, scale):
            box = _pad_image_box(tuple(round(v) for v in (rect.x0, rect.y0, rect.x1, rect.y1)), img.size)
            if box not in drawn:
                drawn.add(box)
                _draw_mask(draw, box, candidate, mode, aliases)
    output = _unique_output(out_dir, _safe_output_stem(path), path.suffix.lower())
    # 촬영 기기·GPS가 담긴 EXIF는 결과물로 옮기지 않는다.
    if output.suffix.lower() in {".jpg", ".jpeg"}:
        img.save(output, quality=95, exif=b"")
    else:
        img.save(output)
    if mode == "delete":
        if chosen and not drawn:
            output.unlink(missing_ok=True)
            raise RuntimeError("삭제 영역 재탐색 실패: 선택 항목의 이미지 좌표를 찾지 못했습니다.")
        _verify_image_output(output, chosen, drawn)
    return output


def _verify_image_output(output: Path, chosen: list[Candidate], boxes: set[tuple[int, int, int, int]]) -> None:
    """저장된 픽셀 마스크와 OCR 잔존 문자열을 모두 확인한다."""
    from PIL import Image, ImageStat
    with Image.open(output) as image:
        rgb = image.convert("RGB")
        for box in boxes:
            left, top, right, bottom = box
            inset = max(1, min(3, (bottom - top) // 8))
            inner = (left + inset, top + inset, right - inset, bottom - inset)
            if inner[2] <= inner[0] or inner[3] <= inner[1]:
                inner = box
            means = ImageStat.Stat(rgb.crop(inner)).mean
            if max(means) > 30:
                output.unlink(missing_ok=True)
                raise RuntimeError("저장 후 검증 실패: 이미지 삭제 영역에 원본 픽셀이 남았습니다.")
        remaining: list[str] = []
        for config in ("--psm 6", "--psm 11"):
            data = _ocr_data(rgb, config=config)
            text = " ".join(str(value) for value in data["text"] if str(value).strip())
            remaining.extend(_residual_values(text, chosen))
    if remaining:
        _fail_verification(output, sorted(set(remaining)))


def _ocr_honorific_boxes(data: dict, scale: float) -> list[tuple[int, int, int, int]]:
    boxes = []
    valid = [i for i, raw in enumerate(data["text"]) if raw.strip()]
    for suffix in (i for i in valid if "님" in re.sub(r"\s+", "", data["text"][i])):
        center_y = data["top"][suffix] + data["height"][suffix] / 2
        choices = [
            i for i in valid
            if data["left"][i] + data["width"][i] <= data["left"][suffix]
            and abs((data["top"][i] + data["height"][i] / 2) - center_y) <= max(data["height"][suffix], data["height"][i]) * 1.5
            and not re.search(r"\d", data["text"][i])
        ]
        if not choices:
            continue
        height = max(1, data["height"][suffix])
        left = max(min(data["left"][i] for i in choices), data["left"][suffix] - height * 12) / scale
        top = min(data["top"][i] for i in choices) / scale
        right = data["left"][suffix] / scale
        bottom = max(data["top"][i] + data["height"][i] for i in choices) / scale
        if right > left:
            boxes.append(tuple(round(v) for v in (left, top, right, bottom)))
    return boxes


def _ocr_labeled_line_boxes(data: dict, selected_kinds: set[str], scale: float):
    """라벨 자체는 남기고 같은 행의 값 영역 전체를 반환한다."""
    lines: dict[tuple[int, int, int], list[int]] = {}
    for i, raw in enumerate(data["text"]):
        if raw.strip():
            key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
            lines.setdefault(key, []).append(i)
    labels = {"이름": {"예금주", "성명", "이름"}, "계좌번호": {"계좌번호", "통장번호"}}
    boxes = []
    for indexes in lines.values():
        compact = ""
        spans = []
        for i in indexes:
            token = re.sub(r"\s+", "", data["text"][i])
            spans.append((len(compact), len(compact) + len(token), i))
            compact += token
        for kind in selected_kinds & labels.keys():
            label = next((name for name in labels[kind] if name in compact), None)
            if not label:
                continue
            value_start = compact.find(label) + len(label)
            value_indexes = [i for start, end, i in spans if end > value_start]
            if not value_indexes:
                continue
            left = min(data["left"][i] for i in value_indexes) / scale
            top = min(data["top"][i] for i in value_indexes) / scale
            right = max(data["left"][i] + data["width"][i] for i in value_indexes) / scale
            bottom = max(data["top"][i] + data["height"][i] for i in value_indexes) / scale
            boxes.append((kind, tuple(round(v) for v in (left, top, right, bottom))))
        if "이름" in selected_kinds:
            suffixes = [i for i in indexes if "님" in re.sub(r"\s+", "", data["text"][i])]
            if suffixes:
                suffix = suffixes[0]
                before_indexes = [i for i in indexes if data["left"][i] < data["left"][suffix]]
                if not before_indexes:
                    continue
                height = max(1, data["height"][suffix])
                line_left = min(data["left"][i] for i in indexes)
                left = max(line_left, data["left"][suffix] - height * 12) / scale
                top = min(data["top"][i] for i in before_indexes) / scale
                right = data["left"][suffix] / scale
                bottom = max(data["top"][i] + data["height"][i] for i in before_indexes) / scale
                if right > left:
                    boxes.append(("이름", tuple(round(v) for v in (left, top, right, bottom))))
    return boxes


def _blank_png() -> bytes:
    """미리보기 이미지를 지우되 파일 목록은 유지해 한글에서 열리도록 한다."""
    from PIL import Image
    buffer = io.BytesIO()
    Image.new("RGB", (1, 1), "white").save(buffer, format="PNG")
    return buffer.getvalue()


def _sanitize_hwpx_preview_text(data: bytes, chosen: list[Candidate], mode: str, aliases: dict) -> bytes:
    codecs = ("utf-16", "utf-8", "cp949") if data[:2] in (b"\xff\xfe", b"\xfe\xff") else ("utf-8", "utf-16", "cp949")
    for codec in codecs:
        try:
            text = data.decode(codec)
        except (UnicodeDecodeError, UnicodeError):
            continue
        return _replace_text(text, chosen, mode, aliases).encode(codec)
    # 읽지 못한 미리보기는 원문이 남는 것보다 비우는 편이 안전하다.
    return b""


def _sanitize_hwpx_package_metadata(data: bytes) -> bytes:
    """content.hpf의 작성자·제목 등 문서 속성을 비운다."""
    from lxml import etree
    try:
        root = etree.fromstring(data)
    except etree.XMLSyntaxError:
        return data
    for node in root.xpath("//*[local-name()='creator' or local-name()='title' or local-name()='subject' or local-name()='description' or local-name()='publisher' or local-name()='contributor']"):
        node.text = ""
    for node in root.xpath("//*[local-name()='meta']"):
        if node.get("name") in {"creator", "lastsaveby", "author", "title"}:
            node.set("content", "")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _process_hwpx(path: Path, chosen: list[Candidate], mode: str, out_dir: Path, aliases: dict) -> Path:
    from lxml import etree
    output = _unique_output(out_dir, _safe_output_stem(path), ".hwpx")
    with zipfile.ZipFile(path, "r") as source, zipfile.ZipFile(output, "w") as target:
        for info in source.infolist():
            data = source.read(info.filename)
            # 미리보기 텍스트·이미지에는 본문이 그대로 들어 있어 함께 처리하지 않으면 원문이 남는다.
            if info.filename.startswith("Preview/") and info.filename.lower().endswith(".txt"):
                data = _sanitize_hwpx_preview_text(data, chosen, mode, aliases)
            elif info.filename.startswith("Preview/"):
                data = _blank_png()
            elif info.filename.endswith(".hpf"):
                data = _sanitize_hwpx_package_metadata(data)
            elif info.filename.startswith("Contents/section") and info.filename.endswith(".xml"):
                root = etree.fromstring(data)
                # 한 셀의 값이 여러 텍스트 런으로 분리된 경우 먼저 셀 전체를 합쳐 치환한다.
                handled_nodes = set()
                for cell in root.xpath("//*[local-name()='tc']"):
                    nodes = cell.xpath(".//*[local-name()='t']")
                    combined = "".join(node.text or "" for node in nodes)
                    replaced = _replace_text(combined, chosen, mode, aliases)
                    if nodes and replaced != combined:
                        nodes[0].text = replaced
                        for node in nodes[1:]:
                            node.text = ""
                        handled_nodes.update(id(node) for node in nodes)
                for node in root.xpath("//*[local-name()='t']"):
                    if node.text and id(node) not in handled_nodes:
                        node.text = _replace_text(node.text, chosen, mode, aliases)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            target.writestr(info, data)
    _verify_hwpx_output(output, chosen)
    return output


def _verify_hwpx_output(output: Path, chosen: list[Candidate]) -> None:
    if not chosen:
        return
    remaining = _residual_values(_hwpx_text(output), chosen)
    if remaining:
        _fail_verification(output, remaining)


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def write_report(results: list[FileResult], output_dir: Path) -> Path:
    report = _unique_output(output_dir, "처리결과", ".csv")
    stamp = datetime.now().isoformat(timespec="seconds")
    with report.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["처리 시각", "원본 파일", "원본 SHA-256", "결과 파일", "상태", "처리 건수", "신뢰도 부족 건수", "메시지"])
        for r in results:
            try:
                digest = sha256(r.source)
            except OSError as exc:
                digest = f"해시 계산 실패({exc.strerror})"
            writer.writerow([stamp, f"원본_{digest[:12]}", digest, r.output.name if r.output else "", r.status, r.processed, r.low_confidence, r.message])
    return report
