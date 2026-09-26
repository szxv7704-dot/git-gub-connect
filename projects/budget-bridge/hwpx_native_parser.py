"""HWPX 사업계획서를 파이썬 표준 라이브러리만으로 직접 읽는다.

외부 변환도구(kordoc CLI/MCP)에 의존하지 않는다. 배포한 실행 파일이 다른 PC에서
그대로 동작해야 하므로, 설치 여부를 알 수 없는 외부 명령을 호출하지 않는다.

HWPX는 ZIP이고 본문은 Contents/section*.xml 이다. 이 모듈은 문서 순서대로
문단과 표를 훑어 kordoc 청크와 같은 모양의 dict 목록을 만든다. 덕분에 기존
parse_kordoc_chunks 이후의 처리 흐름을 그대로 재사용한다.
"""

from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Iterable

HP = "{http://www.hancom.co.kr/hwpml/2011/paragraph}"
HEADING = re.compile(r"^\s*\d+\.\s+\S")


def is_private_use(char: str) -> bool:
    """한글 문서의 기호 불릿은 표준 문자가 아니라 사설영역(PUA) 글리프다.

    사업계획서의 4단계 불릿은 U+F02FB 처럼 보이지 않는 코드포인트로 저장된다.
    '▸' 로 문자열 비교하면 4단계 항목을 통째로 놓치고, 상위 합계는 그대로라
    결과 파일만 보면 누락을 알아챌 수 없다. 반드시 범위로 판정한다.
    """
    code = ord(char)
    return 0xE000 <= code <= 0xF8FF or 0xF0000 <= code <= 0xFFFFD or 0x100000 <= code <= 0x10FFFD


def _text(element) -> str:
    """셀 텍스트를 모은다. 중첩된 표 안으로는 들어가지 않는다.

    설명서 서식은 한 사업 전체를 감싸는 바깥 표 안에 실제 내용 표가 들어 있다.
    그냥 모든 하위 텍스트를 이어붙이면 사업 한 건이 거대한 셀 하나로 뭉개진다.
    """
    parts: list[str] = []

    def collect(node) -> None:
        for child in node:
            if child.tag == HP + "tbl":
                continue
            if child.tag == HP + "t":
                parts.append(child.text or "")
            collect(child)

    collect(element)
    return "".join(parts)


def _child_tables(element) -> list:
    """바로 아래 단계의 표만 찾는다. 표를 만나면 그 아래로 내려가지 않는다."""
    found: list = []

    def scan(node) -> None:
        for child in node:
            if child.tag == HP + "tbl":
                found.append(child)
            else:
                scan(child)

    scan(element)
    return found


def _escape(value: str) -> str:
    return value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _table_html(table) -> str:
    rows = []
    for tr in table.findall(HP + "tr"):
        cells = []
        for tc in tr.findall(HP + "tc"):
            span = tc.find(HP + "cellSpan")
            col = span.get("colSpan", "1") if span is not None else "1"
            row = span.get("rowSpan", "1") if span is not None else "1"
            attrs = ""
            if col != "1":
                attrs += f' colspan="{col}"'
            if row != "1":
                attrs += f' rowspan="{row}"'
            cells.append(f"<td{attrs}>{_escape(_text(tc).strip())}</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return "<table>" + "".join(rows) + "</table>"


def _emit_table(table, output: list) -> None:
    """바깥 표를 먼저 싣고, 그 안에 든 표를 이어서 싣는다."""
    output.append(("table", table))
    for inner in _child_tables(table):
        _emit_table(inner, output)


def _walk(element, output: list) -> None:
    """문서 순서를 유지하며 문단 텍스트와 표를 모은다.

    표 안의 문단까지 훑으면 셀 내용이 본문 문단으로 새어 나오므로, 표를 만나면
    그 아래로는 내려가지 않는다.
    """
    for child in element:
        if child.tag == HP + "tbl":
            _emit_table(child, output)
            continue
        if child.tag == HP + "p":
            inner = _child_tables(child)
            if inner:
                for table in inner:
                    _emit_table(table, output)
            else:
                text = _text(child).strip()
                if text:
                    output.append(("text", text))
            continue
        _walk(child, output)


def native_chunks(path: str) -> list[dict]:
    """HWPX를 kordoc 청크와 같은 구조의 dict 목록으로 변환한다."""
    chunks: list[dict] = []
    breadcrumb = ""
    pending: list[str] = []

    def flush() -> None:
        if pending:
            chunks.append({"type": "text", "text": "\n".join(pending), "breadcrumb": [breadcrumb] if breadcrumb else []})
            pending.clear()

    with zipfile.ZipFile(path) as archive:
        # 구역이 열 개를 넘으면 이름순은 문서 순서가 아니다. section10 이 section2
        # 앞에 오고, 그러면 사업 차례(order)가 뒤섞여 UBIS 대조가 엉뚱한 사업과
        # 붙는다. 번호로 센다.
        names = [name for name in archive.namelist()
                 if re.fullmatch(r"Contents/section\d+\.xml", name)]
        names.sort(key=lambda name: int(re.search(r"(\d+)\.xml$", name).group(1)))
        if not names:
            raise ValueError("HWPX 본문(Contents/section0.xml)을 찾지 못했습니다. 한글 문서가 맞는지 확인해 주세요.")
        for name in names:
            items: list = []
            _walk(ET.fromstring(archive.read(name)), items)
            for kind, value in items:
                if kind == "text":
                    if HEADING.match(value):
                        flush()
                        breadcrumb = value
                    pending.append(value)
                else:
                    flush()
                    chunks.append({"type": "table", "text": _table_html(value), "breadcrumb": [breadcrumb] if breadcrumb else []})
    flush()
    return chunks


def parse_hwpx_native(path: str, candidates: Iterable) -> list:
    """HWPX 경로를 받아 사업 목록을 만든다. kordoc 경로를 대체한다."""
    from kordoc_adapter import parse_kordoc_chunks

    return parse_kordoc_chunks(native_chunks(path), list(candidates))
