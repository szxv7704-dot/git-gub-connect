"""화면이 공통으로 쓰는 표기 규칙과 색.

숫자는 사람이 자릿수를 세지 않아도 되게 적는다. 천단위 쉼표, 소수점 없음,
오른쪽 정렬. 차액은 부호를 붙여 늘었는지 줄었는지 한눈에 보이게 한다.
"""

from __future__ import annotations

from typing import Optional

# 평면 색만 쓴다. tkinter에는 그라데이션도 그림자도 없다.
GROUND = "#F5F7FA"
PANEL = "#FFFFFF"
INK = "#243447"
BODY = "#405568"
MUTED = "#657486"
FAINT = "#9AA1AB"
LINE = "#DEE6EC"
HEAD = "#EAF3F7"
ACCENT = "#387C9C"
ACCENT_DIM = "#2D6884"

ERROR = "#C0392B"
ERROR_BG = "#FDECEA"
ERROR_INK = "#7A3229"
WARN = "#B26A00"
WARN_BG = "#FFF4E0"
OK = "#278579"
OK_BG = "#E8F5EE"
GRAY_BG = "#F0F2F5"

# 사용법 안내 말풍선. 작업 카드(흰 바탕)와 한눈에 갈리도록 색을 뒤집는다 — 안내는
# 어두운 바탕에 밝은 글자. 예전에는 안내도 흰 카드라 어느 것이 안내인지 헷갈렸다.
GUIDE_BG = "#243447"
GUIDE_SOFT = "#34495E"      # 말풍선 안의 '자료 받는 경로' 상자, 버튼을 올렸을 때
GUIDE_TEXT = "#FFFFFF"
GUIDE_BODY = "#D5E0E8"
GUIDE_TAG = "#8FD3C8"       # '사용법 안내' 머리표

SEVERITY_TAG = {"오류": "error", "확인 필요": "warn", "안내": "info"}
SEVERITY_COLOR = {"오류": ERROR, "확인 필요": WARN, "안내": MUTED}

FONT = "맑은 고딕"

# 화면에는 내부 번호(목표1·2A…) 대신 무엇을 검사했는지를 적는다. 담당자는 '목표 2A'가
# 무엇인지 외우지 않는다. 내부 코드는 README·테스트와 맞추려고 그대로 둔다.
GOAL_LABEL = {
    "목표1": "설명서 검산",
    "목표2A": "UBIS 대조",
    "목표2B": "입력 후 대조",
    "목표3": "비목",
    "목표4": "사업 분류",
    "과제카드": "과제카드",
    "추경연결": "직전 차수 연결",
}


def goal_label(goal: str) -> str:
    return GOAL_LABEL.get(goal, goal)


def year_of(name: str):
    """파일 이름에 적힌 회계연도. '2027검토조서(서식1)_2026.9.21_9_15_28' → 2027.

    뒤에 붙는 출력 날짜(2026.9.21)는 연도가 아니므로 '숫자.숫자'로 이어지는 것은 뺀다.
    이름에 연도가 없으면 None — 모르면 말하지 않는다.
    """
    import re

    found = re.search(r"(?<!\d)(20\d{2})(?![\d.])", name or "")
    return int(found.group(1)) if found else None


def money(value: Optional[float]) -> str:
    """3240000.0 → '3,240,000' · 없으면 빈 칸."""
    if value is None or value == "":
        return ""
    try:
        return f"{round(float(value)):,}"
    except (TypeError, ValueError):
        return str(value)


def signed(value: Optional[float]) -> str:
    """차액. 0은 '0', 양수는 +, 음수는 진짜 빼기 기호로."""
    if value is None or value == "":
        return ""
    try:
        number = round(float(value))
    except (TypeError, ValueError):
        return str(value)
    if number == 0:
        return "0"
    return f"+{number:,}" if number > 0 else f"−{abs(number):,}"


def gap_color(value: Optional[float]) -> str:
    if value is None:
        return BODY
    try:
        number = float(value)
    except (TypeError, ValueError):
        return BODY
    if number == 0:
        return BODY
    return ERROR if number > 0 else ACCENT


def apply_tags(tree) -> None:
    """심각도를 색으로. 글자만으로는 오류 1건이 안내 8건에 묻힌다."""
    tree.tag_configure("error", background=ERROR_BG, foreground=ERROR_INK)
    tree.tag_configure("warn", background=WARN_BG, foreground="#6B4A08")
    tree.tag_configure("info", background=PANEL, foreground=MUTED)
    tree.tag_configure("ok", background=OK_BG, foreground="#1E5C3B")
    tree.tag_configure("child", background=PANEL, foreground=BODY)
    tree.tag_configure("strong", background=PANEL, foreground=INK)


def open_path(path: str) -> None:
    """파일을 그 형식의 기본 프로그램(한글·엑셀)으로 연다. 실패하면 OSError."""
    import os
    import subprocess
    import sys

    if sys.platform.startswith("win"):
        os.startfile(path)                                  # noqa: S606 - 사용자가 고른 자료 파일
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])                    # noqa: S603,S607
    else:
        subprocess.Popen(["xdg-open", path])                # noqa: S603,S607
