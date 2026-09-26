"""최근 쓴 파일, 확정한 비목, 창 크기를 기억한다.

실행 파일 옆이 아니라 사용자 폴더에 둔다. 업무망 PC는 프로그램 폴더에 쓰기가
막혀 있는 경우가 있고, 여러 사람이 같은 폴더의 exe를 쓸 수도 있기 때문이다.
읽기·쓰기 어느 쪽이 실패해도 프로그램은 그대로 동작해야 한다.
"""

from __future__ import annotations

import json
import os
from pathlib import Path

APP = "예산요구입력본"
FILENAME = "settings.json"
DEFAULTS = {
    "recent": {"plan": "", "ubis": "", "last_year": "", "classes": "", "cards": ""},
    "bimok": {},        # "210-01" 또는 "210-01|운영용품" → 7자리
    "extra_years": [],  # 비목 학습에만 쓰는 추가 확정본
    "window": "",       # "1280x840"
    "checked": [],      # 담당자가 '확인함'으로 표시한 지적
    "tutorial_seen": False,   # 첫 실행 안내를 '다음부터 열지 않기'로 닫았는가
}


def folder() -> Path:
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("XDG_CONFIG_HOME") or str(Path.home())
    return Path(base) / APP


def path() -> Path:
    return folder() / FILENAME


def load() -> dict:
    data = {key: (value.copy() if isinstance(value, (dict, list)) else value)
            for key, value in DEFAULTS.items()}
    try:
        stored = json.loads(path().read_text(encoding="utf-8"))
    except Exception:
        return data
    if not isinstance(stored, dict):
        return data
    for key, value in stored.items():
        if key in data and isinstance(value, type(data[key])):
            data[key] = value
    return data


def save(data: dict) -> bool:
    try:
        folder().mkdir(parents=True, exist_ok=True)
        path().write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        return True
    except Exception:
        # 기억하지 못하는 것은 불편할 뿐 오류가 아니다. 조용히 넘어간다.
        return False


def bimok_key(code5: str, name: str = "") -> str:
    return f"{code5}|{name}" if name else code5
