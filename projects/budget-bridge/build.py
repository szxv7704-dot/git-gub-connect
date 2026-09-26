"""실행 파일 만들기. 배치 파일은 이 스크립트를 부르기만 한다.

왜 배치가 아니라 파이썬인가:
cmd 는 배치 파일을 바이트 위치로 되짚어 읽는다. 파일 중간에서 `chcp 65001` 로
코드페이지를 바꾸면 그 뒤 줄들의 위치가 어긋나 명령이 잘린다. 실제로
`openpyxl` 이 `enpyxl` 로 잘리고 한글이 전부 깨졌다.
그래서 배치에는 ASCII만 남기고, 한글 안내와 판단은 전부 여기서 한다.
파이썬 3.6+ 는 콘솔에 유니코드를 직접 쓰므로 코드페이지와 무관하게 한글이 나온다.
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
VERSION = "2.16.0"
LABEL = "UI개선"
NAME = f"예산요구_입력본_만들기_{VERSION}_{LABEL}"
DIST = "dist_ui"          # 기존 dist\ 의 실행 파일을 덮어쓰지 않는다.
WORK = "build_ui"
ENTRY = "budget_bridge.py"
HIDDEN = ["converter", "plan_parser", "edufine_form", "bimok_resolver", "bimok_groups",
          "crosscheck", "ubis_review", "hwpx_native_parser", "ui_common", "settings",
          "report_export", "program_classes", "program_cards", "tutorial", "native_ui", "supplement",
          "supplement_form"]
SAMPLES = [
    "(제거)(본예산)부서별사업별 설명서(작성)_2026.9.17_17_42_49.hwpx",
    "(에듀파인)2026세출예산요구내역.xlsx",
    "(본예산세출요구)검토조서(서식1)_2026.9.17_17_24_48.xlsx",
]
SAMPLE_FOLDER = HERE.parent.parent / "3 예산이 뭐예요" / "ubis꺼져"


def say(text: str = "") -> None:
    print(text, flush=True)


def stop(message: str, code: int = 1) -> None:
    say()
    say(f"[중단] {message}")
    say()
    input("닫으려면 Enter 를 누르세요...")
    sys.exit(code)


def run(arguments: list[str], why: str) -> None:
    result = subprocess.run(arguments)
    if result.returncode != 0:
        stop(why)


def sample_folder() -> Path | None:
    """표본 파일이 있는 폴더. 환경변수 UBIS_SAMPLES → 이 폴더 → 작업 폴더 순으로 찾는다."""
    env = [Path(one) for one in os.environ.get("UBIS_SAMPLES", "").split(os.pathsep) if one]
    for folder in (*env, HERE, SAMPLE_FOLDER):
        if all((folder / name).exists() for name in SAMPLES):
            return folder
    return None


def option(flag: str, fallback: str) -> str:
    """--name=... 또는 --name ... 둘 다 받는다."""
    for index, argument in enumerate(sys.argv):
        if argument.startswith(flag + "="):
            return argument.split("=", 1)[1]
        if argument == flag and index + 1 < len(sys.argv):
            return sys.argv[index + 1]
    return fallback


def main() -> None:
    skip = "--skip-samples" in sys.argv
    name = option("--name", NAME)
    dist = option("--dist", DIST)
    say("=" * 48)
    say("  예산요구 입력본 만들기 — 실행 파일 생성")
    say("=" * 48)
    say()
    say(f"[1/4] 파이썬  {sys.version.split()[0]}  ({sys.executable})")
    say(f"       만들 파일: {dist}\\{name}.exe")

    say()
    say("[2/4] 패키지 설치 (openpyxl, pyinstaller)")
    run([sys.executable, "-m", "pip", "install", "--quiet", "--disable-pip-version-check",
         "--no-warn-script-location", "--upgrade", "openpyxl", "pyinstaller"],
        "패키지 설치에 실패했습니다. 인터넷 연결이나 사내 프록시 설정을 확인해 주세요.")

    say()
    folder = sample_folder()
    if folder:
        say(f"[표본] {folder}")
    elif skip:
        say("[주의] 표본 없이 진행합니다. 핵심 회귀 테스트를 건너뜁니다.")
    else:
        say("표본 파일을 찾지 못했습니다. 아래 세 개가 필요합니다.")
        for name in SAMPLES:
            say(f"   - {name}")
        say()
        say("이 셋이 없으면 사설영역 불릿 계층, 비목 추출, 레벨·순번 규칙 검증이")
        say("전부 건너뛰어집니다. 통과한 것처럼 보여도 아무것도 검증하지 못합니다.")
        say()
        say("해결: 아래 폴더에 두거나, 이 폴더에 복사하세요.")
        say(f"   {SAMPLE_FOLDER}")
        say()
        stop("그래도 빌드하려면 build_exe.bat 에 --skip-samples 를 붙여 실행하세요.")

    say()
    say("[3/4] 테스트")
    # 화면 테스트(test_native_ui)도 같이 돌린다. 창을 실제로 띄워 좌표를 재는
    # 시험이라 빌드하는 그 PC에서 돌려야 뜻이 있다. 빼 놓으면 있으나 마나다.
    result = subprocess.run([sys.executable, "-m", "unittest",
                             "test_budget_bridge", "test_native_ui"], cwd=str(HERE))
    if result.returncode != 0:
        stop("테스트가 실패했습니다. 실행 파일을 만들지 않습니다.\n"
             "       깨진 프로그램을 배포하는 것보다 만들지 않는 편이 안전합니다.")
    say("      통과 (자료 검사 + 화면)")

    say()
    say("[4/4] 실행 파일 생성 (몇 분 걸립니다)")
    arguments = [sys.executable, "-m", "PyInstaller", "--onefile", "--noconsole", "--clean",
                 "--name", name, "--distpath", dist, "--workpath", WORK]
    for module in HIDDEN:
        arguments += ["--hidden-import", module]
    arguments.append(ENTRY)
    run(arguments, "빌드에 실패했습니다.")

    target = HERE / dist / f"{name}.exe"
    say()
    say("=" * 48)
    say(f"  완료: {target}")
    say("=" * 48)
    say()
    say(f"{WORK}\\ 와 {dist}\\ 는 저장소에 커밋하지 마세요.")
    say()
    try:
        os.startfile(str(HERE / dist))       # noqa: S606 - Windows 탐색기 열기
    except Exception:
        pass
    input("닫으려면 Enter 를 누르세요...")


if __name__ == "__main__":
    main()
