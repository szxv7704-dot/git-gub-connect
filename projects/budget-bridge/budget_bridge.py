"""예산요구 입력본 만들기 — 화면.

판단은 전부 converter / crosscheck / bimok_groups 에 있다. 여기서 하는 일은
파일을 고르고, 결과를 보기 좋게 늘어놓고, 담당자의 확정을 받아 넘기는 것뿐이다.
"""

from __future__ import annotations

import queue
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import settings as store

from converter import Session
from crosscheck import issue_row
from report_export import export_issues
from native_ui import FilePage, RoundedButton
from ui_common import (ACCENT, BODY, ERROR, ERROR_BG, ERROR_INK, FAINT, FONT, GROUND, HEAD, INK, LINE, MUTED, OK, OK_BG,
                       PANEL, WARN, WARN_BG, SEVERITY_TAG, apply_tags, gap_color, goal_label, money, open_path,
                       signed)

TITLE = "예산요구 입력본 만들기"
STEPS = ["1 자료 선택", "2 검사 결과", "3 비목 확정", "4 입력본 만들기", "5 입력 후 대조"]
PAGES = ["files", "result", "bimok", "make", "verify"]
SUPPLEMENT_CAPTIONS = {"goal1": "설명서 검산 오류 (추경)", "goal2": "UBIS 추경 대조 오류",
                       "goal3": "직전 차수 연결", "total": "추경 증감 합계 (천원)"}
WRITING = "작성 점검"
WRITING_PREFIX = "작성"


def _kind(goal: str, issue) -> str:
    """목록 구분. 작성요령 점검은 '작성 점검', 그 밖은 심각도(오류)."""
    return WRITING if goal.startswith(WRITING_PREFIX) else issue.severity


VERIFY_IDLE = ("아직 대조하지 않았습니다. 4단계에서 만든 입력본이 있으면 한 줄씩, 없으면 사업별 금액만 맞춰 봅니다.")


# --------------------------------------------------------------------------- 작은 조각

def label(parent, text, size=10, bold=False, fg=BODY, bg=None, **kw):
    return tk.Label(parent, text=text, font=(FONT, size, "bold" if bold else "normal"),
                    fg=fg, bg=bg or parent["bg"], **kw)


def button(parent, text, command, primary=False, width=None):
    return RoundedButton(parent, text, command, primary=primary, width=width)


def make_tree(parent, columns, widths, anchors=None, height=16):
    holder = tk.Frame(parent, bg=PANEL)
    tree = ttk.Treeview(holder, columns=columns, show="headings", height=height)
    anchors = anchors or ["w"] * len(columns)
    for name, width, anchor in zip(columns, widths, anchors):
        tree.heading(name, text=name)
        tree.column(name, width=width, anchor=anchor, stretch=(width == 0))
    bar = ttk.Scrollbar(holder, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=bar.set)
    tree.pack(side="left", fill="both", expand=True)
    bar.pack(side="right", fill="y")
    apply_tags(tree)
    return holder, tree


def unique_issues(pairs) -> list:
    """(검사, 지적) 목록에서 글자 하나 다르지 않은 지적을 한 번만 남긴다."""
    unique, seen = [], set()
    for goal, issue in pairs:
        key = (goal, issue.severity, issue.project, issue.item, issue.message)
        if key not in seen:
            seen.add(key)
            unique.append((goal, issue))
    return unique


def friendly(error: Exception) -> str:
    """파이썬 예외를 담당자가 할 일을 알 수 있는 문장으로.

    'File is not a zip file' 같은 영어 한 줄로는 어느 파일을 어떻게 하라는 건지 모른다.
    이 프로그램이 직접 던진 ValueError 는 이미 한국어라 그대로 둔다.
    """
    import zipfile

    name = type(error).__name__
    if isinstance(error, zipfile.BadZipFile) or name == "InvalidFileException":
        return ("파일을 열 수 없습니다. 사업설명서는 한글에서 'HWPX'로, 나머지는 엑셀 '.xlsx'로 저장한 "
                "파일이어야 합니다(.hwp · .xls · .csv 는 읽지 못합니다). 파일이 손상됐을 수도 있습니다.\n\n"
                f"(자세히: {error})")
    if isinstance(error, PermissionError):
        return f"파일을 읽을 권한이 없습니다. 엑셀·한글에서 열려 있으면 닫고 다시 시도해 주세요.\n\n(자세히: {error})"
    if isinstance(error, FileNotFoundError):
        return f"파일을 찾을 수 없습니다. 옮기거나 지웠을 수 있습니다. 다시 선택해 주세요.\n\n(자세히: {error})"
    if isinstance(error, ValueError):
        return str(error)
    return f"예상하지 못한 오류로 불러오지 못했습니다. 자료 파일이 맞는지 확인해 주세요.\n\n({name}: {error})"


# --------------------------------------------------------------------------- 본체

class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(TITLE)
        self.configure(bg=GROUND)
        self.settings = store.load()
        self.geometry(self.settings.get("window") or "1440x900")
        self.minsize(1080, 680)

        self.loading = False
        self.session = Session()
        self.generated = ""
        self.issues: list = []          # 목록에 실린 지적, 표시 순서대로
        self.current = 0
        self.filter = "오류"
        self.extra_years: list = []
        self.goal2_report = None
        self.verify_report = None
        self.groups: list = []
        self.page = "files"
        self.query = tk.StringVar()
        self.progress_queue: queue.Queue = queue.Queue()

        self._style()
        self._header()
        self.body = tk.Frame(self, bg=GROUND)
        self.body.pack(fill="both", expand=True)
        self.pages = {}
        for name, builder in (("files", self._page_files), ("result", self._page_result),
                              ("detail", self._page_detail), ("bimok", self._page_bimok),
                              ("make", self._page_make), ("verify", self._page_verify)):
            frame = tk.Frame(self.body, bg=GROUND)
            frame.place(relwidth=1, relheight=1)
            self.pages[name] = frame
            builder(frame)
            if name != "files":
                self._empty_state(frame)
        self._status()
        self._shortcuts()
        self.extra_years = list(self.settings.get("extra_years") or [])
        if self.extra_years:
            self.year_note.configure(text=f"{len(self.extra_years) + 1}부 기억함 · 서식 본은 "
                                          f"{Path(self.paths['last_year'].get()).name}", fg=OK)
        self.show("files")
        self.protocol("WM_DELETE_WINDOW", self._close)
        if not self.settings.get("tutorial_seen"):
            # 처음 켠 사람에게는 파일 칸 네 개보다 '무엇을 왜 넣는지'가 먼저다.
            self.after(240, self.show_tutorial)

    # ----------------------------------------------------------------- 뼈대
    def _style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Treeview", font=(FONT, 9), rowheight=26,
                        background=PANEL, fieldbackground=PANEL, foreground=BODY, borderwidth=0)
        style.configure("Treeview.Heading", font=(FONT, 9, "bold"), background=HEAD,
                        foreground=BODY, relief="flat", padding=(6, 5))
        style.map("Treeview", background=[("selected", "#D6E4F0")], foreground=[("selected", INK)])
        style.configure("Bar.Horizontal.TProgressbar", troughcolor=HEAD, background=ACCENT,
                        borderwidth=0, thickness=8)

    def _header(self) -> None:
        top = tk.Frame(self, bg=PANEL)
        top.pack(fill="x")
        line = tk.Frame(top, bg=PANEL)
        line.pack(fill="x", padx=22, pady=(12, 0))
        label(line, TITLE, 14, True, INK, PANEL).pack(side="left")
        self.subtitle = label(line, "", 9, False, MUTED, PANEL)
        self.subtitle.pack(side="left", padx=10)
        self.tutorial_button = button(line, "사용법 다시 보기  (F1)", self.show_tutorial)
        self.tutorial_button.pack(side="right")

        tabs = tk.Frame(top, bg=PANEL)
        tabs.pack(fill="x", padx=18, pady=(10, 0))
        self.tabs = []
        for index, text in enumerate(STEPS):
            holder = tk.Frame(tabs, bg=PANEL, cursor="hand2")
            holder.pack(side="left")
            text_label = label(holder, f"  {text}  ", 10, False, FAINT, PANEL)
            text_label.pack(pady=(4, 5))
            underline = tk.Frame(holder, bg=PANEL, height=3)
            underline.pack(fill="x")
            for widget in (holder, text_label):
                widget.bind("<Button-1>", lambda _event, i=index: self.show(PAGES[i]))
            self.tabs.append((text_label, underline))
        tk.Frame(top, bg=LINE, height=1).pack(fill="x")

        self.banner = tk.Frame(self, bg=WARN_BG)
        self.banner_text = label(self.banner, "", 9, False, "#6B4A08", WARN_BG)
        self.banner_text.pack(side="left", padx=22, pady=9)
        self.banner_button = button(self.banner, "첫 항목으로 이동 (F3)", self.next_issue)
        self.banner_button.pack(side="right", padx=22, pady=6)

    def _empty_state(self, frame) -> None:
        """자료를 불러오기 전에 뒤 단계 탭을 누르면 보이는 안내.

        예전에는 빈 표와 '—' 만 보여서 고장 난 것처럼 보였다.
        """
        cover = tk.Frame(frame, bg=GROUND)
        box = tk.Frame(cover, bg=PANEL, highlightbackground=LINE, highlightthickness=1)
        box.place(relx=0.5, rely=0.4, anchor="center")
        label(box, "아직 불러온 자료가 없습니다", 14, True, INK, PANEL).pack(padx=40, pady=(28, 6))
        label(box, "1단계에서 사업설명서와 지난 K-에듀파인 파일을 고르고\n"
                   "'자료 불러오고 검사하기'(F5)를 누르면 이 화면이 채워집니다.",
              10, False, MUTED, PANEL, justify="center").pack(padx=40)
        button(box, "1 자료 선택으로 가기", lambda: self.show("files"), primary=True).pack(pady=(16, 28))
        frame.cover = cover

    def _update_tabs(self) -> None:
        """탭 글자 앞에 끝낸 단계는 ✓ 를 붙인다. 지금 어디까지 왔는지 한눈에 보이게."""
        loaded = bool(self.session.projects)
        done = set()
        if loaded:
            done.add(0)
            if not self.session.blocking() or self.session.is_supplement:
                done.add(1)
            if not self.session.unsettled:
                done.add(2)
        if self.generated:
            done.add(3)
        if self.verify_report is not None:
            done.add(4)
        self.done_steps = done
        supplement = self.session.is_supplement
        if supplement:
            # 추경은 3·4단계(비목·입력본)가 없다. 5단계 입력 후 대조는 한다.
            done = {0, 1} | ({4} if self.verify_report is not None else set())
            self.done_steps = done
        step = PAGES.index(self.page) if self.page in PAGES else 1
        for index, (text_label, underline) in enumerate(self.tabs):
            on = index == step
            mark = "✓ " if index in done else ""
            reachable = (loaded or index == 0) and not (supplement and index in (2, 3))
            text_label.configure(text=f"  {mark}{STEPS[index]}  ",
                                 fg=ACCENT if on else (OK if index in done else BODY if reachable else FAINT),
                                 font=(FONT, 10, "bold" if on else "normal"))
            underline.configure(bg=ACCENT if on else PANEL)
        stage = min((i for i in range(5) if i not in done), default=4)
        self.file_view.flow.set_stage(stage, done)

    def _status(self) -> None:
        bar = tk.Frame(self, bg=HEAD, height=26)
        bar.pack(fill="x", side="bottom")
        self.status = label(bar, "1번 사업설명서와 2번 지난 K-에듀파인 파일을 고르고 F5 를 누르세요.",
                            9, False, MUTED, HEAD, anchor="w")
        self.status.pack(side="left", padx=14, pady=4)
        self.hint = label(bar, "", 9, False, FAINT, HEAD, anchor="e")
        self.hint.pack(side="right", padx=14)

    def _shortcuts(self) -> None:
        self.bind_all("<F1>", lambda _e: self.show_tutorial())
        self.bind_all("<F5>", lambda _e: self.load())
        self.bind_all("<Control-s>", lambda _e: self.save())
        self.bind_all("<Control-S>", lambda _e: self.save())
        self.bind_all("<Control-f>", lambda _e: self.focus_search())
        self.bind_all("<Control-F>", lambda _e: self.focus_search())
        self.bind_all("<F3>", lambda _e: self.next_issue())
        self.bind_all("<Shift-F3>", lambda _e: self.previous_issue())
        self.bind_all("<Escape>", lambda _e: self.file_view.close_guide() if self.file_view.active >= 0 else self.show("result") if self.page == "detail" else None)
        for index in range(5):
            self.bind_all(f"<Control-Key-{index + 1}>", lambda _e, i=index: self.show(PAGES[i]))

    def show(self, name: str) -> None:
        if name in ("bimok", "make") and self.session.is_supplement:
            messagebox.showinfo("추경 점검", "추경 설명서는 산출식이 증감분만 있어 K-에듀파인 입력본의 원천이 될 수 "
                                            "없습니다.\n\n비목 확정 · 입력본 만들기는 본예산 설명서로 불러왔을 때만 씁니다. "
                                            "추경은 K-에듀파인에 입력한 뒤 5 입력 후 대조에서 내려받은 파일을 맞춰 봅니다.")
            name = "result"
        if name != "files" and hasattr(self, "file_view") and self.file_view.active >= 0:
            self.file_view.close_guide()
        self.page = name
        self.pages[name].tkraise()
        cover = getattr(self.pages[name], "cover", None)
        if cover is not None:
            if self.session.projects:
                cover.place_forget()
            else:
                cover.place(relwidth=1, relheight=1)
                cover.tkraise()
        if name == "make":
            self.refresh_make()
        # 비목 화면에서 '비목 확정으로' 버튼은 제자리 걸음이다. 그 화면에서는 감춘다.
        soft_here = name == "bimok" and self.session.projects and not self.session.blocking()
        if soft_here and self.banner_button.winfo_manager():
            self.banner_button.pack_forget()
        elif not soft_here and not self.banner_button.winfo_manager():
            self.banner_button.pack(side="right", padx=22, pady=6)
        self._update_tabs()
        if self.session.projects and name in ("bimok", "make", "verify"):
            left = len(self.session.unsettled)
            self.status.configure(text={
                "bimok": f"비목 미확정 {left:,}행" if left else "비목이 모두 정해졌습니다.",
                "make": f"마지막으로 만든 입력본: {self.generated}" if self.generated else "아직 입력본을 만들지 않았습니다.",
                "verify": "입력 후 내려받은 파일을 골라 대조하세요.",
            }[name])
        elif self.session.projects and name == "result":
            self.status.configure(text=f"{len(self.issues):,}건 표시 중")
        hints = {
            "files": "F1 사용법 · F5 불러오기",
            "result": "F3 다음 지적 · Enter 상세 · Ctrl+F 사업 찾기 · Ctrl+S 입력본 만들기",
            "detail": "Esc 목록 · F3 다음 지적",
            "bimok": "Enter 비목 정하기 · → 펼치기 · ← 접기",
            "make": "Ctrl+S 입력본 만들기",
            "verify": "Ctrl+1~5 단계 이동",
        }
        self.hint.configure(text=hints.get(name, ""))

    def show_tutorial(self) -> None:
        self.file_view.open_guide(0)

    def _close(self) -> None:
        self.settings["window"] = f"{self.winfo_width()}x{self.winfo_height()}"
        store.save(self.settings)
        self.destroy()

    # ----------------------------------------------------------------- ① 파일
    # (키, 번호, 자료명, 필수/선택, 한 줄 요약, 이 자료로 하는 일, 받는 곳, 예시 파일명, 확장자)
    FILE_CARDS = (
        ("plan", "1", "사업설명서 HWPX", "필수",
         "산출근거의 원문입니다. 본예산·추경 모두 받습니다.",
         "사업명, 요구내용, 산출식, 금액, 단위를 읽고 계산 오류를 찾습니다. "
         "추경 설명서를 넣으면 증감(A−B)·산출식·합계를 점검하는 추경 점검으로 바뀝니다.",
         "유비스 → 예산요구 → 사업별설명서 → 사업별 설명서 출력 → 부서별(사업별 설명서)",
         "2026사업별 설명서.hwpx",
         [("한글 문서", "*.hwpx")]),
        ("last_year", "2", "지난 K-에듀파인 입력본·다운로드 파일", "필수",
         "입력 양식과 비목을 배웁니다. 추경 점검에는 필요 없습니다.",
         "올해 입력본의 모양을 정하고, 과거에 확정한 원가통계비목을 자동으로 연결합니다. "
         "여러 부를 선택할 수 있습니다.",
         "K-에듀파인 → 재정사업관리 → 예산관리 → 예산편성 → 사업관리카드 → 세출 조회 → "
         "단위과제카드번호 → 예산편성 → 본예산 → 파일다운로드",
         "2026(k에듀파인)세출예산요구내역.xlsx",
         [("엑셀", "*.xlsx")]),
        ("ubis", "3", "UBIS 세출요구 검토조서", "선택",
         "사업별 금액을 대조합니다.",
         "사업설명서의 합계액과 UBIS 검토조서의 요구액이 같은지 확인합니다. "
         "추경이면 추경 검토조서를 넣습니다 — 증감과 추경요구액을 맞추고, 같은 폴더의 직전 차수 "
         "검토조서로 기정액이 이어지는지도 봅니다. 넣지 않으면 이 금액 대조만 건너뜁니다.",
         "유비스 → 예산요구 → 세출요구출력",
         "검토조서(서식1).xlsx",
         [("엑셀", "*.xlsx")]),
        ("classes", "4", "세출예산 사업별 분류표 (별표 3)", "선택",
         "사업 위계를 확인합니다.",
         "정책·단위·세부사업 분류와 코드를 대조하여 K-에듀파인 입력 행의 위계를 더 정확하게 검토합니다.",
         "유비스 → 예산요구 → 세출요구",
         "세출예산 사업별 분류.hwpx",
         [("한글 문서", "*.hwpx")]),
        ("cards", "5", "단위과제카드 목록", "선택",
         "내가 입력할 사업만 남깁니다.",
         "K-에듀파인은 자기 카드에 딸린 사업만 입력할 수 있습니다. 넣으면 과 전체 설명서에서 "
         "내 사업만 골라 검사하고 입력본을 만듭니다. 넣지 않으면 과 전체가 나옵니다.",
         "K-에듀파인 → 재정사업관리 → 사업담당 → 재정관리 → 예산관리 → 사업관리카드 → 세출 조회 → 파일다운로드",
         "과제카드세출예산요구목록.xlsx",
         [("엑셀", "*.xlsx")]),
    )
    FLOW = (("plan", "1 사업설명서"), ("last_year", "2 K-에듀파인"),
            ("ubis", "3 UBIS 검토조서"), ("classes", "4 사업별 분류표"),
            ("cards", "5 과제카드"))

    def _page_files(self, frame) -> None:
        self.paths = {key: tk.StringVar(value=self.settings["recent"].get(key, ""))
                      for key in ("plan", "ubis", "last_year", "classes", "cards")}
        self.file_view = FilePage(self, frame)
        self.file_marks = self.file_view.numbers
        # FilePage connects self.progress_box, self.year_note and the five FILE_CARDS.

    def _layout_cards(self, event=None) -> None:
        """Native rows resize with their content frame, including expanded source paths."""
        self.file_view._resized(event)

    def clear(self, key: str) -> None:
        """고른 파일을 뺀다.

        기억해 둔 값까지 같이 지운다. 화면에서만 지우면 다음에 켤 때 되살아나
        '분명히 뺐는데 또 있다'가 된다. 다른 해 자료로 갈아탈 때 꼭 필요하다.
        """
        if self.loading:
            return
        self.paths[key].set("")
        self.settings["recent"][key] = ""
        if key == "last_year":
            self.extra_years = []
            self.settings["extra_years"] = []
            self.year_note.configure(text="여러 부를 한 번에 고르면 비목이 더 많이 자동으로 채워집니다.",
                                     fg=MUTED)
        store.save(self.settings)
        self._mark_files()
        title = next((f"{card[1]}번 {card[2]}" for card in self.FILE_CARDS if card[0] == key), key)
        self.status.configure(text=f"{title}을(를) 뺐습니다.")

    def clear_all(self) -> None:
        if self.loading:
            return
        for key in list(self.paths):
            self.clear(key)
        self.status.configure(text="고른 자료를 모두 뺐습니다. 다시 골라 주세요.")

    def _mark_files(self) -> None:
        self.file_view.mark_files()

    def plan_is_supplement(self) -> bool:
        """1번에 고른 설명서가 추경 서식인가. 추경이면 2번(K-에듀파인 파일)이 필요 없다.

        파일을 여는 일이라 경로·수정 시각이 같으면 다시 보지 않는다.
        """
        path = self.paths["plan"].get().strip() if hasattr(self, "paths") else ""
        if not path or not Path(path).exists():
            return False
        key = (path, Path(path).stat().st_mtime)
        cached = getattr(self, "_supplement_cache", None)
        if cached and cached[0] == key:
            return cached[1]
        from supplement import is_supplement_file

        found = is_supplement_file(path)
        self._supplement_cache = (key, found)
        return found

    def pick(self, key: str, types) -> None:
        if self.loading:
            return
        self.file_view.close_guide()
        if key == "last_year":
            # 확정본은 여러 부일수록 비목이 더 많이 자동으로 채워진다.
            # 첫 번째 파일의 모양으로 입력본을 만들고, 나머지는 학습에만 쓴다.
            chosen = filedialog.askopenfilenames(title="지난 K-에듀파인 다운로드 (여러 부 선택 가능)",
                                                 filetypes=list(types) + [("모든 파일", "*.*")])
            if not chosen:
                return
            files = list(chosen)
            self.paths[key].set(files[0])
            self.extra_years = files[1:]
            self.year_note.configure(
                text=(f"{len(files)}부 선택 · 서식 본은 {Path(files[0]).name}" if len(files) > 1
                      else "한 부만 골랐습니다. 여러 부를 고르면 비목이 더 많이 채워집니다."),
                fg=OK if len(files) > 1 else MUTED)
            self._mark_files()
            return
        chosen = filedialog.askopenfilename(filetypes=list(types) + [("모든 파일", "*.*")])
        if chosen:
            self.paths[key].set(chosen)
            self._mark_files()

    # ----------------------------------------------------------------- 불러오기
    def load(self) -> None:
        if self.loading:
            return
        supplement = self.plan_is_supplement()
        needed = ("plan",) if supplement else ("plan", "last_year")
        if not all(self.paths[k].get().strip() for k in needed):
            self.show("files")
            messagebox.showwarning("필수 자료 확인", "1번 사업설명서와 2번 지난 K-에듀파인 파일을 선택해 주세요."
                                   if not self.paths["plan"].get().strip() or not supplement
                                   else "1번 사업설명서를 선택해 주세요.")
            return
        missing = [f"{card[1]}번 {card[2]} — {Path(self.paths[card[0]].get().strip()).name}"
                   for card in self.FILE_CARDS
                   if self.paths[card[0]].get().strip() and not Path(self.paths[card[0]].get().strip()).exists()]
        missing += [f"2번 추가 확정본 — {Path(one).name}" for one in self.extra_years if not Path(one).exists()]
        if missing:
            self.show("files")
            messagebox.showwarning("파일을 찾을 수 없습니다",
                                   "고른 파일이 그 자리에 없습니다. 옮기거나 지웠을 수 있습니다.\n\n - "
                                   + "\n - ".join(missing) + "\n\n해당 칸에서 다시 선택해 주세요.")
            return
        self.file_view.close_guide()
        self.loading = True
        self.file_view.mark_files()
        self.tutorial_button.configure(state="disabled")
        # 새 세션은 다 읽은 뒤에야 바꿔 끼운다. 먼저 바꿔 두면 읽다 실패했을 때 화면의
        # 지난 결과는 그대로인데 세션은 빈 것이 되어, 목록을 누르면 엉뚱한 사업이 뜬다.
        fresh = Session(plan_path=self.paths["plan"].get().strip(),
                        ubis_path=self.paths["ubis"].get().strip(),
                        last_year_path=self.paths["last_year"].get().strip(),
                        extra_last_year=list(self.extra_years),
                        class_path=self.paths["classes"].get().strip(),
                        card_path=self.paths["cards"].get().strip())
        self.pending_session = fresh
        self.show("files")
        self.progress_box.pack(fill="x", pady=(0, 14))
        self.progress["value"] = 0
        self.progress_text.configure(text="시작합니다…")

        def work():
            try:
                fresh.load(progress=lambda message, ratio: self.progress_queue.put(("step", message, ratio)))
                self.progress_queue.put(("done", "", 1.0))
            except Exception as error:  # noqa: BLE001 - 사용자에게 그대로 보여준다
                traceback.print_exc()
                self.progress_queue.put(("fail", friendly(error), 0.0))

        threading.Thread(target=work, daemon=True).start()
        self.after(60, self._pump)

    def _pump(self) -> None:
        try:
            while True:
                kind, message, ratio = self.progress_queue.get_nowait()
                if kind == "step":
                    self.progress["value"] = ratio * 100
                    self.progress_text.configure(text=f"{round(ratio * 100)}%  ·  {message}")
                elif kind == "fail":
                    self.loading = False
                    self.file_view.mark_files()
                    self.tutorial_button.configure(state="normal")
                    self.progress_box.pack_forget()
                    self.status.configure(text="불러오지 못했습니다. 지난 결과가 있으면 그대로 남아 있습니다.")
                    messagebox.showerror("불러오기 실패", message)
                    return
                else:
                    self.loading = False
                    self.session = self.pending_session
                    self.generated = ""
                    self.goal2_report = None
                    self.verify_report = None
                    self.current = 0
                    self.expanded.clear()
                    self.verify_tree.delete(*self.verify_tree.get_children())
                    self.verify_summary.configure(text=VERIFY_IDLE, fg=MUTED, font=(FONT, 10, "normal"))
                    self.file_view.mark_files()
                    self.tutorial_button.configure(state="normal")
                    self.progress_box.pack_forget()
                    self._after_load()
                    return
        except queue.Empty:
            pass
        self.after(60, self._pump)

    def _after_load(self) -> None:
        for key, variable in self.paths.items():
            self.settings["recent"][key] = variable.get().strip()
        self.settings["extra_years"] = list(self.extra_years)
        applied = self.session.apply_remembered(self.settings.get("bimok"))
        store.save(self.settings)
        self._mark_files()
        summary = self.session.summary()
        if self.session.is_supplement:
            self.subtitle.configure(text=f"{Path(self.session.plan_path).name}  ·  {summary['점검 종류']} · "
                                         f"사업 {summary['사업 수']}개 · 증감 줄 {summary['증감 줄']:,}개")
        else:
            self.subtitle.configure(text=f"{Path(self.session.plan_path).name}  ·  사업 {summary['사업 수']}개 · "
                                         f"산출근거 {summary['산출근거 행']:,}행"
                                         + (f" · 기억한 비목 {applied}건 적용" if applied else ""))
        self.refresh()
        self.set_filter("오류")
        self.show("result")
        has_errors = (self.session.plan_report.errors or (self.goal2_report and self.goal2_report.errors))
        self.status.configure(text="검사를 마쳤습니다. 빨간 줄(오류)부터 확인해 주세요."
                              if (self.session.blocking() and not self.session.is_supplement) or has_errors
                              else "검사를 마쳤습니다.")

    # ----------------------------------------------------------------- ② 검사 결과
    def _page_result(self, frame) -> None:
        cards = tk.Frame(frame, bg=GROUND)
        cards.pack(fill="x", padx=22, pady=(14, 0))
        self.cards = {}
        self.card_captions = {}
        for key, caption, color, target in (
            ("goal1", "설명서 검산 오류", ERROR, "goal1"),
            ("goal2", "UBIS 금액 대조 오류", ERROR, "goal2"),
            ("goal3", "비목 미확정", WARN, "bimok"),
            ("total", "입력할 요구액 (천원)", INK, "total"),
        ):
            card = tk.Frame(cards, bg=PANEL, highlightbackground=LINE, highlightthickness=1, cursor="hand2")
            card.pack(side="left", fill="both", expand=True, padx=(0, 8))
            strip = tk.Frame(card, bg=color, height=3)
            strip.pack(fill="x")
            inner = tk.Frame(card, bg=PANEL)
            inner.pack(fill="both", expand=True, padx=14, pady=10)
            caption_label = label(inner, caption, 9, False, MUTED, PANEL)
            caption_label.pack(anchor="w")
            self.card_captions[key] = (caption_label, caption)
            value = label(inner, "—", 18, True, color, PANEL)
            value.pack(anchor="w")
            note_label = label(inner, "", 8, False, MUTED, PANEL)
            note_label.pack(anchor="w")
            self.cards[key] = (value, note_label, strip)
            for widget in (card, inner, value, note_label):
                widget.bind("<Button-1>", lambda _e, t=target: self._card_click(t))

        tools = tk.Frame(frame, bg=GROUND)
        tools.pack(fill="x", padx=22, pady=(14, 8))
        self.filter_buttons = {}
        chips = tk.Frame(tools, bg=LINE)          # 1px 선으로 이어 붙인 한 덩어리 선택지
        chips.pack(side="left")
        # 2.15.0 — 목록에는 오류만 싣는다. 2.17.0 — 작성요령 점검은 '작성 점검'에 따로 모은다.
        # 틀렸다고 단정할 수 없는 것들이라 오류와 섞으면 정작 오류가 묻힌다.
        for name in ("오류", WRITING):
            chip = tk.Button(chips, text=name, font=(FONT, 9), relief="flat", bd=0,
                             padx=14, pady=5, cursor="hand2", highlightthickness=0,
                             activebackground=HEAD, command=lambda n=name: self.set_filter(n))
            chip.pack(side="left", padx=(0 if name == "오류" else 1, 0), pady=1)
            self.filter_buttons[name] = chip
        label(tools, "사업 찾기", 9, False, MUTED, GROUND).pack(side="left", padx=(16, 6))
        self.search = tk.Entry(tools, textvariable=self.query, font=(FONT, 10), relief="solid", bd=1,
                               highlightthickness=1, highlightcolor=ACCENT, highlightbackground=LINE)
        self.search.pack(side="left", fill="x", expand=True, ipady=4)
        self.query.trace_add("write", lambda *_a: self.refresh_list())
        button(tools, "지적 목록 엑셀로 저장", self.export).pack(side="left", padx=(10, 0))

        foot = tk.Frame(frame, bg=GROUND)
        foot.pack(side="bottom", fill="x", padx=22, pady=(0, 12))
        label(foot, "행을 두 번 누르거나 Enter: 상세 화면 · 금액 단위는 모두 천원", 9, False, MUTED,
              GROUND).pack(side="left")
        self.result_next = button(foot, "4 입력본 만들기로  →", lambda: self.show("make"), primary=True)
        self.result_next.pack(side="right")
        self.result_bimok = button(foot, "3 비목 확정으로", lambda: self.show("bimok"))
        self.result_bimok.pack(side="right", padx=8)

        # 고른 지적의 전체 문장. 표의 '내용' 칸은 잘려서, 무엇이 왜 틀렸는지 끝까지
        # 읽으려면 상세 화면을 열어야 했다. 고르기만 해도 아래에 전부 보인다.
        self.preview = tk.Frame(frame, bg=PANEL, highlightbackground=LINE, highlightthickness=1)
        self.preview.pack(side="bottom", fill="x", padx=22, pady=(0, 10))
        # 고른 지적의 원본 파일을 바로 여는 버튼. 목록에서 줄만 골라도 옆에 나온다.
        self.preview_actions = tk.Frame(self.preview, bg=PANEL)
        self.preview_actions.pack(side="right", anchor="n", padx=(0, 10), pady=8)
        self.preview_title = label(self.preview, "", 10, True, INK, PANEL, anchor="w")
        self.preview_title.pack(fill="x", padx=14, pady=(10, 2))
        self.preview_text = label(self.preview, "", 10, False, BODY, PANEL, anchor="w", justify="left",
                                  wraplength=1180)
        self.preview_text.pack(fill="x", padx=14, pady=(0, 10))
        # 오른쪽 '원본 열기' 버튼 자리만큼 글을 좁힌다. 안 그러면 문장 끝이 버튼 밑으로 잘린다.
        self.preview.bind("<Configure>", lambda e: self.preview_text.configure(
            wraplength=max(300, e.width - 50 - self.preview_actions.winfo_reqwidth())))

        holder, self.issue_tree = make_tree(
            frame,
            ("구분", "검사", "사업", "내용", "설명서 금액", "비교 금액", "차액"),
            (86, 96, 250, 0, 120, 120, 120),
            ("w", "w", "w", "w", "e", "e", "e"))
        holder.pack(fill="both", expand=True, padx=22, pady=(0, 8))
        self.issue_tree.bind("<Double-1>", lambda _e: self.open_detail())
        self.issue_tree.bind("<Return>", lambda _e: self.open_detail())
        self.issue_tree.bind("<<TreeviewSelect>>", lambda _e: self._preview())

    def _card_click(self, target: str) -> None:
        if target == "bimok" and self.session.is_supplement:
            self.query.set("")
            self.set_filter("전체")
            self._select_first(lambda goal, issue: goal == "추경연결")
            return
        if target == "bimok":
            self.show("bimok")
            return
        if target == "goal1":
            self.query.set("")
            self.set_filter("오류")
            self._select_first(lambda goal, issue: goal == "목표1" and issue.severity == "오류")
        elif target == "goal2":
            self.query.set("")
            self.set_filter("오류")
            self._select_first(lambda goal, issue: goal == "목표2A" and issue.severity == "오류")
        else:
            self.set_filter("전체")

    def _select_first(self, wanted) -> None:
        for index, (goal, issue) in enumerate(self.issues):
            if wanted(goal, issue):
                self._select(index)
                return

    def set_filter(self, name: str) -> None:
        self.filter = name
        self.current = 0
        self.refresh_list()

    def focus_search(self) -> None:
        self.show("result")
        self.search.focus_set()
        self.search.select_range(0, "end")

    # ----------------------------------------------------------------- 목록 채우기
    def refresh(self) -> None:
        if not self.session.projects:
            return
        summary = self.session.summary()
        # 입력본을 만든 뒤에는 그 파일까지 넣어 다시 맞춘다. 빈 문자열로 돌리면 방금
        # 만든 입력본과의 대조 결과가 화면을 새로 그릴 때마다 사라졌다.
        generated = self.generated if self.generated and Path(self.generated).exists() else ""
        goal2 = (self.session.verify_generated(generated)
                 if (self.session.ubis or generated) else None)
        self.goal2_report = goal2
        errors1 = len(self.session.plan_report.errors)
        blocking1 = len(self.session.plan_report.blocking_errors)
        errors2 = len(goal2.errors) if goal2 else 0
        unsettled = len(self.session.unsettled)
        skipped1 = len(self.session.plan_report.skipped)

        value, note, _strip = self.cards["goal1"]
        value.configure(text=f"{errors1:,} 건", fg=ERROR if errors1 else OK)
        note.configure(text=(f"그중 입력본을 막는 오류 {blocking1}건" if blocking1
                             else "입력본을 막는 오류 없음")
                            + (f" · 검사 못 한 줄 {skipped1}" if skipped1 else ""))
        value, note, _strip = self.cards["goal2"]
        if goal2:
            value.configure(text=f"{errors2:,} 건", fg=ERROR if errors2 else OK)
            note.configure(text="생성을 막지 않음"
                                + (f" · 대조 못 한 사업 {len(goal2.skipped)}" if goal2.skipped else ""))
        else:
            value.configure(text="대조 안 함", fg=FAINT)
            note.configure(text="3번 UBIS 검토조서를 넣지 않았습니다")
        supplement = self.session.is_supplement
        for widget, side, pad in ((self.result_next, "right", 0), (self.result_bimok, "right", 8)):
            if supplement and widget.winfo_manager():
                widget.pack_forget()
            elif not supplement and not widget.winfo_manager():
                widget.pack(side=side, padx=pad)
        for key, (caption_label, caption) in self.card_captions.items():
            caption_label.configure(text=SUPPLEMENT_CAPTIONS.get(key, caption) if supplement else caption)
        if supplement:
            self._refresh_supplement_cards(goal2)
        else:
            value, note, _strip = self.cards["goal3"]
            value.configure(text=f"{unsettled:,} 행" if unsettled else "모두 확정", fg=WARN if unsettled else OK)
            note.configure(text=(f"정할 것: {self.session.unsettled_kinds()}" if unsettled
                                 else f"확정 {summary['비목 확정']:,}행"))
            total = sum(project.base or 0 for project in self.session.projects if not project.handover)
            handed = sum(project.base or 0 for project in self.session.projects if project.handover)
            value, note, _strip = self.cards["total"]
            value.configure(text=money(total))
            note.configure(text=f"입력 대상 사업 {sum(1 for p in self.session.projects if not p.handover)}개"
                                + (f" · 재배정(입력 안 함) {money(handed)} 별도" if handed else ""))

        reasons = [] if supplement else self.session.blocking()
        soft = self.session.warnings()
        if supplement:
            self._banner_color(WARN_BG)
            errors = len(self.session.plan_report.errors) + (len(goal2.errors) if goal2 else 0)
            self.banner_text.configure(
                text=f"추경 점검 — {self.session.supplement.round_label or '추경'} 설명서입니다. 증감·산출식·합계와 "
                     "UBIS 추경 검토조서, 직전 차수 기정액을 맞춰 봤습니다. K-에듀파인에 입력한 뒤에는 "
                     "5 입력 후 대조에서 내려받은 파일을 맞춰 보세요." + (f"  오류 {errors}건." if errors else "  오류 없음."),
                fg="#6B4A08", bg=WARN_BG)
            self.banner_button.configure(text="첫 오류 보기 (F3)", command=self._first_error)
            self.banner.pack(fill="x", before=self.body)
        elif reasons:
            self._banner_color(ERROR_BG)
            self.banner_text.configure(text="입력본을 만들 수 없습니다 — " + " · ".join(reasons)
                                            + ".  설명서를 고친 뒤 다시 불러오세요.",
                                       fg=ERROR_INK, bg=ERROR_BG)
            self.banner_button.configure(text="첫 오류 보기 (F3)")
            self.banner.pack(fill="x", before=self.body)
        elif soft:
            self._banner_color(WARN_BG)
            self.banner_text.configure(text="입력본을 만들 수 있습니다. 다만 " + " · ".join(soft),
                                       fg="#6B4A08", bg=WARN_BG)
            self.banner_button.configure(text="3 비목 확정으로", command=lambda: self.show("bimok"))
            self.banner.pack(fill="x", before=self.body)
        else:
            self.banner.pack_forget()
        if reasons:
            self.banner_button.configure(command=self._first_error)
        self.refresh_list()
        if supplement:
            self._update_tabs()
            return
        self.refresh_bimok()
        self.refresh_make()
        self._update_tabs()

    def _refresh_supplement_cards(self, goal2) -> None:
        """추경 점검에서는 비목 카드 대신 직전 차수 연결, 요구액 카드 대신 증감 합계를 보인다."""
        document = self.session.supplement
        value, note, _strip = self.cards["goal1"]
        report = self.session.plan_report
        note.configure(text=(f"검사 못 한 것 {len(report.skipped)}건" if report.skipped
                             else "증감 · 산출식 · 합계 · 총괄 표"))
        value, note, _strip = self.cards["goal2"]
        if goal2 is not None and self.session.supplement_ubis is not None:
            upstream = [one for one in goal2.issues if one.goal != "추경연결"]
            errors = sum(1 for one in upstream if one.severity == "오류")
            value.configure(text=f"{errors:,} 건", fg=ERROR if errors else OK)
            note.configure(text="증감 = UBIS 추경요구액(B) · 기정 = 기정액(A)")
        value, note, _strip = self.cards["goal3"]
        chain = [one for one in (goal2.issues if goal2 else []) if one.goal == "추경연결"]
        if self.session.previous_used:
            problems = len(chain)
            value.configure(text=f"{problems:,} 건", fg=ERROR if problems else OK)
            note.configure(text="기정액(A) = 직전 차수가 끝난 금액")
        else:
            value.configure(text="대조 안 함", fg=FAINT)
            note.configure(text="UBIS 파일 옆에 직전 차수 검토조서가 없습니다")
        total = document.summary_total
        change = total.change if total and total.change is not None else sum(
            one.change or 0 for one in document.projects)
        value, note, _strip = self.cards["total"]
        value.configure(text=signed(change))
        note.configure(text=(f"추경 {money(total.after)} · 기정 {money(total.before)}" if total
                             else f"사업 쪽 {len(document.projects)}개의 증감 합계"))

    def _banner_color(self, color: str) -> None:
        self.banner.configure(bg=color)
        # RoundedButton 은 bg 를 버튼 면 색으로 받는다. 둥근 모서리 바깥(캔버스) 색은 따로 맞춘다.
        tk.Canvas.configure(self.banner_button, bg=color)

    def _first_error(self) -> None:
        self.show("result")
        self.query.set("")
        self.set_filter("오류")
        self._select_first(lambda goal, issue: issue.severity == "오류" and issue.blocking)

    def _all_issues(self) -> list:
        rows = [("목표1", issue) for issue in self.session.plan_report.issues]
        rows += [("목표4", issue) for issue in self.session.class_report.issues]
        writing = getattr(self.session, "writing_report", None)
        if writing is not None:
            rows += [(issue.goal, issue) for issue in writing.issues]
        if getattr(self, "goal2_report", None):
            rows += [(issue.goal if issue.goal == "추경연결" else "목표2A", issue)
                     for issue in self.goal2_report.issues]
        # 설명서에 같은 쪽이 두 번 실리면 UBIS 대조 문장도 글자 하나 다르지 않게 두 번
        # 나온다. 같은 문장은 한 번만 보인다(겹친 쪽 자체는 설명서 검산이 따로 말한다).
        unique = unique_issues(rows)
        order = {"오류": 0, "확인 필요": 1, "안내": 2}
        unique.sort(key=lambda pair: order.get(pair[1].severity, 3))
        return unique

    def refresh_list(self) -> None:
        tree = self.issue_tree
        tree.delete(*tree.get_children())
        needle = self.query.get().strip()
        self.issues = []
        everything = self._all_issues()
        for goal, issue in everything:
            if self.filter != "전체" and _kind(goal, issue) != self.filter:
                continue
            if needle and needle not in issue.project and needle not in self.session.label_of(issue.project):
                continue
            self.issues.append((goal, issue))
            tree.insert("", "end", values=(issue.severity, goal_label(goal), self.session.label_of(issue.project),
                                           issue.message, money(issue.left), money(issue.right), signed(issue.gap)),
                        tags=(SEVERITY_TAG.get(issue.severity, "info"),))
        counts = {"오류": 0, WRITING: 0}
        for goal, issue in everything:
            counts[_kind(goal, issue)] = counts.get(_kind(goal, issue), 0) + 1
        colors = {"오류": ERROR, WRITING: WARN}
        for name, chip in self.filter_buttons.items():
            on = name == self.filter
            total = sum(counts.values()) if name == "전체" else counts.get(name, 0)
            chip.configure(text=f"{name}  {total}",
                           bg=ACCENT if on else PANEL,
                           fg="#FFFFFF" if on else (colors[name] if total else FAINT),
                           font=(FONT, 9, "bold" if on else "normal"))
        if self.session.projects:
            skipped = len(self.session.plan_report.skipped) + (
                len(self.goal2_report.skipped) if getattr(self, "goal2_report", None) else 0)
            self.status.configure(text=f"{self.filter} {len(self.issues):,}건 표시 중 "
                                       f"(오류 {counts['오류']:,} · 작성 점검 {counts[WRITING]:,})"
                                       + (f" · '{needle}' 로 거름" if needle else "")
                                       + (f" · 검사하지 못한 것 {skipped}건(지적 목록 엑셀 '검사 못 한 것' 시트)"
                                          if skipped else ""))
        self._preview()

    def _preview(self) -> None:
        """고른 지적의 전체 문장을 표 아래에 보인다. 고른 게 없으면 목록 상태를 말한다."""
        selected = self.issue_tree.selection()
        if selected:
            index = self.issue_tree.index(selected[0])
            if 0 <= index < len(self.issues):
                self.current = index
                goal, issue = self.issues[index]
                where = f" · {issue.item}" if issue.item else ""
                self.preview_title.configure(
                    text=f"[{issue.severity}] {goal_label(goal)} · {self.session.label_of(issue.project)}{where}",
                    fg={"오류": ERROR, "확인 필요": WARN}.get(issue.severity, INK))
                self.preview_text.configure(text=issue.message)
                self._source_buttons(self.preview_actions, goal)
                self.update_idletasks()
                self.preview_text.configure(wraplength=max(
                    300, self.preview.winfo_width() - 50 - self.preview_actions.winfo_reqwidth()))
                return
        self._source_buttons(self.preview_actions, None)
        if not self.session.projects:
            title, body = "", ""
        elif not self.issues:
            title = f"'{self.filter}'에 해당하는 지적이 없습니다."
            body = ("찾는 글자를 지우거나 다른 구분을 눌러 보세요." if self.query.get().strip()
                    else "이 구분은 깨끗합니다. 왼쪽 위 다른 구분(오류 · 작성 점검)을 눌러 보세요.")
        else:
            title = f"{len(self.issues):,}건 — 줄을 고르면 여기에 전체 내용이 나옵니다."
            body = "F3 으로 다음 지적, Shift+F3 으로 이전 지적으로 옮겨 갑니다."
        self.preview_title.configure(text=title, fg=INK)
        self.preview_text.configure(text=body)

    # ----------------------------------------------------------------- 지적 이동
    def _select(self, index: int) -> None:
        children = self.issue_tree.get_children()
        if not children:
            return
        index = max(0, min(index, len(children) - 1))
        self.current = index
        self.issue_tree.selection_set(children[index])
        self.issue_tree.focus(children[index])
        self.issue_tree.see(children[index])

    def next_issue(self) -> None:
        if self.page == "detail":
            self._select(self.current + 1)
            self.open_detail()
            return
        self.show("result")
        self._select(self.current + 1 if self.issue_tree.selection() else 0)

    def previous_issue(self) -> None:
        if self.page == "detail":
            self._select(self.current - 1)
            self.open_detail()
            return
        self.show("result")
        self._select(self.current - 1)

    # ----------------------------------------------------------------- ③ 상세
    def _page_detail(self, frame) -> None:
        top = tk.Frame(frame, bg=PANEL)
        top.pack(fill="x")
        inner = tk.Frame(top, bg=PANEL)
        inner.pack(fill="x", padx=22, pady=12)
        button(inner, "← 목록  (Esc)", lambda: self.show("result")).pack(side="left")
        titles = tk.Frame(inner, bg=PANEL)
        titles.pack(side="left", padx=14)
        self.detail_badge = label(titles, "", 8, True, "#FFFFFF", ERROR, padx=8, pady=2)
        self.detail_title = label(titles, "", 13, True, INK, PANEL)
        self.detail_title.pack(anchor="w")
        self.detail_path = label(titles, "", 9, False, MUTED, PANEL)
        self.detail_path.pack(anchor="w")
        button(inner, "다음 지적 (F3)", self.next_issue).pack(side="right")
        button(inner, "이전 (Shift+F3)", self.previous_issue).pack(side="right", padx=8)
        self.detail_count = label(inner, "", 9, False, FAINT, PANEL)
        self.detail_count.pack(side="right", padx=10)
        tk.Frame(top, bg=LINE, height=1).pack(fill="x")

        numbers = tk.Frame(frame, bg=GROUND)
        numbers.pack(fill="x", padx=22, pady=(14, 10))
        self.detail_cards = []
        self.detail_captions = []
        for caption in ("설명서 금액 (천원)", "계산 · 비교 금액", "UBIS 요구액", "차액 (설명서 − 비교)"):
            card = tk.Frame(numbers, bg=PANEL, highlightbackground=LINE, highlightthickness=1)
            card.pack(side="left", fill="both", expand=True, padx=(0, 8))
            caption_label = label(card, caption, 9, False, MUTED, PANEL)
            caption_label.pack(anchor="w", padx=14, pady=(10, 0))
            self.detail_captions.append((caption_label, caption))
            value = label(card, "—", 15, True, INK, PANEL)
            value.pack(anchor="w", padx=14, pady=(0, 10))
            self.detail_cards.append(value)

        self.detail_note = label(frame, "", 10, False, BODY, PANEL, anchor="w", justify="left",
                                 wraplength=1180, padx=14, pady=10)
        self.detail_note.pack(fill="x", padx=22)

        finder = tk.Frame(frame, bg=PANEL, highlightbackground=LINE, highlightthickness=1)
        finder.pack(fill="x", padx=22, pady=(8, 0))
        inner = tk.Frame(finder, bg=PANEL)
        inner.pack(fill="x", padx=14, pady=10)
        label(inner, "한글에서 찾을 글자", 10, True, INK, PANEL).pack(side="left")
        self.find_value = tk.StringVar()
        entry = tk.Entry(inner, textvariable=self.find_value, font=("Consolas", 10),
                         relief="solid", bd=1, bg="#FBFCFD")
        entry.pack(side="left", fill="x", expand=True, padx=10, ipady=3)
        button(inner, "복사", self.copy_find).pack(side="left")
        self.detail_actions = tk.Frame(inner, bg=PANEL)
        self.detail_actions.pack(side="left", padx=(8, 0))
        label(frame, "'설명서 열기'를 누르면 찾을 글자를 복사한 뒤 파일을 엽니다. 열린 창에서 Ctrl+F → Ctrl+V → Enter "
                     "하면 그 자리로 갑니다. 한글 파일에는 쪽 번호가 없어 쪽수 대신 찾을 글자를 드립니다.",
              9, False, FAINT, GROUND).pack(anchor="w", padx=26, pady=(4, 0))

        legend = tk.Frame(frame, bg=GROUND)
        legend.pack(fill="x", padx=26, pady=(10, 0))
        label(legend, "표 색깔 : ", 9, False, MUTED, GROUND).pack(side="left")
        for text_, fore, back in (("이번 지적", ERROR_INK, ERROR_BG),
                                  ("비목 미확정 (금액과 무관)", "#6B4A08", WARN_BG),
                                  ("합계 일치", "#1E5C3B", OK_BG)):
            tk.Label(legend, text=f" {text_} ", font=(FONT, 8), fg=fore, bg=back,
                     padx=4, pady=1).pack(side="left", padx=(0, 8))
        label(legend, "색이 없는 줄은 이상 없음", 9, False, FAINT, GROUND).pack(side="left")

        holder, self.detail_tree = make_tree(
            frame,
            ("원문행", "레벨", "항목", "비목", "산출기초", "금액(천원)"),
            (64, 50, 0, 190, 250, 120),
            ("e", "c", "w", "w", "w", "e"))
        holder.pack(fill="both", expand=True, padx=22, pady=12)
        self.detail_tree.bind("<<TreeviewSelect>>", self._pick_find)

    def _pick_find(self, _event=None) -> None:
        """상세 표에서 줄을 고르면 그 줄을 찾을 글자로 바꾼다."""
        selected = self.detail_tree.focus()
        if not selected:
            return
        values = self.detail_tree.item(selected, "values")
        if len(values) > 4 and values[4]:
            self.find_value.set(values[4])
        elif len(values) > 2 and values[2]:
            self.find_value.set(str(values[2]).strip())

    def copy_find(self) -> None:
        text = self.find_value.get()
        if not text:
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        self.status.configure(text=f"복사했습니다 — 한글에서 Ctrl+F 에 붙여넣으세요: {text}")

    # ----------------------------------------------------------------- 원본 파일 열기
    # 지적을 보고 원본과 맞춰 보려면 예전에는 탐색기에서 파일을 찾아 열고, 찾을 글자를
    # 따로 복사해야 했다. 지적 옆 버튼 하나로 '찾을 글자 복사 + 파일 열기'를 한 번에 한다.
    # 한글·엑셀을 원격 조종해 그 줄로 옮기지는 않는다. 보안 승인·설치 판본에 따라 PC마다
    # 다르게 움직여서, 되는 PC와 안 되는 PC가 생기는 것보다 Ctrl+F 한 번이 확실하다.
    SOURCES = (("plan", "plan_path", "설명서"), ("ubis", "ubis_path", "UBIS 검토조서"),
               ("classes", "class_path", "분류표"), ("cards", "card_path", "과제카드 목록"))
    GOAL_SOURCES = {"목표2A": "ubis", "추경연결": "ubis", "목표4": "classes", "과제카드": "cards"}

    def _source_path(self, key: str) -> str:
        attribute = next(one[1] for one in self.SOURCES if one[0] == key)
        return (getattr(self.session, attribute, "") or "").strip()

    def _sources_for(self, goal) -> list:
        """이 지적을 맞춰 볼 파일들. 설명서는 늘, 비교 상대가 있으면 그 파일도."""
        if goal is None:
            return []
        keys = ["plan"] + ([self.GOAL_SOURCES[goal]] if goal in self.GOAL_SOURCES else [])
        return [key for key in keys if self._source_path(key)]

    def _source_buttons(self, holder, goal) -> None:
        for child in holder.winfo_children():
            child.destroy()
        for key in self._sources_for(goal):
            name = next(one[2] for one in self.SOURCES if one[0] == key)
            button(holder, f"{name} 열기", lambda k=key: self.open_source(k),
                   primary=key == "plan").pack(side="left", padx=(0, 6))

    def _find_for(self, key: str, issue) -> str:
        """원본 파일에서 Ctrl+F 로 찾을 글자. 설명서는 상세 화면과 같은 글자, UBIS 는 거기 적힌 사업명."""
        if key == "ubis":
            from crosscheck import ubis_name

            return ubis_name(self.session.ubis or {}, issue.project)
        if key != "plan":
            return issue.project
        if self.page == "detail" and self.find_value.get().strip():
            return self.find_value.get().strip()
        if str(getattr(issue, "goal", "")).startswith(WRITING_PREFIX) and issue.item:
            return issue.item
        if self.session.is_supplement and issue.item and issue.item not in ("총괄", "현황", "합계", "세부"):
            return issue.item
        project = next((one for one in self.session.projects if one.name == issue.project), None)
        return project.find_text() if project else issue.project

    def open_source(self, key: str) -> None:
        if not self.issues:
            return
        _goal, issue = self.issues[max(0, min(self.current, len(self.issues) - 1))]
        path = self._source_path(key)
        name = next(one[2] for one in self.SOURCES if one[0] == key)
        if not path or not Path(path).exists():
            messagebox.showwarning("파일을 찾을 수 없습니다",
                                   f"불러올 때 썼던 {name} 파일이 그 자리에 없습니다. 옮기거나 지웠을 수 있습니다."
                                   f"\n\n{path}")
            return
        find = self._find_for(key, issue)
        if find:
            self.clipboard_clear()
            self.clipboard_append(find)
        try:
            open_path(path)
        except OSError as error:
            messagebox.showerror("파일을 열지 못했습니다",
                                 f"{Path(path).name} 을(를) 열 프로그램을 찾지 못했습니다. "
                                 f"탐색기에서 직접 열어 주세요.\n\n(자세히: {error})")
            return
        self.status.configure(text=f"{name} 파일을 열었습니다 — 열린 창에서 Ctrl+F → Ctrl+V → Enter: {find}"
                              if find else f"{name} 파일을 열었습니다.")

    def open_detail(self) -> None:
        children = self.issue_tree.get_children()
        selected = self.issue_tree.selection()
        if selected:
            self.current = children.index(selected[0])
        if not self.issues:
            return
        self.current = max(0, min(self.current, len(self.issues) - 1))
        goal, issue = self.issues[self.current]
        project = next((one for one in self.session.projects if one.name == issue.project), None)

        self.detail_badge.configure(text=f" {issue.severity} ",
                                    bg={"오류": ERROR, "확인 필요": WARN}.get(issue.severity, MUTED))
        self.detail_title.configure(text=self.session.label_of(issue.project))
        if project is not None and self.session.classes is not None:
            from program_classes import describe

            path = describe(project, self.session.classes).replace("\n", "  ")
        else:
            path = " > ".join(x for x in [getattr(project, "policy", ""), getattr(project, "unit", ""),
                                          getattr(project, "program", "")] if x)
        self.detail_path.configure(text=f"{goal_label(goal)} · {path}"[:220]
                                        + (f" · {project.organisation}" if project else ""))
        self.detail_count.configure(text=f"{self.current + 1} / {len(self.issues)}")

        ubis = self.session.ubis.get(issue.project)
        values = [issue.left, issue.right, ubis, issue.gap]
        for card, value in zip(self.detail_cards, values):
            card.configure(text=money(value) if value is not None else "—")
        self.detail_cards[3].configure(text=signed(issue.gap), fg=gap_color(issue.gap))

        note = issue.message
        if project and issue.left is not None and ubis is not None and abs((issue.right or 0) - ubis) <= 1:
            note += "\n산출근거 합계와 UBIS 요구액이 같습니다. 설명서의 총사업비 표기만 다를 가능성이 큽니다."
        note += "\n도구는 값을 고치지 않습니다. K-에듀파인과 설명서에서 직접 확인해 주세요."
        self.detail_note.configure(text=note)

        self.find_value.set(project.find_text() if project else issue.project)
        if goal.startswith(WRITING_PREFIX) and issue.item:
            # 작성 점검은 걸린 글자 자체를 찾는다. 사업 제목으로 가면 그 사업 안을 다시 훑어야 한다.
            self.find_value.set(issue.item)
        self._source_buttons(self.detail_actions, goal)

        tree = self.detail_tree
        tree.delete(*tree.get_children())
        for (caption_label, caption), other in zip(self.detail_captions, (None, None, "UBIS 추경요구액(B)", None)):
            caption_label.configure(text=other if other and self.session.is_supplement else caption)
        if self.session.is_supplement:
            if issue.item and issue.item not in ("총괄", "현황", "합계", "세부"):
                self.find_value.set(issue.item)
            self._show_supplement(tree, issue)
            self.show("detail")
            return
        if project:
            target = issue_row(project.items, issue)
            marked = None
            for index, item in enumerate(project.items):
                code, _detail, _source = self.session.code_for(project, index)
                if index == target:
                    tag = "error"                       # 이번 지적이 가리키는 바로 그 줄
                elif item.is_leaf and not code:
                    tag = "warn"                        # 비목만 비었다. 금액과는 무관하다.
                elif item.depth <= 2:
                    tag = "strong"
                else:
                    tag = "child"
                row_id = tree.insert("", "end", values=(item.row, item.depth, "   " * (item.depth - 1) + item.name,
                                                        f"{item.bimok_name}({item.bimok_code5})"
                                                        if item.bimok_code5 else "",
                                                        item.formula, money(item.amount)), tags=(tag,))
                if index == target:
                    marked = row_id
            gap = None if project.base is None else project.base - project.item_total
            matched = gap is not None and abs(gap) <= 1
            note = ("합계가 올해 요구액과 같습니다" if matched
                    else f"차액 {signed(gap)}" if gap is not None else "총사업비를 찾지 못했습니다")
            tree.insert("", "end", values=("", "", "산출근거 합계", "",
                                           f"올해 요구액 {money(project.base)}  ·  {note}"
                                           + (f"   (총사업비 {money(project.request)})"
                                              if project.request is not None
                                              and project.request != project.base else ""),
                                           money(project.item_total)),
                        tags=("ok" if matched else "error",))
            self._show_outline(tree, project, issue)
            if marked is not None:
                # 사업이 길면 칠한 줄이 화면 아래에 있어 안 보인다. 그 줄까지 내려 준다.
                tree.see(marked)
        self.show("detail")

    def _show_supplement(self, tree, issue) -> None:
        """추경 사업의 증감 표를 펼친다. 지적이 가리키는 줄은 빨갛게."""
        document = self.session.supplement
        found = next((one for one in document.projects if one.name == issue.project), None)
        if found is None:
            return
        tree.insert("", "end", values=("", "", "사업 머리 표", "",
                                       f"추경안 {money(found.after)} · 기정 {money(found.before)}",
                                       signed(found.change)), tags=("strong",))
        for table in found.tables:
            title = table.section or ("미반영 사업" if table.kind == "미반영" else "증감 산출내역")
            wanted = issue.row if issue.block == table.order else 0
            cost = f"사업비 {signed(table.section_cost)}" if table.section_cost is not None else ""
            whole = issue.block == table.order and not issue.row
            tree.insert("", "end", values=("", "", f"― {title} ({table.kind} 표) ―", "", cost,
                                           signed(table.value("change"))), tags=("error" if whole else "strong",))
            for group in table.groups:
                hit = bool(wanted) and group.row == wanted
                tree.insert("", "end", values=(group.row, 1, group.title or group.business, "",
                                               f"{money(group.after)} − {money(group.before)}"
                                               + (f"  ({group.formula})" if group.formula else ""),
                                               signed(group.change)),
                            tags=("error" if hit else "strong",))
                for line in group.lines:
                    hit = bool(wanted) and line.row == wanted
                    tree.insert("", "end", values=(line.row, line.depth + 1, "   " * line.depth + line.name, "",
                                                   line.formula, signed(line.amount)),
                                tags=("error" if hit else "child",))
            if table.total is not None:
                hit = bool(wanted) and table.total.row == wanted
                tree.insert("", "end", values=(table.total.row, "", "합계", "",
                                               f"{money(table.total.after)} − {money(table.total.before)}",
                                               signed(table.total.change)), tags=("error" if hit else "ok",))

    def _show_outline(self, tree, project, issue) -> None:
        """지적이 사업계획의 '산출내역' 표를 가리키면 그 표도 아래에 펼친다.

        설명서는 같은 사업의 산출 내역을 두 서식으로 적는다. 지적이 뒤쪽 표에서 났는데
        화면에 앞쪽 표만 보이면, 담당자는 멀쩡한 줄들을 훑으며 어디가 틀렸는지 찾게 된다.
        틀린 줄이 있는 표를 보여 주는 것이 먼저다.
        """
        form = getattr(issue, "form", "")
        blocks = getattr(project, "outline_items", None) or []
        number = getattr(issue, "block", 0)
        if not form or number >= len(blocks):
            return
        tree.insert("", "end", values=("", "", f"― 사업계획 {form} ―", "",
                                       "같은 사업을 다시 적은 표입니다. 두 표의 금액은 같아야 합니다.", ""),
                    tags=("strong",))
        wanted = getattr(issue, "row", 0)
        for item in blocks[number]:
            if item.row and item.row == wanted:
                tag = "error"
            elif item.depth <= 2:
                tag = "strong"
            else:
                tag = "child"
            tree.insert("", "end", values=(item.row, item.depth,
                                           "   " * (item.depth - 1) + item.name, "",
                                           item.formula, money(item.amount)), tags=(tag,))

    # ----------------------------------------------------------------- ④ 비목
    def _page_bimok(self, frame) -> None:
        top = tk.Frame(frame, bg=GROUND)
        top.pack(fill="x", padx=22, pady=(14, 8))
        self.bimok_headline = label(top, "", 12, True, INK, GROUND)
        self.bimok_headline.pack(anchor="w")
        self.bimok_progress = ttk.Progressbar(top, style="Bar.Horizontal.TProgressbar",
                                              mode="determinate", maximum=100)
        self.bimok_progress.pack(fill="x", pady=6)
        self.bimok_note = label(top, "", 9, False, MUTED, GROUND, anchor="w", justify="left")
        self.bimok_note.pack(anchor="w")
        label(top, "설명서에 목코드(210-01 같은 앞 5자리)가 있으면 뒤 2자리만 정하면 됩니다. "
                   "목코드가 없는 줄은 항목명마다 7자리를 정합니다. 작년 기록이 없는 값은 추측해 채우지 않습니다.",
              9, False, MUTED, GROUND, anchor="w", justify="left", wraplength=1300).pack(anchor="w", pady=(2, 0))

        foot = tk.Frame(frame, bg=GROUND)
        foot.pack(side="bottom", fill="x", padx=22, pady=(0, 12))
        self.remember = tk.BooleanVar(value=True)
        tk.Checkbutton(foot, text="정한 비목을 기억해 다음 실행에도 씁니다", variable=self.remember,
                       font=(FONT, 9), bg=GROUND, fg=BODY, activebackground=GROUND,
                       highlightthickness=0, selectcolor=PANEL).pack(side="left")
        button(foot, "4 입력본 만들기로  →", lambda: self.show("make"), primary=True).pack(side="right")
        button(foot, "다음 미확정 줄", self._next_unsettled).pack(side="right", padx=8)

        legend = tk.Frame(frame, bg=GROUND)
        legend.pack(side="bottom", fill="x", padx=26, pady=(0, 8))
        for text_, fore, back in (("미확정 — 정해 주세요", "#6B4A08", WARN_BG),
                                  ("항목명마다 다름 — 펼쳐서 정하세요", ERROR_INK, ERROR_BG),
                                  ("자동 확정 (작년 기록)", "#1E5C3B", OK_BG)):
            tk.Label(legend, text=f" {text_} ", font=(FONT, 8), fg=fore, bg=back,
                     padx=4, pady=1).pack(side="left", padx=(0, 8))
        label(legend, "줄을 두 번 누르거나 Enter 로 정합니다 · ▸ 줄은 → 로 펼칩니다", 9, False, MUTED,
              GROUND).pack(side="left", padx=(6, 0))

        holder, self.bimok_tree = make_tree(
            frame,
            ("비목 · 목코드", "행 수", "항목명", "원가통계비목", "근거", "상태"),
            (230, 64, 0, 110, 250, 104),
            ("w", "e", "w", "w", "w", "w"))
        holder.pack(fill="both", expand=True, padx=22, pady=(8, 8))
        self.bimok_empty = label(frame, "", 11, False, OK, PANEL)
        self.bimok_tree.bind("<Double-1>", lambda _e: self.edit_code())
        self.bimok_tree.bind("<Return>", lambda _e: self.edit_code())
        self.bimok_tree.bind("<Right>", lambda _e: self.expand(True))
        self.bimok_tree.bind("<Left>", lambda _e: self.expand(False))
        self.expanded: set = set()

    def refresh_bimok(self) -> None:
        if not self.session.projects:
            return
        groups = self.session.groups()
        self.groups = groups
        tree = self.bimok_tree
        tree.delete(*tree.get_children())
        done = sum(group.settled_rows for group in groups)
        rows = sum(group.rows for group in groups)
        left = rows - done
        if left:
            self.bimok_headline.configure(text=f"비목 미확정 {left:,}행 — 정할 것: {self.session.unsettled_kinds()}",
                                          fg=INK)
        else:
            self.bimok_headline.configure(text="비목이 모두 정해졌습니다" if rows else "정할 비목이 없습니다",
                                          fg=OK)
        self.bimok_progress["value"] = (done / rows * 100) if rows else 100
        split = sum(1 for g in groups if g.split and not g.no_code)
        self.bimok_note.configure(
            text=f"확정 {done:,} / {rows:,}행 ({done * 100 // max(rows, 1)}%)"
                 + (f" · 항목명에 따라 갈리는 목코드 {split}종" if split else "")
                 + " · 재배정(입력 안 함) 줄은 여기 나오지 않습니다")
        if not groups:
            self.bimok_empty.configure(text="입력본에 들어갈 산출근거 줄이 없습니다.")
            self.bimok_empty.place(relx=0.5, rely=0.45, anchor="center")
        else:
            self.bimok_empty.place_forget()

        for index, group in enumerate(groups):
            opened = group.no_code or group.split or group.code5 in self.expanded
            mark = "▾" if opened else "▸"
            if group.no_code:
                left_rows = group.rows - group.settled_rows
                state = "모두 확정" if group.settled else f"{left_rows:,}행 미확정"
                tag = "ok" if group.settled else "warn"
                names_text = f"항목명 {group.name_count:,}종 — 아래에서 하나씩 정합니다"
                code_text, source_text = "", "설명서에 목코드가 없어 작년 확정본에서 항목명으로 찾았습니다"
                title = f"{mark} 목코드 없음 (산출내역 표에서 읽은 줄)"
            else:
                state = "자동 확정" if group.settled and not group.split else ("항목명마다 다름" if group.split else "미확정")
                tag = "ok" if group.settled and not group.split else ("error" if group.split else "warn")
                names = " · ".join(sorted({name for sub in group.subgroups for name in sub.names})[:3])
                extra = group.name_count - 3
                names_text = names + (f"  외 {extra}종" if extra > 0 else "")
                code_text = group.code or ("항목명마다 다름" if group.split else "")
                source_text = group.source if not group.split else "작년엔 항목명마다 달랐습니다"
                title = f"{mark} {group.bimok_label}  {group.code5}"
            tree.insert("", "end", iid=f"g{index}",
                        values=(title, f"{group.rows:,}", names_text, code_text, source_text, state), tags=(tag,))
            if not opened:
                continue
            for order, sub in enumerate(group.subgroups):
                tree.insert("", "end", iid=f"s{index}.{order}",
                            values=("     └ 항목명", f"{sub.rows:,}", sub.label,
                                    sub.code or "미확정",
                                    sub.source + (f"  후보 {', '.join(sub.candidates)}" if sub.candidates else ""),
                                    "자동 확정" if sub.settled else "미확정"),
                            tags=("child" if sub.settled else "warn",))

    def expand(self, open_it: bool) -> None:
        selected = self.bimok_tree.focus()
        if not selected.startswith("g"):
            return
        group = self.groups[int(selected[1:])]
        if open_it:
            self.expanded.add(group.code5)
        else:
            self.expanded.discard(group.code5)
        self.refresh_bimok()
        if self.bimok_tree.exists(selected):
            self.bimok_tree.selection_set(selected)
            self.bimok_tree.focus(selected)

    def edit_code(self) -> None:
        selected = self.bimok_tree.focus()
        if not selected:
            return
        if selected.startswith("g"):
            group = self.groups[int(selected[1:])]
            if group.no_code:
                messagebox.showinfo("항목명마다 정합니다",
                                    "설명서에 목코드가 없는 줄들입니다. 앞 5자리를 모르므로 한 값으로 "
                                    "일괄 확정하지 않습니다.\n\n아래 '└ 항목명' 줄을 골라 하나씩 정해 주세요.")
                return
            if group.split:
                self.expanded.add(group.code5)
                self.refresh_bimok()
                messagebox.showinfo("항목명마다 다릅니다",
                                    f"{group.code5} 는 작년 확정본에서 항목명에 따라 다른 통계목을 썼습니다.\n"
                                    "아래 항목명 줄에서 각각 정해 주세요.")
                return
            target, by_name, current = group, False, group.code
        else:
            index, order = selected[1:].split(".")
            group = self.groups[int(index)]
            target, by_name = group.subgroups[int(order)], True
            current = target.code
        choices = self.session.known_codes() if group.no_code else []
        answer = AskCode(self, group, target, current, choices).result
        if answer is None:
            return
        if by_name:
            for one in target.names:
                self.session.overrides_by_name[(group.code5, one)] = answer
        else:
            self.session.overrides_by_code[group.code5] = answer
        if self.remember.get():
            self.settings["bimok"] = self.session.remembered()
            store.save(self.settings)
        self.refresh()
        self.show("bimok")
        self._next_unsettled()
        self.status.configure(text=f"{target.rows:,}행에 {answer} 를 넣었습니다. 남은 미확정 {len(self.session.unsettled):,}행.")

    def _next_unsettled(self) -> None:
        for item in self.bimok_tree.get_children():
            values = self.bimok_tree.item(item, "values")
            if item.startswith("g") and self.groups[int(item[1:])].no_code:
                continue          # 묶음 줄은 정할 수 없다. 그 아래 항목명 줄로 간다
            if values[5] in ("미확정", "항목명마다 다름"):
                self.bimok_tree.selection_set(item)
                self.bimok_tree.focus(item)
                self.bimok_tree.see(item)
                self.bimok_tree.focus_set()
                return

    # ----------------------------------------------------------------- ⑤ 생성 · 검증
    def _page_make(self, frame) -> None:
        box = tk.Frame(frame, bg=GROUND)
        box.pack(fill="both", expand=True, padx=22, pady=18)
        label(box, "입력본 만들기", 14, True, INK, GROUND).pack(anchor="w")
        label(box, "만들기 전에 아래를 확인합니다. 빨간 항목이 하나라도 있으면 파일을 만들지 않습니다 — "
                   "완성돼 보이는 파일을 경고와 함께 내주면 그대로 입력되기 때문입니다.",
              10, False, MUTED, GROUND, justify="left", anchor="w", wraplength=1100).pack(anchor="w", pady=(4, 12))
        self.checklist = tk.Frame(box, bg=PANEL, highlightbackground=LINE, highlightthickness=1)
        self.checklist.pack(fill="x")
        self.make_note = label(box, "", 10, False, BODY, GROUND, justify="left", anchor="w", wraplength=1100)
        self.make_note.pack(anchor="w", pady=(12, 4))
        row = tk.Frame(box, bg=GROUND)
        row.pack(anchor="w", pady=8)
        self.make_button = button(row, "입력본 만들기  (Ctrl+S)", self.save, primary=True)
        self.make_button.pack(side="left")
        label(box, "만든 파일의 시트: 입력본(K-에듀파인 화면 순서) · 대본(한 줄씩 확인용, 빨간 줄 = 비목 미확정) · "
                   "재배정(입력 안 함)", 9, False, MUTED, GROUND, justify="left").pack(anchor="w", pady=(6, 0))

    def refresh_make(self) -> None:
        """입력본을 만들기 전 점검표. 무엇이 막고 무엇이 막지 않는지 한 곳에서 본다."""
        if not hasattr(self, "checklist"):
            return
        for child in self.checklist.winfo_children():
            child.destroy()
        if not self.session.projects:
            return
        checks = []
        blocking1 = self.session.plan_report.blocking_errors
        checks.append(("error" if blocking1 else "ok", "설명서 금액 검산",
                       f"입력본 값을 틀리게 만드는 오류 {len(blocking1)}건 — 설명서를 고친 뒤 다시 불러오세요"
                       if blocking1 else "통과"))
        if self.session.classes is not None:
            wrong = self.session.class_report.errors
            checks.append(("error" if wrong else "ok", "사업 분류 (4번 자료)",
                           f"분류 오류 {len(wrong)}건" if wrong else "통과"))
        unsettled = len(self.session.unsettled)
        checks.append(("warn" if unsettled else "ok", "원가통계비목",
                       f"{unsettled:,}행이 빈칸으로 나갑니다 ({self.session.unsettled_kinds()}). "
                       "만들 수는 있지만 K-에듀파인에서 직접 골라야 합니다" if unsettled else "모두 확정"))
        goal2 = self.goal2_report
        if self.session.ubis:
            errors2 = len(goal2.errors) if goal2 else 0
            checks.append(("warn" if errors2 else "ok", "UBIS 금액 대조",
                           f"금액이 다른 사업 {errors2}건 — 생성은 막지 않습니다. 어느 쪽이 맞는지 확인하세요"
                           if errors2 else "모든 사업 금액 일치"))
        else:
            checks.append(("info", "UBIS 금액 대조", "3번 자료가 없어 건너뜀"))
        icons = {"ok": ("✓", OK), "warn": ("!", WARN), "error": ("✕", ERROR), "info": ("–", MUTED)}
        for index, (kind, name, detail) in enumerate(checks):
            line = tk.Frame(self.checklist, bg=PANEL)
            line.pack(fill="x", padx=16, pady=(12 if index == 0 else 4, 12 if index == len(checks) - 1 else 4))
            mark, color = icons[kind]
            tk.Label(line, text=mark, font=(FONT, 12, "bold"), fg=color, bg=PANEL, width=2).pack(side="left")
            label(line, name, 10, True, INK, PANEL, width=16, anchor="w").pack(side="left", padx=(4, 8))
            label(line, detail, 10, False, color if kind in ("error", "warn") else BODY, PANEL,
                  anchor="w", justify="left").pack(side="left", fill="x", expand=True)
        blocked = bool(self.session.blocking())
        self.make_button.configure(state="disabled" if blocked else "normal")
        if self.generated:
            self.make_note.configure(text=f"마지막으로 만든 파일: {self.generated}", fg=OK)
        elif blocked:
            self.make_note.configure(text="빨간 항목을 해결해야 만들 수 있습니다.", fg=ERROR)
        else:
            self.make_note.configure(text="준비됐습니다. 저장할 곳을 고르면 파일을 만듭니다.", fg=BODY)

    def _page_verify(self, frame) -> None:
        box = tk.Frame(frame, bg=GROUND)
        box.pack(fill="x", padx=22, pady=(16, 8))
        label(box, "입력 후 대조", 14, True, INK, GROUND).pack(anchor="w")
        label(box, "K-에듀파인에 입력을 마친 뒤 내려받은 파일을 설명서·UBIS·만든 입력본과 한 줄씩 맞춰 봅니다.",
              10, False, MUTED, GROUND, anchor="w").pack(anchor="w", pady=(4, 0))
        row = tk.Frame(box, bg=GROUND)
        row.pack(fill="x", pady=10)
        button(row, "내려받은 파일 고르고 대조", self.verify, primary=True).pack(side="left")
        self.verify_summary = label(row, VERIFY_IDLE, 10, False, MUTED, GROUND, anchor="w")
        self.verify_summary.pack(side="left", padx=14)
        holder, self.verify_tree = make_tree(
            frame,
            ("구분", "사업 · 행", "내용", "설명서 · 입력본", "비교 금액", "차액"),
            (86, 268, 0, 132, 132, 132),
            ("w", "w", "w", "e", "e", "e"))
        holder.pack(fill="both", expand=True, padx=22, pady=(0, 14))

    def save(self) -> None:
        if self.loading:
            return
        if not self.session.projects:
            messagebox.showwarning("먼저 불러오기", "1단계에서 자료를 고르고 F5 로 먼저 불러와 주세요.")
            return
        reasons = self.session.blocking()
        if reasons:
            self.show("result")
            messagebox.showwarning("만들 수 없습니다", "다음을 해결해야 입력본을 만들 수 있습니다.\n\n - "
                                   + "\n - ".join(reasons)
                                   + "\n\n금액과 분류가 틀린 채로 나간 파일은 눈으로 알아챌 수 없습니다.")
            self._first_error()
            return
        soft = self.session.warnings()
        if soft and not messagebox.askyesno(
                "비목 미확정",
                "\n".join(soft) + "\n\n비목 칸을 비운 채로 만들까요?\n"
                "'대본' 시트에서 빨간 줄이 그 행입니다.\n\n"
                "[아니요]를 누르면 3 비목 확정 화면으로 갑니다."):
            self.show("bimok")
            self._next_unsettled()
            return
        target = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                              initialfile="K-에듀파인_입력본.xlsx",
                                              filetypes=[("엑셀", "*.xlsx")])
        if not target:
            return
        try:
            self.session.save(target, allow_unsettled=True)
        except PermissionError:
            messagebox.showerror("저장할 수 없습니다",
                                 f"{Path(target).name} 이(가) 엑셀에서 열려 있는 것 같습니다.\n"
                                 "엑셀을 닫거나 다른 이름으로 저장해 주세요.")
            return
        except Exception as error:  # noqa: BLE001
            traceback.print_exc()
            messagebox.showerror("만들지 못했습니다", friendly(error))
            return
        self.generated = target
        self.refresh()
        self.show("make")
        self.status.configure(text=f"입력본을 만들었습니다: {target}")
        messagebox.showinfo("만들었습니다",
                            f"{Path(target).name}\n\n'입력본' 시트를 보며 K-에듀파인에 입력하세요.\n"
                            "'대본' 시트는 한 줄씩 확인할 때 씁니다.\n\n"
                            "입력을 마치면 5단계에서 내려받은 파일을 대조하세요.")

    def verify(self) -> None:
        if self.loading:
            return
        if not self.session.projects:
            messagebox.showwarning("먼저 불러오기", "1단계에서 자료를 고르고 F5 로 먼저 불러와 주세요.")
            return
        if self.session.is_supplement:
            # 추경은 사업(카드)마다 내려받으므로 여러 부를 한 번에 맞춘다.
            chosen = list(filedialog.askopenfilenames(title="K-에듀파인 추경 세출예산요구내역 (여러 부 선택 가능)",
                                                      filetypes=[("엑셀", "*.xlsx"), ("모든 파일", "*.*")]))
        else:
            one = filedialog.askopenfilename(title="K-에듀파인에서 내려받은 파일",
                                             filetypes=[("엑셀", "*.xlsx"), ("모든 파일", "*.*")])
            chosen = [one] if one else []
        if not chosen:
            return
        from crosscheck import Report

        report = Report()
        try:
            for downloaded in chosen:
                report.extend(self.session.verify_downloaded(downloaded, self.generated))
        except Exception as error:  # noqa: BLE001
            traceback.print_exc()
            messagebox.showerror("대조하지 못했습니다", friendly(error))
            return
        label_ = Path(chosen[0]).name if len(chosen) == 1 else f"{len(chosen)}부"
        self.show_verify(report, label_)
        messagebox.showinfo("대조 완료", f"오류 {len(report.errors)}건" if report.errors else "오류가 없습니다.")

    def show_verify(self, report, downloaded: str = "") -> None:
        tree = self.verify_tree
        tree.delete(*tree.get_children())
        shown = [issue for _goal, issue in unique_issues(("목표2B", one) for one in report.issues)]
        for issue in shown:
            tree.insert("", "end", values=(issue.severity, self.session.label_of(issue.project), issue.message,
                                           money(issue.left), money(issue.right), signed(issue.gap)),
                        tags=(SEVERITY_TAG.get(issue.severity, "info"),))
        self.verify_report = report
        errors = len(report.errors)
        name = Path(downloaded).name if downloaded else ""
        skipped = f" · 계산하지 못한 산출기초 {len(report.skipped)}건" if report.skipped else ""
        if errors:
            text, color = f"오류 {errors}건{skipped}  ({name})", ERROR
        else:
            text, color = f"✓ 오류가 없습니다{skipped}  ({name})", OK
        self.verify_summary.configure(text=text, fg=color, font=(FONT, 10, "bold"))
        if not self.generated and not self.session.is_supplement:
            self.verify_summary.configure(text=text + "  · 이번 실행에서 만든 입력본이 없어 줄 단위 비교는 건너뜀")
        self.show("verify")
        self.status.configure(text=f"대조했습니다: {downloaded}" if downloaded else "대조했습니다.")

    def export(self) -> None:
        if self.loading:
            return
        if not self.session.projects:
            messagebox.showwarning("먼저 불러오기", "1단계에서 자료를 고르고 F5 로 먼저 불러와 주세요.")
            return
        target = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                              initialfile="예산검사_지적목록.xlsx",
                                              filetypes=[("엑셀", "*.xlsx")])
        if not target:
            return
        # 화면 목록에 나온 것은 파일에도 다 있어야 한다. 예전에는 사업 분류(목표4) 지적이
        # 화면에만 있고 파일에서 빠졌다.
        reports = {"목표1": self.session.plan_report}
        if self.session.class_report.issues:
            reports["목표4"] = self.session.class_report
        if getattr(self, "goal2_report", None):
            from crosscheck import Report

            upstream, chain = Report(), Report()
            for issue in self.goal2_report.issues:
                (chain if issue.goal == "추경연결" else upstream).issues.append(issue)
            reports["목표2A"] = upstream
            if chain.issues:
                reports["추경연결"] = chain
        if getattr(self, "verify_report", None):
            reports["목표2B"] = self.verify_report
        try:
            export_issues(target, self.session, reports, getattr(self.session, "writing_report", None))
        except PermissionError:
            messagebox.showerror("저장할 수 없습니다", f"{Path(target).name} 이(가) 엑셀에서 열려 있는 것 같습니다.")
            return
        except Exception as error:  # noqa: BLE001
            traceback.print_exc()
            messagebox.showerror("저장 실패", str(error))
            return
        messagebox.showinfo("저장 완료", Path(target).name)


class AskCode(tk.Toplevel):
    """원가통계비목을 받는 작은 창.

    목코드가 있으면 앞 5자리는 고정이고 뒤 2자리만 받는다. 목코드가 없는 줄은 7자리를
    다 받되, 작년 확정본에 실제로 있던 코드 목록에서 고르게 한다. 외워서 치면 오타가
    그대로 입력본에 실린다.
    """

    def __init__(self, parent, group, target, current: str, choices=None) -> None:
        super().__init__(parent)
        self.title("원가통계비목 정하기")
        self.configure(bg=PANEL)
        self.resizable(False, False)
        self.result = None
        self.full = not group.digits
        digits = group.digits
        heading = "목코드 없음 — 7자리를 정합니다" if self.full else f"{group.bimok_label}  ({group.code5})"
        label(self, heading, 12, True, INK, PANEL).pack(anchor="w", padx=18, pady=(16, 2))
        label(self, target.label if hasattr(target, "label") else "", 10, False, BODY, PANEL,
              wraplength=440, justify="left").pack(anchor="w", padx=18)
        parents = getattr(target, "parents", [])
        spread = (f" 상위 항목 {len(parents)}곳: " + ", ".join(parents[:4])
                  + (f" 외 {len(parents) - 4}곳" if len(parents) > 4 else "")) if len(parents) > 1 else ""
        label(self, f"{target.rows:,}행에 한꺼번에 들어갑니다.{spread}", 9, False, MUTED, PANEL,
              wraplength=440, justify="left").pack(anchor="w", padx=18, pady=(6, 0))
        if len(parents) > 1:
            label(self, "상위 항목마다 비목이 다르면 여기서 한 값으로 정하지 말고 비워 둔 뒤 K-에듀파인에서 고르세요.",
                  9, False, WARN, PANEL, wraplength=440, justify="left").pack(anchor="w", padx=18, pady=(2, 0))

        row = tk.Frame(self, bg=PANEL)
        row.pack(anchor="w", padx=18, pady=12)
        if not self.full:
            label(row, digits, 15, True, MUTED, PANEL).pack(side="left")
        self.entry = tk.Entry(row, font=("Consolas", 15), width=9 if self.full else 4, justify="center",
                              relief="solid", bd=1, highlightthickness=1, highlightcolor=ACCENT)
        if self.full:
            self.entry.insert(0, current if len(current) == 7 else "")
        else:
            self.entry.insert(0, current[5:] if len(current) == 7 else "")
        self.entry.pack(side="left", padx=6, ipady=3)
        self.entry.focus_set()
        self.entry.select_range(0, "end")
        candidates = [one for one in getattr(target, "candidates", []) if len(one) == 7][:4]
        if candidates:
            label(row, "작년 후보:", 9, False, MUTED, PANEL).pack(side="left", padx=(8, 2))
        for candidate in candidates:
            shown = candidate if self.full else candidate[5:]
            tk.Button(row, text=shown, font=(FONT, 9), relief="solid", bd=1, padx=8, bg=PANEL,
                      cursor="hand2", command=lambda value=shown: self._set(value)).pack(side="left", padx=3)

        self.choices = list(choices or [])
        if self.full and self.choices:
            # 작년에 실제로 쓴 코드 목록. 비목명으로 찾아 고르면 칸에 들어간다.
            picker = tk.Frame(self, bg=PANEL)
            picker.pack(fill="x", padx=18)
            label(picker, "작년 확정본에 있던 코드에서 고르기 (비목명·숫자로 찾기)", 9, False, MUTED,
                  PANEL).pack(anchor="w")
            self.find = tk.StringVar()
            find = tk.Entry(picker, textvariable=self.find, font=(FONT, 10), relief="solid", bd=1)
            find.pack(fill="x", pady=(4, 4), ipady=2)
            self.listbox = tk.Listbox(picker, height=7, font=("Consolas", 10), activestyle="none",
                                      relief="solid", bd=1, highlightthickness=0,
                                      selectbackground="#D6E4F0", selectforeground=INK)
            self.listbox.pack(fill="x")
            self.find.trace_add("write", lambda *_a: self._filter())
            self.listbox.bind("<<ListboxSelect>>", lambda _e: self._pick())
            self.listbox.bind("<Double-1>", lambda _e: self._done())
            self._filter()

        note = ("앞 5자리를 모르는 줄이라 7자리 전부가 필요합니다. 모르면 비워 두고 K-에듀파인에서 고르세요."
                if self.full else "작년 확정본에 기록이 없으면 담당자가 정해야 합니다. 추측해서 채우지 않습니다.")
        label(self, note, 9, False, MUTED, PANEL, wraplength=440, justify="left").pack(anchor="w", padx=18,
                                                                                       pady=(10, 0))

        buttons = tk.Frame(self, bg=PANEL)
        buttons.pack(anchor="e", padx=18, pady=14)
        button(buttons, "취소", self.destroy).pack(side="right", padx=4)
        button(buttons, "정하기 (Enter)", self._done, primary=True).pack(side="right", padx=4)
        self.entry.bind("<Return>", lambda _e: self._done())
        self.bind("<Escape>", lambda _e: self.destroy())
        self.digits = digits
        self.transient(parent)
        # 부모 창 가운데에 띄운다. 창 관리자에 맡기면 화면 구석에 떠서 못 보고 지나친다.
        self.update_idletasks()
        x = parent.winfo_rootx() + max(0, (parent.winfo_width() - self.winfo_reqwidth()) // 2)
        y = parent.winfo_rooty() + max(0, (parent.winfo_height() - self.winfo_reqheight()) // 3)
        self.geometry(f"+{x}+{y}")
        self.grab_set()
        parent.wait_window(self)

    def _filter(self) -> None:
        needle = self.find.get().strip()
        self.listbox.delete(0, "end")
        self.shown = [(code, name) for code, name in self.choices
                      if not needle or needle in code or needle in name]
        for code, name in self.shown:
            self.listbox.insert("end", f"{code}  {name}")

    def _pick(self) -> None:
        chosen = self.listbox.curselection()
        if chosen:
            self._set(self.shown[chosen[0]][0])

    def _set(self, value: str) -> None:
        self.entry.delete(0, "end")
        self.entry.insert(0, value)

    def _done(self) -> None:
        tail = self.entry.get().strip()
        if self.full:
            if not tail.isdigit() or len(tail) != 7:
                messagebox.showwarning("일곱 자리 숫자", "원가통계비목은 7자리 숫자입니다. 예: 2100143", parent=self)
                return
            self.result = tail
        else:
            if not tail.isdigit() or len(tail) != 2:
                messagebox.showwarning("두 자리 숫자", "통계목은 두 자리 숫자입니다. 예: 43", parent=self)
                return
            self.result = self.digits + tail
        self.destroy()


def main() -> None:
    App().mainloop()


if __name__ == "__main__":
    main()
