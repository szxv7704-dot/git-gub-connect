"""첫 실행 안내. 이 프로그램이 무엇을 하는지, 무엇을 넣어야 하는지 먼저 말한다.

만든 사람은 안다. 처음 켠 담당자는 모른다. 파일 고르는 칸 네 개만 덩그러니
보여 주면 '무엇을 왜 넣는지'를 스스로 추측해야 하고, 그러면 선택 자료를 빼놓고
돌린 뒤 "검사가 부실하다"고 생각하게 된다. 목적을 먼저 말하는 편이 싸다.

한 번에 한 가지만 말한다. 1쪽 '무엇을 하는가', 2쪽 '무엇을 넣는가',
3쪽 '어떻게 진행되는가'. 세 쪽을 넘기면 끝이다.
"""

from __future__ import annotations

import tkinter as tk

from ui_common import ACCENT, BODY, FONT, GROUND, INK, LINE, MUTED, OK_BG, PANEL

# (쪽 제목, 한 줄 요약, [(머리, 본문, 덧말), …], 맺음말)
PAGES = [
    (
        "옮겨 적다 생기는 실수를 막습니다",
        "K-에듀파인에는 엑셀 업로드가 없어 칸마다 더블클릭해 직접 칩니다. "
        "그 과정에서 숫자 하나가 틀어져도 아무도 모릅니다. 이 프로그램이 세 가지를 봅니다.",
        [
            ("목적 1 · 설명서 안에서 금액이 맞는가",
             "산출식을 직접 계산해 적힌 금액과 비교합니다. 하위 항목을 더한 값이 상위 금액과 "
             "같은지, 산출근거를 모두 더한 값이 그 사업 요구액과 같은지도 봅니다.",
             "곱셈이 한 번 어긋난 줄, 항목 하나가 빠진 사업을 찾아냅니다."),
            ("목적 2 · 설명서 · UBIS · 입력본 세 곳의 금액이 같은가",
             "사업설명서 요구액 = UBIS 세출요구 검토조서 = K-에듀파인 입력본. "
             "셋이 어긋나면 어느 파일이 얼마인지 짚어 드립니다.",
             "입력 전(만든 입력본)과 입력 후(다시 내려받은 파일) 두 시점에 각각 검사합니다."),
            ("목적 3 · 올해 입력본을 대신 만듭니다",
             "작년 확정 파일의 모양을 그대로 따라 올해 입력본을 만듭니다. 레벨·순번·산출기초 "
             "표기를 K-에듀파인 규칙대로 채우고, 과거에 확정한 원가통계비목을 항목명과 상위 "
             "항목으로 찾아 자동으로 붙입니다.",
             "표본에서 산출근거 1,118행 중 910행(81%)이 자동으로 정해졌습니다. "
             "못 정한 줄은 빈칸으로 두고 따로 표시합니다."),
        ],
        "이 프로그램은 값을 고치지 않습니다. 어디가 어긋났는지만 알려 드립니다. "
        "원본 파일도 덮어쓰지 않습니다.",
    ),
    (
        "넣어야 할 자료 4가지",
        "다음 화면의 카드 4장이 이 자료들입니다. 카드마다 어디서 받는지 경로가 적혀 있습니다.",
        [
            ("1 · 사업설명서 HWPX     — 필수",
             "산출근거의 원문입니다. 사업명·요구내용·산출식·금액·단위를 여기서 읽습니다. "
             "목적 1이 이 파일 하나로 돌아갑니다.",
             "유비스 → 예산요구 → 사업별설명서 → 부서별(사업별 설명서)"),
            ("2 · 지난 K-에듀파인 입력본·다운로드 파일     — 필수",
             "올해 입력본의 모양을 여기서 배웁니다. 과거에 확정한 원가통계비목도 여기서 "
             "가져옵니다. 여러 부를 한 번에 고를 수 있습니다.",
             "부수가 늘수록 비목이 더 많이 자동으로 채워집니다. 1부 49% → 13부 81%"),
            ("3 · UBIS 세출요구 검토조서     — 선택",
             "설명서의 사업별 합계액과 UBIS 요구액이 같은지 대조합니다. "
             "넣지 않으면 이 금액 대조만 건너뜁니다.",
             "유비스 → 예산요구 → 세출요구출력"),
            ("4 · 세출예산 사업별 분류표 (별표 3)     — 선택",
             "정책·단위·세부사업 분류와 코드를 대조해 입력 행의 위계를 검토합니다.",
             "유비스 → 예산요구 → 세출요구"),
            ("5 · 단위과제카드 목록     — 선택",
             "내가 입력할 사업만 남깁니다. K-에듀파인은 자기 카드에 딸린 사업만 입력할 수 "
             "있는데, 과 전체 설명서에는 다른 담당자 사업까지 들어 있습니다.",
             "표본에서 69개 사업 중 5개만 남았습니다.\nK-에듀파인 → 재정사업관리 → 사업담당 → 재정관리 → 예산관리 → 사업관리카드 → 세출 조회 → 파일다운로드"),
        ],
        "3·4·5번을 빼도 진행됩니다. 5번을 넣으면 과 전체가 아니라 내 사업만 검사하고 "
        "입력본을 만듭니다.",
    ),
    (
        "무엇이 나오는가",
        "엑셀 파일 하나가 만들어집니다. 시트가 세 장이고, 쓰임이 다릅니다.",
        [
            ("입력본",
             "K-에듀파인에 그대로 옮겨 칠 표입니다. 레벨·순번·산출기초 표기를 "
             "K-에듀파인 규칙대로 채워 두었습니다.",
             "다운로드 서식 그대로라 열을 더하거나 빼지 않았습니다."),
            ("재배정(입력 안 함)",
             "총액배분사업비와 재원배분 사업입니다. 예산과가 학교로 바로 재배정하거나 "
             "다른 과가 나눠 주므로 이 과에서 입력하지 않습니다.",
             "입력본 시트에 섞어 두면 그대로 입력하게 되고, 아예 빼면 '설명서에 있던 "
             "사업이 왜 없지' 하게 됩니다. 그래서 빼되 사유와 함께 여기 남깁니다."),
            ("대본",
             "한 줄씩 짚어 가며 입력하라고 만든 표입니다. 확인 칸이 있고, 비목이 빈 줄은 "
             "빨갛게 칠해집니다.",
             "맨 위에 전체 행 수와 비목 미확정 행 수가 적힙니다."),
        ],
        "비목이 빈칸인 채로 나올 수 있습니다. 그건 막지 않습니다. 줄은 그대로 있고 칸만 "
        "비어 있어 입력하다 반드시 마주치니, K-에듀파인에서 직접 고르시면 됩니다.",
    ),
    (
        "이렇게 진행됩니다",
        "위쪽 탭 다섯 개가 순서입니다. 앞 단계를 마치지 않아도 탭을 눌러 오갈 수 있습니다.",
        [
            ("1 파일 불러오기  →  2 검사 결과",
             "자료를 고르고 F5를 누르면 읽기가 시작됩니다. 끝나면 지적 목록이 뜹니다. "
             "줄을 두 번 누르면 그 사업의 설명서 원문 행을 전부 보여 줍니다.",
             "지적에는 설명서에 적힌 사업 번호가 함께 나옵니다. 예: 15. 영유아교육내실화 지원"),
            ("지적은 세 등급입니다",
             "빨강 오류 · 주황 확인 필요 · 회색 안내. 오류가 남아 있으면 입력본을 만들지 "
             "않습니다. 금액이나 분류가 틀린 파일은 결과만 봐서는 알 수 없고, 그대로 "
             "입력되면 예산이 틀어지기 때문입니다.",
             "'확인 필요'는 설명서 안에서 기준이 갈리거나 이 프로그램이 단계를 다르게 "
             "읽었을 수 있는 줄입니다. 틀렸다는 뜻이 아닙니다."),
            ("3 비목 확정  →  4 입력본 생성  →  5 입력 후 검증",
             "자동으로 정하지 못한 비목을 목코드 단위로 묶어 한 번에 확정합니다. 한 번 정한 "
             "값은 다음 실행에서 자동으로 다시 적용됩니다.",
             "K-에듀파인 입력을 마치고 다시 내려받은 파일을 넣으면 마지막 대조를 합니다."),
        ],
        "F1 사용법 · F5 불러오기 · Ctrl+F 찾기 · F3 다음 지적 · Ctrl+S 입력본 생성 · Ctrl+1~5 화면 이동",
    ),
]


class Tutorial(tk.Toplevel):
    """첫 실행 안내 창. 언제든 '사용법' 버튼으로 다시 연다."""

    def __init__(self, parent, on_close=None) -> None:
        super().__init__(parent)
        self.parent, self.on_close = parent, on_close
        self.index = 0
        self.again = tk.BooleanVar(value=False)

        self.title("사용법")
        self.configure(bg=PANEL)
        self.resizable(True, True)
        self.minsize(560, 420)
        self.transient(parent)

        # 버튼 줄을 **먼저** 깐다. tkinter 의 pack 은 순서대로 자리를 나눠 주므로,
        # 본문을 먼저 깔면 내용이 긴 쪽에서 아래 버튼 줄이 0픽셀로 밀려 사라진다.
        # 실제로 2쪽에서 '다음'과 '다음부터 열지 않기'가 통째로 보이지 않았다.
        self._foot()

        area = tk.Frame(self, bg=PANEL)
        area.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(area, bg=PANEL, highlightthickness=0, bd=0)
        self.scroll = tk.Scrollbar(area, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.body = tk.Frame(self.canvas, bg=PANEL)
        self.window = self.canvas.create_window((0, 0), window=self.body, anchor="nw")
        self.canvas.bind("<Configure>",
                         lambda event: self.canvas.itemconfigure(self.window, width=event.width))
        self.body.bind("<Configure>", lambda _event: self._fit())
        for target in (self, self.canvas):
            target.bind("<MouseWheel>", self._wheel)
        self._draw()

        self.update_idletasks()
        # 업무망 노트북은 세로 768px 인 경우가 많다. 화면 밖으로 나가면 아래 버튼 줄을
        # 다시 못 보게 되므로 화면 안에 들어오도록 줄인다. 내용은 스크롤된다.
        width = min(820, self.winfo_screenwidth() - 60)
        height = min(680, self.winfo_screenheight() - 120)
        x = max(parent.winfo_rootx() + (parent.winfo_width() - width) // 2, 10)
        y = max(parent.winfo_rooty() + (parent.winfo_height() - height) // 3, 10)
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.after(0, self._fit)
        self.protocol("WM_DELETE_WINDOW", self._finish)
        self.bind("<Escape>", lambda _event: self._finish())
        self.bind("<Return>", lambda _event: self._next())
        self.grab_set()
        self.focus_set()

    def _fit(self) -> None:
        """내용이 창보다 길 때만 스크롤 막대를 보인다."""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        need = self.body.winfo_reqheight() > self.canvas.winfo_height()
        if need and not self.scroll.winfo_ismapped():
            self.scroll.pack(side="right", fill="y")
        elif not need and self.scroll.winfo_ismapped():
            self.scroll.pack_forget()

    def _wheel(self, event) -> None:
        self.canvas.yview_scroll(-1 if getattr(event, "delta", 0) > 0 else 1, "units")

    # ------------------------------------------------------------------ 내용
    def _draw(self) -> None:
        for child in self.body.winfo_children():
            child.destroy()
        title, lead, blocks, footer = PAGES[self.index]

        head = tk.Frame(self.body, bg=PANEL)
        head.pack(fill="x", padx=34, pady=(28, 0))
        tk.Label(head, text=f"{self.index + 1} / {len(PAGES)}", font=(FONT, 9, "bold"),
                 fg=ACCENT, bg=PANEL).pack(anchor="w")
        tk.Label(head, text=title, font=(FONT, 18, "bold"), fg=INK, bg=PANEL,
                 anchor="w", justify="left", wraplength=700).pack(fill="x", pady=(6, 0))
        tk.Label(head, text=lead, font=(FONT, 10), fg=MUTED, bg=PANEL,
                 anchor="w", justify="left", wraplength=700).pack(fill="x", pady=(8, 0))
        tk.Frame(self.body, bg=LINE, height=1).pack(fill="x", padx=34, pady=(18, 0))

        holder = tk.Frame(self.body, bg=PANEL)
        holder.pack(fill="both", expand=True, padx=34, pady=(16, 0))
        for heading, text, aside in blocks:
            row = tk.Frame(holder, bg=PANEL)
            row.pack(fill="x", pady=(0, 15))
            tk.Frame(row, bg=ACCENT, width=3).pack(side="left", fill="y")
            inner = tk.Frame(row, bg=PANEL)
            inner.pack(side="left", fill="x", expand=True, padx=(11, 0))
            tk.Label(inner, text=heading, font=(FONT, 11, "bold"), fg=INK, bg=PANEL,
                     anchor="w", justify="left", wraplength=680).pack(fill="x")
            tk.Label(inner, text=text, font=(FONT, 10), fg=BODY, bg=PANEL,
                     anchor="w", justify="left", wraplength=680).pack(fill="x", pady=(4, 0))
            tk.Label(inner, text=aside, font=(FONT, 9), fg=MUTED, bg=PANEL,
                     anchor="w", justify="left", wraplength=680).pack(fill="x", pady=(4, 0))

        note = tk.Frame(self.body, bg=OK_BG)
        note.pack(fill="x", padx=34, pady=(4, 20))
        tk.Label(note, text=footer, font=(FONT, 9), fg="#1E5C3B", bg=OK_BG,
                 anchor="w", justify="left", wraplength=690, padx=12, pady=9).pack(fill="x")

        self.canvas.yview_moveto(0)
        def attach(widget):
            widget.bind("<MouseWheel>", self._wheel)
            for child in widget.winfo_children():
                attach(child)

        attach(self.body)
        for position, dot in enumerate(self.dots):
            dot.configure(bg=ACCENT if position == self.index else LINE)
        self.back.configure(state="normal" if self.index else "disabled")
        self.forward.configure(text="시작하기" if self.index == len(PAGES) - 1 else "다음  →")

    def _foot(self) -> None:
        tk.Frame(self, bg=LINE, height=1).pack(fill="x", side="bottom")
        bar = tk.Frame(self, bg=GROUND)
        bar.pack(fill="x", side="bottom")
        inner = tk.Frame(bar, bg=GROUND)
        inner.pack(fill="x", padx=34, pady=14)

        tk.Checkbutton(inner, text="다음부터 열지 않기", variable=self.again,
                       font=(FONT, 9), fg=MUTED, bg=GROUND, activebackground=GROUND,
                       selectcolor=PANEL, bd=0, highlightthickness=0).pack(side="left")

        self.forward = tk.Button(inner, text="다음  →", command=self._next, font=(FONT, 10, "bold"),
                                 bg=ACCENT, fg="#FFFFFF", activebackground="#163A5A",
                                 activeforeground="#FFFFFF", relief="flat", padx=22, pady=7,
                                 cursor="hand2")
        self.forward.pack(side="right")
        self.back = tk.Button(inner, text="←  이전", command=self._previous, font=(FONT, 9),
                              bg="#F4F5F7", fg=BODY, relief="solid", bd=1, padx=14, pady=5,
                              cursor="hand2")
        self.back.pack(side="right", padx=8)

        dots = tk.Frame(inner, bg=GROUND)
        dots.pack(side="right", padx=16)
        self.dots = []
        for _ in PAGES:
            dot = tk.Frame(dots, bg=LINE, width=8, height=8)
            dot.pack(side="left", padx=3)
            self.dots.append(dot)

    # ------------------------------------------------------------------ 이동
    def _next(self) -> None:
        if self.index < len(PAGES) - 1:
            self.index += 1
            self._draw()
        else:
            self._finish()

    def _previous(self) -> None:
        if self.index:
            self.index -= 1
            self._draw()

    def _finish(self) -> None:
        if self.on_close:
            self.on_close(self.again.get())
        self.grab_release()
        self.destroy()
