"""Native Tk widgets for the approved UI. No browser, service, or network dependency."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk, font as tkfont
from pathlib import Path

from ui_common import (ACCENT, ACCENT_DIM, BODY, ERROR, FAINT, FONT, GROUND, GUIDE_BG, GUIDE_BODY, GUIDE_SOFT,
                       GUIDE_TAG, GUIDE_TEXT, HEAD, INK, LINE, MUTED, OK, OK_BG, PANEL, WARN, year_of)

FLOW_IDLE = "#8FA7B8"       # 아직 가지 않은 단계 사이 화살표. 연하면 흐름이 안 보인다


def rounded(canvas, x1, y1, x2, y2, radius=14, **kw):
    r = min(radius, max(0, (x2-x1)/2), max(0, (y2-y1)/2))
    return canvas.create_polygon(
        x1+r, y1, x2-r, y1, x2, y1, x2, y1+r,
        x2, y2-r, x2, y2, x2-r, y2, x1+r, y2,
        x1, y2, x1, y2-r, x1, y1+r, x1, y1,
        smooth=True, splinesteps=24, **kw)


class RoundedButton(tk.Canvas):
    """Keyboard-operable rounded button; retains the configure/invoke contract."""
    def __init__(self, parent, text, command=None, primary=False, width=None, hover=None, outline=None):
        self._text, self._command = text, command
        # 어두운 안내 말풍선 위에서는 기본 색(밝은 바탕에 맞춘 것)을 쓸 수 없다. 올렸을 때 색과
        # 테두리를 따로 받는다. 없으면 예전 그대로다.
        self._hover_fill, self._outline = hover, outline
        self._state, self._hover = "normal", False
        self._fill = ACCENT if primary else PANEL
        self._fg = PANEL if primary else BODY
        self._primary = primary
        self._font = (FONT, 10, "bold" if primary else "normal")
        self._fixed_width = width
        self._parent_bg = parent.cget("bg")
        super().__init__(parent, height=44, width=100, bg=self._parent_bg,
                         highlightthickness=0, bd=0, takefocus=True, cursor="hand2")
        self._measure()
        self.bind("<Configure>", self._draw)
        self.bind("<Enter>", lambda e: self._enter(True))
        self.bind("<Leave>", lambda e: self._enter(False))
        self.bind("<FocusIn>", self._draw)
        self.bind("<FocusOut>", self._draw)
        self.bind("<Button-1>", self._click)
        self.bind("<Return>", lambda e: self.invoke())
        self.bind("<space>", lambda e: self.invoke())

    def _measure(self):
        f = tkfont.Font(self, font=self._font)
        width = f.measure(self._text) + 36
        if self._fixed_width:
            width = max(width, f.measure("가") * self._fixed_width + 24)
        super().configure(width=max(60, width), height=max(44, f.metrics("linespace") + 20))

    def _enter(self, hover):
        self._hover = hover
        self._draw()

    def _click(self, _event):
        if self._state != "disabled":
            self.focus_set()
            self.invoke()

    def invoke(self):
        if self._state != "disabled" and self._command:
            return self._command()

    def configure(self, cnf=None, **kw):
        if isinstance(cnf, dict):
            kw = {**cnf, **kw}
        elif cnf is not None:
            return super().configure(cnf, **kw)
        if not kw:
            return super().configure()
        changed = False
        for key, attr in (("text", "_text"), ("command", "_command"), ("state", "_state"),
                          ("bg", "_fill"), ("background", "_fill"), ("fg", "_fg"),
                          ("foreground", "_fg"), ("font", "_font")):
            if key in kw:
                setattr(self, attr, kw.pop(key))
                changed = True
        for key in ("relief", "padx", "pady", "activebackground", "activeforeground"):
            kw.pop(key, None)
        if kw:
            super().configure(**kw)
        if changed:
            super().configure(takefocus=self._state != "disabled",
                              cursor="arrow" if self._state == "disabled" else "hand2")
            self._measure()
            self._draw()

    config = configure

    def set_primary(self, primary: bool) -> None:
        """강조 여부를 바꾼다. 다음에 누를 버튼 하나만 채운 색으로 보이게 할 때 쓴다."""
        if primary == self._primary:
            return
        self._primary = primary
        self._fill = ACCENT if primary else PANEL
        self._fg = PANEL if primary else BODY
        self._font = (FONT, 10, "bold" if primary else "normal")
        self._measure()
        self._draw()

    def cget(self, key):
        if key in ("text", "state", "command"):
            return getattr(self, "_" + key)
        if key in ("fg", "foreground"):
            return self._fg
        return super().cget(key)

    def __getitem__(self, key):
        return self.cget(key)

    def _draw(self, _event=None):
        self.delete("all")
        w, h = self.winfo_width(), self.winfo_height()
        if w < 3:
            return
        disabled = self._state == "disabled"
        fill = HEAD if disabled else (self._hover_fill or (ACCENT_DIM if self._primary else HEAD)) \
            if self._hover else self._fill
        foreground = MUTED if disabled else self._fg
        focused = self.focus_get() == self
        rounded(self, 2, 2, w-2, h-2, 13, fill=fill,
                outline=ACCENT if focused else self._outline or (self._fill if self._primary else LINE),
                width=2 if focused else 1)
        self.create_text(w/2, h/2, text=self._text, fill=foreground, font=self._font)


class RoundedPanel(tk.Canvas):
    """Auto-height rounded container with a normal Tk content frame."""
    def __init__(self, parent, fill=PANEL, border=LINE, padding=16):
        super().__init__(parent, bg=parent.cget("bg"), highlightthickness=0, bd=0, height=80)
        self.fill, self.border, self.padding = fill, border, padding
        self.inner = tk.Frame(self, bg=fill)
        self._shape = rounded(self, 1, 1, 30, 30, 18, fill=fill, outline=border)
        self._window = self.create_window(padding, padding, anchor="nw", window=self.inner)
        self.bind("<Configure>", self._layout)
        self.inner.bind("<Configure>", self._layout)

    def set_border(self, color):
        self.border = color
        self.itemconfigure(self._shape, outline=color, width=2 if color == ACCENT else 1)

    def _layout(self, _event=None):
        width = max(80, self.winfo_width())
        self.itemconfigure(self._window, width=max(40, width-2*self.padding))
        height = self.inner.winfo_reqheight() + 2*self.padding
        if abs(height - int(float(self.cget("height")))) > 1:
            super().configure(height=height)
        self.delete(self._shape)
        self._shape = rounded(self, 1, 1, width-1, height-1, 18, fill=self.fill,
                              outline=self.border, width=2 if self.border == ACCENT else 1)
        self.tag_lower(self._shape)


def text(parent, content, size=10, bold=False, color=BODY, **kw):
    return tk.Label(parent, text=content, bg=parent.cget("bg"), fg=color,
                    font=(FONT, size, "bold" if bold else "normal"), **kw)


class ProcessFlow(tk.Canvas):
    """위쪽 탭과 **같은** 다섯 단계. 예전에는 탭과 이 그림이 서로 다른 다섯 단계를
    말해서('비목 확정'이 한쪽에만 있었다) 지금 몇 번째 단계인지 헷갈렸다."""
    TITLES = ("자료 선택", "검사 결과 확인", "비목 확정", "입력본 만들기", "입력 후 대조")
    NOTES = ("1·2번 필수(추경은 1번만)", "계산 오류 · UBIS 금액 차이", "원가통계비목 7자리 채우기",
             "입력본 · 대본 · 재배정 시트", "K-에듀파인 다운로드본과 비교")
    HEIGHT = 128

    def __init__(self, parent):
        super().__init__(parent, height=self.HEIGHT, bg=GROUND, bd=0, highlightthickness=0)
        self.stage = 0
        self.done: set = set()
        self.bind("<Configure>", self.draw)

    def set_stage(self, stage, done=None):
        self.stage = stage
        if done is not None:
            self.done = set(done)
        self.draw()

    def draw(self, _event=None):
        self.delete("all")
        w = self.winfo_width()
        if w < 10:
            return
        bottom = self.HEIGHT - 2
        rounded(self, 1, 1, w-1, bottom, 20, fill=PANEL, outline=LINE)
        self.create_text(20, 20, anchor="w", text="작업 순서  ·  위쪽 탭과 같은 다섯 단계입니다",
                         font=(FONT, 9, "bold"), fill=INK)
        col = (w-40)/5
        for i, (title, note) in enumerate(zip(self.TITLES, self.NOTES)):
            x = 20 + i*col
            finished = i in self.done or i < self.stage
            current = i == self.stage
            rounded(self, x, 38, x+32, 70, 11,
                    fill=ACCENT if current else OK if finished else HEAD, outline="")
            self.create_text(x+16, 54, text="✓" if finished and not current else f"{i+1}",
                             fill=PANEL if (current or finished) else MUTED, font=(FONT, 10, "bold"))
            if i < 4:
                # 단계 사이 화살표. 예전 것은 가늘고(2px) 연한 물결선에 촉도 작아서 흐름이라기보다
                # 밑줄처럼 보였다. 곧게 긋고 3px·큰 촉으로 바꾸고, 색으로 상태를 말한다:
                # 끝낸 구간은 초록, 지금 단계에서 다음으로 가는 구간은 강조색, 나머지는 회청색.
                leg = OK if finished and not current else ACCENT if current else FLOW_IDLE
                self.create_line(x+44, 54, x+col-14, 54, arrow="last", arrowshape=(13, 15, 7),
                                 width=3, capstyle="round", fill=leg)
            self.create_text(x, 88, anchor="w", text=title, font=(FONT, 9, "bold"),
                             fill=ACCENT if current else INK)
            self.create_text(x, 108, anchor="w", text=note, font=(FONT, 8), fill=MUTED)


def year_warnings(paths: dict) -> dict:
    """파일 이름의 회계연도로 짝이 안 맞는 자료를 짚는다. {키: 문구}.

    다른 해 자료를 섞어 넣어도 프로그램은 돌아가고 결과도 그럴듯하다. 금액 대조가
    전부 어긋나야 비로소 알게 된다. 고르는 자리에서 바로 알려 준다. 이름에 연도가
    없으면 말하지 않는다(모르는 것을 틀렸다고 하지 않는다). 막지는 않는다.
    """
    years = {key: year_of(Path(value).name) for key, value in paths.items() if value}
    plan = years.get("plan")
    found: dict = {}
    if plan:
        ubis = years.get("ubis")
        if ubis and ubis != plan:
            found["ubis"] = f"연도 확인: 설명서는 {plan}년, 이 검토조서는 {ubis}년 자료입니다."
        elif ubis:
            from supplement import round_of

            plan_round, ubis_round = round_of(Path(paths["plan"]).name), round_of(Path(paths["ubis"]).name)
            if plan_round and ubis_round and plan_round != ubis_round:
                word = {0: "본예산", 1: "1회 추경", 2: "2회 추경"}
                found["ubis"] = (f"차수 확인: 설명서는 {word.get(plan_round[1], plan_round[1])}, 이 검토조서는 "
                                 f"{word.get(ubis_round[1], ubis_round[1])} 자료로 보입니다.")
        last = years.get("last_year")
        if last and last >= plan:
            found["last_year"] = (f"연도 확인: 설명서가 {plan}년인데 이 파일은 {last}년입니다. "
                                  f"보통은 앞선 해({plan - 1}년) 확정본을 넣습니다.")
        cards = years.get("cards")
        if cards and cards != plan:
            found["cards"] = f"연도 확인: 설명서는 {plan}년, 이 카드 목록은 {cards}년 자료입니다."
    return found


class FilePage(tk.Frame):
    """Five material rows, real source paths and an inline, anchored tutorial."""
    def __init__(self, app, parent):
        super().__init__(parent, bg=GROUND)
        self.app = app
        self.active = -1
        self.return_focus = None
        self.rows, self.coaches, self.routes, self.states, self.numbers = {}, {}, {}, {}, {}
        self.places, self.checks, self.pick_buttons, self.drop_buttons = {}, {}, {}, {}
        self.route_buttons, self.file_buttons, self.wraps = {}, [], []
        self.pack(fill="both", expand=True)
        bottom = tk.Frame(self, bg=PANEL, highlightbackground=LINE, highlightthickness=1)
        bottom.pack(side="bottom", fill="x")
        self.run = RoundedButton(bottom, "자료 불러오고 검사하기  (F5)  →", app.load, True)
        self.run.pack(side="right", padx=24, pady=12)
        self.reset = RoundedButton(bottom, "모두 빼기", app.clear_all)
        self.reset.pack(side="right", pady=12)
        self.ready = text(bottom, "필수 자료 1·2번을 선택해 주세요.", 10, color=MUTED, anchor="w", justify="left")
        self.ready.pack(side="left", padx=24, fill="x", expand=True)

        progress_holder = tk.Frame(self, bg=GROUND)
        progress_holder.pack(side="bottom", fill="x", padx=24)
        app.progress_box = tk.Frame(progress_holder, bg=GROUND)
        app.progress_text = text(app.progress_box, "", 9)
        app.progress_text.pack(fill="x", pady=(8, 2))
        app.progress = ttk.Progressbar(app.progress_box, style="Bar.Horizontal.TProgressbar", maximum=100)
        app.progress.pack(fill="x", pady=(0, 8))

        area = tk.Frame(self, bg=GROUND)
        area.pack(fill="both", expand=True)
        self.canvas = tk.Canvas(area, bg=GROUND, bd=0, highlightthickness=0)
        bar = ttk.Scrollbar(area, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=bar.set)
        bar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.sheet = tk.Frame(self.canvas, bg=GROUND)
        window = self.canvas.create_window(0, 0, window=self.sheet, anchor="nw")
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(window, width=e.width))
        self.sheet.bind("<Configure>", self._resized)
        pad = tk.Frame(self.sheet, bg=GROUND)
        pad.pack(fill="x", padx=24, pady=(18, 18))
        # 첫 화면의 주인공은 '넣어야 할 자료'다. 예전에는 큰 제목과 옆 안내판이 윗부분을
        # 차지해 1440x900 에서도 자료 칸이 두 개만 보였다. 설명은 한 줄로 줄였다.
        self.desc = text(pad, "사업설명서의 금액을 검산하고, K-에듀파인 입력본(엑셀)을 만든 뒤, 입력 결과를 다시 "
                              "대조합니다. 원본 파일은 고치지 않고 틀린 곳만 알려 드립니다.",
                         10, color=MUTED, justify="left", anchor="w")
        self.desc.pack(fill="x", pady=(0, 12))
        self._wrap(pad, self.desc)
        self.flow = ProcessFlow(pad)
        self.flow.pack(fill="x", pady=(0, 16))
        self.flow_host = tk.Frame(pad, bg=GROUND)
        self.flow_host.pack(fill="x")

        left = tk.Frame(pad, bg=GROUND)
        self.left = left
        left.pack(fill="both", expand=True)
        heading = tk.Frame(left, bg=GROUND)
        heading.pack(fill="x", pady=(0, 4))
        text(heading, "넣어야 할 자료", 15, True, INK).pack(side="left")
        self.counter = text(heading, "", 10, True, ACCENT)
        self.counter.pack(side="right")
        text(left, "1·2번은 꼭 필요합니다. 3·4·5번은 넣으면 검사 범위가 넓어집니다.", 9, color=MUTED,
             anchor="w").pack(fill="x", pady=(0, 12))
        self.list = tk.Frame(left, bg=GROUND)
        self.list.pack(fill="x")
        for index, entry in enumerate(app.FILE_CARDS):
            key, number, name, need, short, what, where, sample, types = entry
            row = RoundedPanel(self.list, padding=14)
            row.pack(fill="x", pady=(0, 10))
            self.rows[key] = row
            inner = row.inner
            number_canvas = tk.Canvas(inner, bg=PANEL, width=38, height=40, bd=0, highlightthickness=0)
            number_canvas.pack(side="left", fill="y", padx=(0, 12))
            number_canvas.bind("<Configure>", lambda e, k=key: self._number(k, bool(app.paths[k].get().strip())))
            self.numbers[key] = number_canvas
            actions = tk.Frame(inner, bg=PANEL)
            actions.pack(side="right", anchor="n", padx=(12, 0))
            help_button = RoundedButton(actions, "안내", lambda n=index+4: self.open_guide(n))
            help_button.pack(side="left", padx=(0, 4))
            # 파일 선택 버튼은 매 행 같은 자리·같은 너비다(DESIGN.md). 글자가 바뀌어도 밀리지 않는다.
            pick = RoundedButton(actions, "파일 선택", lambda k=key, t=types: app.pick(k, t), width=6)
            pick.pack(side="left")
            drop = RoundedButton(actions, "빼기", lambda k=key: app.clear(k))
            drop.pack(side="left", padx=(4, 0))
            self.pick_buttons[key], self.drop_buttons[key] = pick, drop
            self.file_buttons += [help_button]
            content = tk.Frame(inner, bg=PANEL)
            content.pack(side="left", fill="both", expand=True)
            title_line = tk.Frame(content, bg=PANEL)
            title_line.pack(fill="x")
            title = text(title_line, f"{number}. {name}", 10, True, INK, anchor="w", justify="left")
            title.pack(side="left")
            badge = tk.Label(title_line, text=f" {need} ", font=(FONT, 8, "bold"),
                             fg=PANEL if need == "필수" else MUTED,
                             bg=ACCENT if need == "필수" else HEAD, padx=4, pady=0)
            badge.pack(side="left", padx=(8, 0))
            detail = text(content, short, 9, color=MUTED, anchor="w", justify="left")
            detail.pack(fill="x", pady=(3, 2))
            route_button = tk.Button(content, text="▸ 자료 받는 경로", command=lambda k=key: self.toggle_route(k),
                                     bg=PANEL, fg=ACCENT, activebackground=PANEL, activeforeground=ACCENT_DIM,
                                     relief="flat", bd=0, highlightthickness=0, font=(FONT, 9),
                                     cursor="hand2", anchor="w", takefocus=True)
            route_button.pack(anchor="w")
            self.route_buttons[key] = route_button
            route = text(content, where, 9, color=BODY, anchor="w", justify="left")
            self.routes[key] = route
            state = text(content, "", 9, color=MUTED, anchor="w", justify="left")
            state.pack(fill="x", pady=(3, 0))
            self.states[key] = state
            place = text(content, "", 8, color=FAINT, anchor="w", justify="left")
            self.places[key] = place
            check = text(content, "", 9, color=WARN, anchor="w", justify="left")
            self.checks[key] = check
            if key == "last_year":
                app.year_note = text(content, "여러 파일을 함께 선택할 수 있습니다.", 8, color=MUTED,
                                     anchor="w", justify="left")
                app.year_note.pack(fill="x", pady=(3, 0))
                self._wrap(content, app.year_note)
            for widget in (detail, route, state, place, check):
                self._wrap(content, widget)
            self.coaches[key] = tk.Frame(self.list, bg=GROUND)
            self.coaches[key].pack(fill="x")
        text(pad, "고른 자료는 다음 실행 때도 기억합니다. 다른 해 작업을 할 때는 자료를 바꿔 주세요.",
             9, color=MUTED).pack(anchor="w", pady=(8, 0))
        self.hide_again = tk.BooleanVar(value=bool(app.settings.get("tutorial_seen")))
        self.coach = None
        self.bind_wheel(area)
        self.mark_files()

    def _wrap(self, parent, widget):
        def fit(e=None):
            widget.configure(wraplength=max(120, parent.winfo_width()-6))
        parent.bind("<Configure>", fit, add="+")
        widget.configure(wraplength=380)

    def _resized(self, _event=None):
        region = self.canvas.bbox("all")
        self.canvas.configure(scrollregion=region)
        if not region:
            return
        # 내용이 줄었는데 스크롤이 그대로면 빈 곳을 보게 된다. 안내를 닫았을 때가 그렇다.
        span = max(1, region[3]-region[1])
        top = self.canvas.canvasy(0)
        overflow = top+self.canvas.winfo_height()-region[3]
        if overflow > 0:
            self.canvas.yview_moveto(max(0.0, top-overflow)/span)

    def bind_wheel(self, root):
        def wheel(e):
            if self.app.page != "files":
                return
            self.canvas.yview_scroll(-1 if e.delta > 0 else 1, "units")
            return "break"
        root.bind("<MouseWheel>", wheel)
        for child in root.winfo_children():
            self.bind_wheel(child)

    def toggle_route(self, key):
        route = self.routes[key]
        if route.winfo_manager():
            route.pack_forget()
            self.route_buttons[key].configure(text="▸ 자료 받는 경로")
        else:
            route.pack(fill="x", before=self.states[key], pady=(5, 6))
            self.route_buttons[key].configure(text="▾ 자료 받는 경로")

    def _show(self, widget, content: str, before=None) -> None:
        """글이 있으면 보이고 없으면 자리째 숨긴다. 빈 줄이 칸마다 남지 않게."""
        widget.configure(text=content)
        if content and not widget.winfo_manager():
            widget.pack(fill="x", pady=(2, 0), **({"before": before} if before is not None else {}))
        elif not content and widget.winfo_manager():
            widget.pack_forget()

    def mark_files(self):
        busy = getattr(self.app, "loading", False)
        paths = {key: self.app.paths[key].get().strip() for key in self.rows}
        warnings = year_warnings(paths)
        # 추경 설명서는 입력본을 만들지 않으므로 2번(K-에듀파인 파일)이 필요 없다.
        supplement = bool(getattr(self.app, "plan_is_supplement", lambda: False)())
        needed = ("plan",) if supplement else ("plan", "last_year")
        for key, row in self.rows.items():
            name = paths[key]
            # 지난번에 고른 파일을 기억해 두므로, 그새 옮기거나 지운 파일이 '✓ 선택 완료'로
            # 남아 있을 수 있다. 불러오기를 눌러서야 영어 오류로 알게 하지 않는다.
            gone = bool(name) and not Path(name).exists()
            if gone:
                warnings[key] = "파일을 찾을 수 없습니다. 옮기거나 지웠을 수 있습니다 — 다시 선택해 주세요."
            self.states[key].configure(text=("✕ " if gone else "✓ ") + Path(name).name if name
                                       else "아직 선택하지 않았습니다.",
                                       fg=ERROR if gone else OK if name else MUTED,
                                       font=(FONT, 9, "bold" if name else "normal"))
            after = self.app.year_note if key == "last_year" else None
            self._show(self.checks[key], warnings.get(key, ""), before=after)
            self._show(self.places[key], f"위치: {Path(name).parent}" if name else "",
                       before=self.checks[key] if self.checks[key].winfo_manager() else after)
            self._number(key, name != "")
            active = self._guide_key(self.active) == key
            row.set_border(ACCENT if active else WARN if key in warnings else OK if name else LINE)
            pick, drop = self.pick_buttons[key], self.drop_buttons[key]
            # 다음에 누를 버튼 하나만 채운 색으로. 필수 자료가 비어 있으면 그 칸의 '파일 선택'이다.
            pick.set_primary(not name and key in needed)
            pick.configure(text="다른 파일" if name else "파일 선택", state="disabled" if busy else "normal")
            drop.configure(state="disabled" if busy or not name else "normal")
        required = sum(1 for k in needed if paths[k])
        optional = sum(1 for k in ("ubis", "classes", "cards") if paths.get(k))
        self.counter.configure(text=(f"추경 점검 · 필수 {required}/1 · 선택 {optional}/3" if supplement
                                     else f"필수 {required}/2 · 선택 {optional}/3"))
        ready = required == len(needed)
        self.run.configure(state="normal" if ready and not busy else "disabled")
        if busy:
            message, color = "자료를 읽고 있습니다… 잠시 기다려 주세요.", MUTED
        elif not ready:
            missing = [label for k, label in (("plan", "1번 사업설명서"), ("last_year", "2번 지난 K-에듀파인 파일"))
                       if not paths[k]]
            message, color = f"{' · '.join(missing)}을(를) 선택해 주세요.", MUTED
        elif warnings:
            message, color = "준비됐습니다. 다만 연도가 맞지 않는 자료가 있습니다 — 노란 칸을 확인해 주세요.", WARN
        elif supplement:
            message, color = ("추경 설명서입니다. 증감·산출식·합계와 UBIS 추경 검토조서를 점검합니다 "
                              "(2번 K-에듀파인 파일은 쓰지 않습니다). F5 로 시작합니다."), OK
        else:
            message, color = "준비됐습니다. 오른쪽 버튼이나 F5 로 검사를 시작합니다.", OK
        self.ready.configure(text=message, fg=color)
        for b in self.file_buttons + list(self.route_buttons.values()) + [self.reset]:
            b.configure(state="disabled" if busy else "normal")
        if not busy and not any(paths.values()):
            self.reset.configure(state="disabled")

    def _number(self, key, selected):
        c = self.numbers[key]
        c.delete("all")
        active = self._guide_key(self.active) == key
        rounded(c, 1, 1, 37, 39, 12, fill=ACCENT if active else OK_BG if selected else HEAD, outline="")
        number = next(x[1] for x in self.app.FILE_CARDS if x[0] == key)
        c.create_text(19, 20, text=number, font=(FONT, 10, "bold"), fill=PANEL if active else OK if selected else MUTED)

    @staticmethod
    def _guide_button(parent, label, command):
        """어두운 말풍선 위의 보조 버튼. 바탕과 같은 색에 밝은 테두리·글자."""
        widget = RoundedButton(parent, label, command, hover=GUIDE_SOFT, outline="#5A7288")
        widget.configure(bg=GUIDE_BG, fg=GUIDE_TEXT)
        return widget

    def _guide_key(self, index):
        if index >= 4:
            return self.app.FILE_CARDS[index-4][0]
        return {1: "plan", 2: "ubis", 3: "last_year"}.get(index)

    def open_guide(self, index=0):
        if getattr(self.app, "loading", False):
            return
        if self.active < 0:
            self.return_focus = self.app.focus_get()
        self.app.show("files")
        self._drop_coach()
        self.active = index
        key = self._guide_key(index)
        host = self.coaches[key] if key else self.flow_host
        outer = tk.Frame(host, bg=GROUND)
        outer.pack(fill="x", pady=(0, 16))
        self.coach = outer
        # 안내는 작업 카드와 **달라 보여야** 한다. 예전에는 안내도 흰 카드라 목록 사이에
        # 끼면 어느 것이 안내이고 어느 것이 넣을 자료인지 구분되지 않았다. 색을 뒤집은
        # 말풍선(어두운 바탕·밝은 글자)으로 바꾸고, 위쪽 꼭지로 설명하는 칸을 직접 가리킨다.
        notch = tk.Canvas(outer, height=14, bg=GROUND, bd=0, highlightthickness=0)
        notch.pack(fill="x")
        notch.create_polygon(30, 14, 44, 0, 58, 14, fill=GUIDE_BG, outline=GUIDE_BG)
        panel = RoundedPanel(outer, fill=GUIDE_BG, border=GUIDE_BG, padding=20)
        panel.pack(fill="x", expand=True)
        inner = panel.inner
        top = tk.Frame(inner, bg=GUIDE_BG)
        top.pack(fill="x")
        kind = "자료 안내" if index >= 4 else "전체 흐름" if index == 0 else "이 도구가 하는 일"
        text(top, f"?  사용법 안내  ·  {kind}", 9, True, GUIDE_TAG).pack(side="left")
        self._guide_button(top, "닫기  (Esc)", self.close_guide).pack(side="right")
        if index >= 4:
            entry = self.app.FILE_CARDS[index-4]
            title, body, route = f"{entry[1]}. {entry[2]}", entry[5], entry[6]
        else:
            title, body = (
                ("전체 흐름을 먼저 살펴보세요", "위쪽 탭과 같은 다섯 단계로 진행합니다. ① 자료를 고르고 ② 검사 결과를 확인한 뒤 "
                 "③ 비목을 확정하고 ④ 입력본을 만듭니다. K-에듀파인에 입력한 뒤에는 ⑤ 내려받은 파일을 다시 대조합니다."),
                ("설명서 검산 · 설명서 안의 금액", "사업설명서의 산출식을 계산하고, 항목 합계가 윗 단계·올해 요구액과 같은지 확인합니다."),
                ("UBIS 대조 · 두 자료의 금액", "1번 사업설명서  ↔  3번 UBIS 세출요구 검토조서\n사업별 금액이 같은지 확인합니다. 3번을 넣지 않으면 이 대조만 건너뜁니다."),
                ("입력본 만들기 · K-에듀파인 형식", "1번 사업설명서의 항목 → 올해 입력본\n2번 지난 K-에듀파인 파일에서 양식과 원가통계비목(7자리)을 배웁니다."),
            )[index]
            route = ""
        heading = text(inner, title, 13, True, GUIDE_TEXT, anchor="w", justify="left")
        heading.pack(fill="x", pady=(10, 8))
        desc = text(inner, body, 10, color=GUIDE_BODY, anchor="w", justify="left")
        desc.pack(fill="x")
        self._wrap(inner, heading)
        self._wrap(inner, desc)
        if route:
            source = tk.Frame(inner, bg=GUIDE_SOFT)
            source.pack(fill="x", pady=(14, 4))
            text(source, "자료 받는 경로", 9, True, GUIDE_TAG).pack(anchor="w", padx=12, pady=(10, 4))
            route_text = text(source, route, 9, color=GUIDE_TEXT, anchor="w", justify="left")
            route_text.pack(fill="x", padx=12, pady=(0, 12))
            self._wrap(source, route_text)
        foot = tk.Frame(inner, bg=GUIDE_BG)
        foot.pack(fill="x", pady=(16, 0))
        self._guide_button(foot, "건너뛰기", self.close_guide).pack(side="left")
        text(foot, f"{index+1} / 9", 9, color=GUIDE_BODY).pack(side="left", padx=12)
        next_button = RoundedButton(foot, "시작하기" if index == 8 else "다음 →",
                                    self.close_guide if index == 8 else lambda: self.open_guide(index+1),
                                    hover=HEAD, outline=PANEL)
        next_button.configure(bg=PANEL, fg=INK, font=(FONT, 10, "bold"))
        next_button.pack(side="right")
        if index > 0:
            self._guide_button(foot, "이전", lambda: self.open_guide(index-1)).pack(side="right", padx=(0, 6))
        tk.Checkbutton(inner, text="다음 실행부터 자동 안내 열지 않기", variable=self.hide_again,
                       bg=GUIDE_BG, fg=GUIDE_BODY, font=(FONT, 9), activebackground=GUIDE_BG,
                       activeforeground=GUIDE_TEXT, selectcolor=GUIDE_SOFT,
                       highlightthickness=0).pack(anchor="w", pady=(10, 0))
        self.bind_wheel(outer)
        self.mark_files()
        self.app.update_idletasks()
        self._resized()
        self.app.after_idle(lambda: self._reveal(key, index))
        next_button.focus_set()

    def _reveal(self, key, index):
        """안내가 화면 안에 들어오게 스크롤한다.

        행과 안내를 합친 높이가 보이는 영역보다 크면 **안내 바닥을 지킨다.**
        거기에 '다음'·'건너뛰기'·'닫기'가 있어서, 잘리면 다음 쪽으로 갈 길이 없다.
        행 위쪽은 맥락일 뿐이라 이때는 밀려 올라가도 된다. 둘 다 지킬 수 없을 때
        무엇을 포기할지가 요점이다. (1080x680 에서 7px 이 잘려 빌드 관문이 잡았다)
        """
        if not self.coach or self.active != index:
            return
        self.app.update_idletasks()
        top_widget = self.rows[key] if key else self.flow
        top = top_widget.winfo_rooty()-self.sheet.winfo_rooty()
        bottom = self.coach.winfo_rooty()-self.sheet.winfo_rooty()+self.coach.winfo_height()
        viewport = self.canvas.winfo_height()
        height = max(1, self.sheet.winfo_height())
        fits = bottom-top+28 <= viewport
        goal = max(0, top-12) if fits else max(0, bottom-viewport+16)
        self.canvas.yview_moveto(goal/height)
        # 계산이 어긋나도 눈으로는 맞아야 한다. 실제로 재 보고 모자란 만큼만 더 내린다.
        self.app.update_idletasks()
        if not self.coach:
            return
        over = ((self.coach.winfo_rooty()+self.coach.winfo_height())
                - (self.canvas.winfo_rooty()+self.canvas.winfo_height()))
        if over > 0:
            self.canvas.yview_moveto(min(1.0, (goal+over+8)/height))

    def _drop_coach(self):
        """안내를 지우고 **그 자리도 같이** 지운다.

        tkinter 의 Frame 은 마지막 자식이 사라져도 이전에 잡아 둔 크기를 그대로 들고
        있다. 자식이 0개인데 요청 높이가 328px 로 남아 행과 행 사이가 그만큼 벌어진
        채 굳는다. 담당자 눈에는 자료 목록이 사라진 것처럼 보인다. 높이를 직접
        1 로 되돌려야 원래 자리로 돌아온다.
        """
        if not self.coach:
            return
        host = self.coach.master
        self.coach.destroy()
        self.coach = None
        if host.winfo_exists() and not host.winfo_children():
            host.configure(height=1)

    def close_guide(self, persist=True):
        self._drop_coach()
        self.active = -1
        self.mark_files()
        if persist:
            self.app.settings["tutorial_seen"] = self.hide_again.get()
            import settings
            settings.save(self.app.settings)
        if self.return_focus and self.return_focus.winfo_exists():
            self.return_focus.focus_set()
        self.app.after_idle(self._restore)

    def _restore(self):
        """안내를 닫은 뒤 화면을 원래대로 되돌린다.

        안내는 마지막 쪽에서 화면을 한참 아래로 내려 둔다. 그 상태에서 안내만 지우면
        스크롤은 그대로인데 내용은 줄어, 담당자 눈에는 자료 목록이 사라진 것처럼 보인다.
        캔버스에 그린 테두리와 흐름선도 다시 그려야 자리가 맞는다.
        """
        if not self.winfo_exists():
            return
        self.app.update_idletasks()
        for row in self.rows.values():
            row._layout()
        self.flow.draw()
        self._resized()
        self.canvas.yview_moveto(0)
