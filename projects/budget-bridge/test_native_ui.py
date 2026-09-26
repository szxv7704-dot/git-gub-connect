"""Native integration tests isolate user settings and exercise real Tk geometry."""
import copy
import unittest
from unittest.mock import patch

import budget_bridge
import settings
from native_ui import RoundedButton


class NativeUITests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        data = copy.deepcopy(settings.DEFAULTS)
        data['tutorial_seen'] = True
        cls.patches = [patch.object(settings, 'load', return_value=data),
                       patch.object(settings, 'save', return_value=True)]
        for p in cls.patches:
            p.start()
        cls.app = budget_bridge.App()
        cls.errors = []
        cls.app.report_callback_exception = lambda kind, value, tb: cls.errors.append(str(value))
        cls.app.update()

    @classmethod
    def tearDownClass(cls):
        cls.app.destroy()
        for p in cls.patches:
            p.stop()

    def test_guides_follow_rows_at_desktop_sizes(self):
        app, page = self.app, self.app.file_view
        for geometry in ('1440x900', '1366x768', '1080x680'):
            app.geometry(geometry)
            app.update()
            for step in range(9):
                with self.subTest(geometry=geometry, step=step):
                    page.open_guide(step)
                    app.update()
                    app.update_idletasks()
                    key = page._guide_key(step)
                    self.assertEqual(page.coach.master, page.coaches[key] if key else page.flow_host)
                    panel = page.coach.winfo_children()[1]
                    self.assertGreater(panel.winfo_width(), 200)
                    self.assertLessEqual(panel.winfo_rootx()+panel.winfo_width(),
                                         page.canvas.winfo_rootx()+page.canvas.winfo_width()+2)
                    # Normal desktop height must show all guide controls above the pinned footer.
                    self.assertLessEqual(page.coach.winfo_rooty()+page.coach.winfo_height(),
                                         page.canvas.winfo_rooty()+page.canvas.winfo_height()+3)
                    if key:
                        self.assertLessEqual(page.rows[key].winfo_rooty()+page.rows[key].winfo_height(),
                                             page.coach.winfo_rooty()+2)
            page.close_guide(False)
        self.assertFalse(self.errors)

    def test_paths_are_complete_and_toggle_in_place(self):
        page = self.app.file_view
        for entry in self.app.FILE_CARDS:
            key = entry[0]
            self.assertEqual(page.routes[key].cget('text'), entry[6])
            page.toggle_route(key)
            self.app.update()
            self.assertEqual(page.routes[key].winfo_manager(), 'pack')
            page.toggle_route(key)
            self.assertFalse(page.routes[key].winfo_manager())
        self.assertIn('단위과제카드번호', page.routes['last_year'].cget('text'))
        self.assertEqual(page.routes['ubis'].cget('text'), '유비스 → 예산요구 → 세출요구출력')

    def test_required_files_and_busy_guards(self):
        app, page = self.app, self.app.file_view
        app.clear_all()
        self.assertEqual(page.run.cget('state'), 'disabled')
        app.paths['plan'].set('plan.hwpx')
        app.paths['last_year'].set('last.xlsx')
        app._mark_files()
        self.assertEqual(page.run.cget('state'), 'normal')
        app.loading = True
        app._mark_files()
        app.clear_all()
        app.load()
        app.show_tutorial()
        self.assertEqual(app.paths['plan'].get(), 'plan.hwpx')
        self.assertEqual(page.run.cget('state'), 'disabled')
        self.assertEqual(page.active, -1)
        app.loading = False
        app.clear_all()

    def test_closing_the_guide_brings_the_list_back(self):
        """안내를 다 보고 닫으면 자료 목록이 **처음 자리 그대로** 돌아와야 한다.

        안내가 놓였던 자리(tk.Frame)는 자식이 사라져도 이전 크기를 그대로 들고 있다.
        자식 0개에 요청 높이 328px 이 남아 행과 행 사이가 그만큼 벌어진 채 굳었고,
        담당자 화면에서는 자료 목록이 사라진 것처럼 보였다. 높이만 보지 말고
        **행 사이 간격과 전체 높이**를 재야 잡힌다.
        """
        app, page = self.app, self.app.file_view
        app.geometry('1440x900')
        app.update()
        app.update_idletasks()
        before_sheet = page.sheet.winfo_height()
        before_rows = {k: page.rows[k].winfo_y() for k in page.rows}
        for step in range(9):
            page.open_guide(step)
            app.update()
        page.close_guide(False)
        app.update()
        app.update_idletasks()
        self.assertIsNone(page.coach)
        self.assertLessEqual(page.canvas.yview()[0], 0.001, '안내를 닫으면 맨 위로 돌아온다')
        for host in list(page.coaches.values()) + [page.flow_host]:
            self.assertLessEqual(host.winfo_reqheight(), 2,
                                 '안내가 놓였던 자리가 높이를 그대로 들고 있다')
        self.assertEqual({k: page.rows[k].winfo_y() for k in page.rows}, before_rows,
                         '행 사이가 벌어진 채 굳었다')
        self.assertEqual(page.sheet.winfo_height(), before_sheet, '전체 높이가 돌아오지 않았다')
        for entry in app.FILE_CARDS:
            row = page.rows[entry[0]]
            self.assertTrue(row.winfo_manager(), f'{entry[0]} 행이 사라졌다')
            self.assertGreater(row.winfo_height(), 40, f'{entry[0]} 행이 찌그러졌다')
        self.assertFalse(self.errors)

    def test_detail_screen_shows_the_table_the_issue_is_actually_in(self):
        """지적이 사업계획 '산출내역' 표에서 났으면 그 표를 보여 주고 그 줄을 칠해야 한다.

        예전에는 상세 화면에 '요구내용 및 산출근거' 표만 나왔다. 뒤쪽 표에서 난 지적을
        누르면 멀쩡한 줄만 가득한 표가 떠서, 담당자는 어디가 틀렸는지 찾을 수 없었다.
        """
        from crosscheck import Issue
        from plan_parser import PlanItem, PlanProject

        app = self.app
        project = PlanProject(name="표본사업", request=10000, year_request=10000,
                              had_own_items=True)
        project.items = [PlanItem(depth=1, name="절", amount=10000, row=1),
                         PlanItem(depth=2, name="가", amount=10000,
                                  formula="10,000,000원×1건", row=2)]
        project.outline_items = [[
            PlanItem(depth=1, name="절", amount=10000, row=1),
            PlanItem(depth=2, name="가", amount=10000, formula="10,000천원×2건", row=2)]]
        app.session.projects = [project]
        app.issues = [("목표1", Issue("목표1", "오류", "표본사업", "가",
                                    "산출내역 표 2행 '가' — 산출식이 금액과 다릅니다.",
                                    10000, 20000, index=None,
                                    form="산출내역 표", block=0, row=2))]
        app.current = 0
        app.open_detail()
        app.update()
        tree = app.detail_tree
        rows = [(tree.item(one, "tags"), tree.item(one, "values"))
                for one in tree.get_children()]
        self.assertTrue(any("산출내역 표" in str(values[2]) for _tags, values in rows),
                        "산출내역 표가 상세 화면에 나오지 않는다")
        mark = next(index for index, (_tags, values) in enumerate(rows)
                    if "산출내역 표" in str(values[2]))
        painted = [values for tags, values in rows[mark:] if "error" in tags]
        self.assertEqual(len(painted), 1, painted)
        self.assertEqual(str(painted[0][4]), "10,000천원×2건")
        app.session.projects = []
        app.issues = []

    def test_later_steps_explain_themselves_before_loading(self):
        """자료를 불러오기 전 뒤 단계 탭은 빈 표가 아니라 '무엇을 먼저 하라'를 보여야 한다."""
        app = self.app
        app.session.projects = []
        for page in ('result', 'bimok', 'make', 'verify'):
            app.show(page)
            app.update()
            self.assertTrue(app.pages[page].cover.winfo_ismapped(), page)
        app.show('files')

    def test_code_less_rows_are_offered_on_the_bimok_screen(self):
        """목코드 없는 줄(2027 설명서 전부)이 비목 화면에 나와야 한다. 예전엔 0줄이었다."""
        from bimok_resolver import BimokResolver
        from converter import Session
        from crosscheck import check_bimok
        from plan_parser import PlanItem, PlanProject

        app = self.app
        project = PlanProject(name="표본사업", request=200, year_request=200, had_own_items=True)
        project.items = [PlanItem(depth=1, name="운영", amount=200, row=1),
                         PlanItem(depth=2, name="운영용품", amount=100, formula="100,000원×1회", row=2),
                         PlanItem(depth=2, name="간식비", amount=100, formula="100,000원×1회", row=3)]
        session = Session()
        session.projects = [project]
        session.resolver = BimokResolver()
        session.bimok_report, session.resolutions = check_bimok(session.projects, session.resolver)
        app.session = session
        app.refresh()
        app.show('bimok')
        app.update()
        self.assertFalse(app.pages['bimok'].cover.winfo_ismapped())
        tree = app.bimok_tree
        rows = [tree.item(one, 'values') for one in tree.get_children()]
        self.assertTrue(any('목코드 없음' in str(values[0]) for values in rows), rows)
        self.assertEqual(sum(1 for values in rows if '└' in str(values[0])), 2, rows)
        self.assertIn('2행', app.bimok_headline.cget('text'))
        # 점검표: 비목 미확정은 막지 않고(노랑), 금액 검산은 통과(초록)
        app.show('make')
        app.update()
        self.assertEqual(app.make_button.cget('state'), 'normal')
        app.session = Session()
        app.show('files')
        self.assertFalse(self.errors)

    def test_failed_load_keeps_the_previous_result(self):
        """다시 불러오다 실패하면 화면의 지난 결과와 세션이 짝을 유지해야 한다."""
        from converter import Session
        from plan_parser import PlanProject

        app = self.app
        previous = Session()
        previous.projects = [PlanProject(name="지난사업")]
        app.session = previous
        import tempfile, os
        folder = tempfile.mkdtemp()
        broken = os.path.join(folder, '깨진설명서.hwpx')
        with open(broken, 'wb') as handle:
            handle.write(b'not a zip')
        app.paths['plan'].set(broken)
        app.paths['last_year'].set(broken)
        with patch('budget_bridge.messagebox.showerror') as shown:
            app.load()
            for _ in range(200):
                app.update()
                if not app.loading:
                    break
                import time
                time.sleep(0.02)
            self.assertTrue(shown.called)
        self.assertIs(app.session, previous)
        app.session = Session()
        app.clear_all()

    def test_missing_remembered_file_is_named_before_loading(self):
        """기억해 둔 파일이 옮겨졌으면 불러오기 전에 한국어로 어느 칸인지 알려야 한다."""
        app = self.app
        app.paths['plan'].set('/없는/폴더/2027사업별 설명서.hwpx')
        app.paths['last_year'].set('/없는/폴더/2026(k에듀파인).xlsx')
        app._mark_files()
        self.assertIn('✕', app.file_view.states['plan'].cget('text'))
        with patch('budget_bridge.messagebox.showwarning') as warned:
            app.load()
        self.assertFalse(app.loading, '없는 파일로 읽기를 시작하면 안 된다')
        self.assertIn('1번 사업설명서', warned.call_args[0][1])
        app.clear_all()

    def test_issue_opens_its_source_files(self):
        """지적을 고르면 맞춰 볼 원본 파일 버튼이 나오고, 누르면 찾을 글자를 복사한 뒤 연다.

        설명서 검산은 설명서만, UBIS 대조는 설명서와 UBIS 검토조서 둘 다. 파일이 그새
        옮겨졌으면 열지 않고 어느 파일인지 말한다.
        """
        import os
        import tempfile
        from converter import Session
        from crosscheck import Issue
        from plan_parser import PlanProject

        app = self.app
        folder = tempfile.mkdtemp()
        plan, ubis = os.path.join(folder, '설명서.hwpx'), os.path.join(folder, '검토조서.xlsx')
        for path in (plan, ubis):
            open(path, 'wb').close()
        session = Session(plan_path=plan, ubis_path=ubis)
        session.projects = [PlanProject(name='표본사업', heading='3. 표본사업')]
        session.ubis = {'표본 사업': 100.0}      # UBIS 는 띄어 쓴다. 찾을 글자는 UBIS 표기여야 한다
        app.session = session
        app.issues = [('목표1', Issue('목표1', '오류', '표본사업', '', '검산 오류', 1, 2)),
                      ('목표2A', Issue('목표2A', '오류', '표본사업', '', 'UBIS 차이', 1, 2))]
        app.show('result')
        app.issue_tree.delete(*app.issue_tree.get_children())
        for _goal, issue in app.issues:
            app.issue_tree.insert('', 'end', values=(issue.severity, '', issue.project, issue.message, '', '', ''))
        labels = lambda: [one.cget('text') for one in app.preview_actions.winfo_children()]

        app._select(0)
        app.update()
        self.assertEqual(labels(), ['설명서 열기'])
        app._select(1)
        app.update()
        self.assertEqual(labels(), ['설명서 열기', 'UBIS 검토조서 열기'])

        with patch('budget_bridge.open_path') as opened:
            app.open_source('ubis')
            opened.assert_called_once_with(ubis)
            self.assertEqual(app.clipboard_get(), '표본 사업')
            app.open_source('plan')
            self.assertEqual(opened.call_args[0][0], plan)
            self.assertEqual(app.clipboard_get(), session.projects[0].find_text())
        self.assertIn('Ctrl+F', app.status.cget('text'))

        app.open_detail()
        app.update()
        self.assertEqual([one.cget('text') for one in app.detail_actions.winfo_children()],
                         ['설명서 열기', 'UBIS 검토조서 열기'])

        os.remove(ubis)
        with patch('budget_bridge.open_path') as opened, \
                patch('budget_bridge.messagebox.showwarning') as warned:
            app.open_source('ubis')
        opened.assert_not_called()
        self.assertTrue(warned.called)
        app.session = Session()
        app.issues = []
        app.show('files')
        self.assertFalse(self.errors)

    def test_writing_checks_have_their_own_list(self):
        """작성 점검은 오류 목록에 섞이지 않고 '작성 점검' 구분에 따로 모인다."""
        from converter import Session
        from crosscheck import Issue
        from plan_parser import PlanProject
        from writing_check import WritingReport

        app = self.app
        session = Session()
        session.projects = [PlanProject(name='표본사업', heading='3. 표본사업')]
        session.writing_report = WritingReport([Issue('작성-표기', '확인 필요', '표본사업', '편성·운영',
                                                      '가운뎃점은 ․ 를 씁니다.', blocking=False)])
        app.session = session
        app.show('result')
        app.set_filter('오류')
        app.update()
        self.assertEqual(app.issues, [])
        self.assertIn('작성 점검  1', app.filter_buttons['작성 점검'].cget('text'))
        app.set_filter('작성 점검')
        app.update()
        self.assertEqual(len(app.issues), 1)
        app._select(0)
        app.update()
        self.assertEqual(app._find_for('plan', app.issues[0][1]), '편성·운영', '걸린 글자로 찾아가야 한다')
        app.open_detail()
        app.update()
        self.assertEqual(app.find_value.get(), '편성·운영')
        app.session = Session()
        app.issues = []
        app.set_filter('오류')
        app.show('files')
        self.assertFalse(self.errors)

    def test_flow_arrows_are_bold(self):
        """단계 사이 화살표는 굵고 곧아야 한다. 가는 연한 물결선은 흐름으로 읽히지 않았다."""
        flow = self.app.file_view.flow
        self.app.show('files')
        self.app.update()
        flow.set_stage(1, {0})
        lines = [one for one in flow.find_all() if flow.type(one) == 'line']
        self.assertEqual(len(lines), 4)
        for one in lines:
            self.assertGreaterEqual(float(flow.itemcget(one, 'width')), 3)
            self.assertEqual(flow.itemcget(one, 'arrow'), 'last')
            self.assertEqual(len(flow.coords(one)), 4, '곧은 선이어야 한다')

    def test_guide_looks_different_from_work_cards(self):
        """안내 말풍선은 작업 카드(흰 바탕)와 다른 바탕이어야 한다."""
        from native_ui import RoundedPanel
        from ui_common import GUIDE_BG, PANEL

        page = self.app.file_view
        page.open_guide(5)
        self.app.update()
        panel = page.coach.winfo_children()[1]
        self.assertIsInstance(panel, RoundedPanel)
        self.assertEqual(panel.fill, GUIDE_BG)
        self.assertNotEqual(panel.fill, PANEL)
        self.assertTrue(all(row.fill == PANEL for row in page.rows.values()))
        page.close_guide(False)

    def test_rounded_button_disabled_and_keyboard(self):
        calls = []
        b = RoundedButton(self.app, '확인', lambda: calls.append(1))
        b.pack()
        b.configure(state='disabled')
        b.invoke()
        self.assertEqual(calls, [])
        b.configure(state='normal', text='다음')
        b.invoke()
        self.assertEqual(calls, [1])
        self.assertEqual(b.cget('text'), '다음')
        b.destroy()

    def test_all_original_pages_open(self):
        for page in ('files','result','bimok','make','verify'):
            self.app.show(page)
            self.app.update()
            self.assertEqual(self.app.page, page)
        self.app.show('files')
        self.assertFalse(self.errors)


if __name__ == '__main__':
    unittest.main()
