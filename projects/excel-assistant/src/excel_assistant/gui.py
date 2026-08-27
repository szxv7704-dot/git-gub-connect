from __future__ import annotations

from pathlib import Path
import sys
import tempfile

from PySide6.QtCore import QObject, QRunnable, QThreadPool, Qt, Signal, Slot
from PySide6.QtWidgets import (
    QApplication, QComboBox, QFileDialog, QFrame, QGridLayout, QHBoxLayout,
    QLabel, QLineEdit, QMainWindow, QMessageBox, QProgressBar, QPushButton, QTableWidget,
    QTableWidgetItem, QTextEdit, QVBoxLayout, QWidget,
)

from .core import RequestError, TaskPlan, analyze_request
from .workbook import execute_plan, inspect_workbook


class WorkerSignals(QObject):
    finished = Signal(object)
    failed = Signal(str)


class BackgroundJob(QRunnable):
    def __init__(self, function, *args) -> None:
        super().__init__()
        self.function, self.args = function, args
        self.signals = WorkerSignals()

    @Slot()
    def run(self) -> None:
        try:
            self.signals.finished.emit(self.function(*self.args))
        except Exception as exc:
            self.signals.failed.emit(str(exc))


def _preview_job(source, source_sheet, header_row, end_row, plan, temp_path):
    try:
        return execute_plan(source, source_sheet, header_row, plan, temp_path, end_row)
    finally:
        try:
            Path(temp_path).unlink(missing_ok=True)
        except OSError:
            pass


class ExcelAssistantApp(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("엑셀 서식 도우미")
        self.resize(1080, 760)
        self.setMinimumSize(880, 640)
        self.source_path: Path | None = None
        self.workbook_info: dict = {}
        self.plan: TaskPlan | None = None
        self.thread_pool = QThreadPool.globalInstance()
        self._jobs: list[BackgroundJob] = []
        self._build_ui()

    @staticmethod
    def _panel() -> tuple[QFrame, QGridLayout]:
        frame = QFrame()
        frame.setProperty("panel", True)
        grid = QGridLayout(frame)
        grid.setContentsMargins(14, 12, 14, 12)
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(7)
        return frame, grid

    def _build_ui(self) -> None:
        self.setStyleSheet(
            "QMainWindow{background:#f4f7f5;}"
            "QFrame[panel='true']{background:white;border:1px solid #d5ded8;border-radius:8px;}"
            "QPushButton{padding:8px 12px;}"
            "QPushButton[primary='true']{background:#17643f;color:white;border:0;border-radius:6px;font-weight:600;}"
            "QComboBox,QTextEdit{background:white;border:1px solid #aebbb3;border-radius:5px;padding:6px;}"
            "QTableWidget{background:white;border:1px solid #d5ded8;gridline-color:#e4e9e6;}"
            "QHeaderView::section{background:#edf4ef;padding:7px;border:0;border-bottom:1px solid #ced9d2;font-weight:600;}"
        )
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(12)

        top = QHBoxLayout()
        title = QLabel("엑셀 서식 도우미")
        title.setStyleSheet("font-size:22px;font-weight:700;color:#173c2a;")
        top.addWidget(title)
        top.addStretch()
        open_button = QPushButton("엑셀 파일 열기")
        open_button.clicked.connect(self.open_file)
        top.addWidget(open_button)
        layout.addLayout(top)

        file_panel, file_layout = self._panel()
        file_layout.addWidget(QLabel("1. 원본 데이터"), 0, 0)
        self.file_label = QLabel("아직 파일을 선택하지 않았습니다.")
        file_layout.addWidget(self.file_label, 1, 0)
        file_layout.addWidget(QLabel("시트"), 0, 1)
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self._sheet_changed)
        file_layout.addWidget(self.sheet_combo, 1, 1)
        file_layout.setColumnStretch(0, 1)
        layout.addWidget(file_panel)

        request_panel, request_layout = self._panel()
        request_layout.addWidget(QLabel("2. 원하는 결과"), 0, 0, 1, 2)
        self.request_text = QTextEdit("이름을 기준으로 금액을 합산하고 큰 순서대로 보여줘")
        self.request_text.setFixedHeight(72)
        request_layout.addWidget(self.request_text, 1, 0, 1, 2)
        examples = QHBoxLayout()
        for label, text in (
            ("이름별 금액 합계", "이름을 기준으로 금액을 합산하고 큰 순서대로 보여줘"),
            ("완료 항목만", "상태가 완료인 행만 보여줘"),
            ("금액 큰 순", "금액을 큰 순서로 정렬해줘"),
            ("이름 중복 찾기", "이름 중복 항목을 찾아줘"),
        ):
            button = QPushButton(label)
            button.clicked.connect(lambda _checked=False, value=text: self.request_text.setPlainText(value))
            examples.addWidget(button)
        examples.addStretch()
        analyze_button = QPushButton("요청 분석")
        analyze_button.setProperty("primary", True)
        analyze_button.clicked.connect(self.analyze)
        examples.addWidget(analyze_button)
        request_layout.addLayout(examples, 2, 0, 1, 2)
        layout.addWidget(request_panel)

        plan_panel, plan_layout = self._panel()
        plan_layout.addWidget(QLabel("3. 작업 확인"), 0, 0, 1, 3)
        self.plan_label = QLabel("파일을 열고 요청을 분석하면 적용할 작업이 표시됩니다.")
        self.plan_label.setWordWrap(True)
        plan_layout.addWidget(self.plan_label, 1, 0, 1, 3)
        plan_layout.addWidget(QLabel("기준 열"), 2, 0)
        plan_layout.addWidget(QLabel("계산·대상 열"), 2, 1)
        self.group_combo, self.value_combo = QComboBox(), QComboBox()
        plan_layout.addWidget(self.group_combo, 3, 0)
        plan_layout.addWidget(self.value_combo, 3, 1)
        preview_button = QPushButton("선택한 열로 미리보기")
        preview_button.clicked.connect(self.preview)
        plan_layout.addWidget(preview_button, 3, 2)
        plan_layout.addWidget(QLabel("추가 기준 열 / 더할 열 (쉼표로 구분)"), 4, 0)
        plan_layout.addWidget(QLabel("조건값"), 4, 1)
        self.extra_columns = QLineEdit()
        self.extra_columns.setPlaceholderText("예: 부서, 세부사업 또는 중앙, 광역, 기초")
        self.condition_value = QLineEdit()
        self.condition_value.setPlaceholderText("필터 작업의 조건값")
        plan_layout.addWidget(self.extra_columns, 5, 0)
        plan_layout.addWidget(self.condition_value, 5, 1)
        plan_layout.setColumnStretch(0, 1)
        plan_layout.setColumnStretch(1, 1)
        layout.addWidget(plan_panel)

        preview_title = QLabel("4. 결과 미리보기")
        preview_title.setStyleSheet("font-weight:600;")
        layout.addWidget(preview_title)
        self.preview_table = QTableWidget()
        self.preview_table.setAlternatingRowColors(True)
        self.preview_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.preview_table.setSelectionBehavior(QTableWidget.SelectRows)
        layout.addWidget(self.preview_table, 1)

        bottom = QHBoxLayout()
        self.status_label = QLabel("준비")
        bottom.addWidget(self.status_label)
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setFixedWidth(130)
        self.progress.hide()
        bottom.addWidget(self.progress)
        bottom.addStretch()
        self.save_button = QPushButton("결과 엑셀 저장")
        self.save_button.setProperty("primary", True)
        self.save_button.setEnabled(False)
        self.save_button.clicked.connect(self.save_result)
        bottom.addWidget(self.save_button)
        layout.addLayout(bottom)

    def open_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "엑셀 파일 선택", "", "지원 파일 (*.xlsx *.xlsm *.csv);;Excel 통합문서 (*.xlsx *.xlsm);;CSV (*.csv)")
        if not path:
            return
        self.status_label.setText("파일 구조와 표를 분석하고 있습니다…")
        self._start_job(inspect_workbook, (path,), lambda info: self._file_opened(path, info), "파일 열기 실패")

    def _file_opened(self, path: str, info: dict) -> None:
        self.workbook_info = info
        self.source_path = Path(path)
        self.sheet_combo.blockSignals(True)
        self.sheet_combo.clear()
        self.sheet_combo.addItems(list(self.workbook_info))
        self.sheet_combo.blockSignals(False)
        self.file_label.setText(self.source_path.name)
        self._sheet_changed(self.sheet_combo.currentText())

    def _sheet_changed(self, sheet_name: str) -> None:
        info = self.workbook_info.get(sheet_name, {})
        headers = info.get("headers", [])
        for combo in (self.group_combo, self.value_combo):
            combo.clear()
            combo.addItems(headers)
        if len(headers) > 1:
            self.value_combo.setCurrentIndex(1)
        self.extra_columns.clear()
        self.condition_value.clear()
        self.plan = None
        self.save_button.setEnabled(False)
        self.preview_table.clear()
        message = f"{info.get('rows', 0):,}행 · {len(headers)}개 열"
        if info.get("table_count", 1) > 1:
            message += f" · 이 시트에서 표 {info['table_count']}개 감지"
        if info.get("uncached_formula_count", 0):
            message += f" · 계산값 없는 수식 {info['uncached_formula_count']}개 주의"
        self.status_label.setText(message)

    def analyze(self) -> None:
        if not self.source_path:
            QMessageBox.information(self, "파일 필요", "먼저 엑셀 파일을 선택해 주세요.")
            return
        info = self.workbook_info[self.sheet_combo.currentText()]
        try:
            self.plan = analyze_request(self.request_text.toPlainText().strip(), info["headers"], info.get("profiles"))
        except RequestError as exc:
            QMessageBox.warning(self, "요청 확인", str(exc))
            return
        self.extra_columns.clear()
        self.condition_value.clear()
        self.plan_label.setText(f"{self.plan.operation_label} · {self.plan.explanation} · 열 매칭 신뢰도 {self.plan.confidence:.0%}")
        first = self.plan.group_column or self.plan.filter_column or self.plan.sort_column
        if first:
            self.group_combo.setCurrentText(first)
        if self.plan.value_column:
            self.value_combo.setCurrentText(self.plan.value_column)
        if self.plan.operation == "group_sum" and len(self.plan.group_columns) > 1:
            self.extra_columns.setText(", ".join(self.plan.group_columns))
        elif self.plan.operation == "sum_check":
            self.extra_columns.setText(", ".join(self.plan.component_columns))
            if self.plan.total_column:
                self.value_combo.setCurrentText(self.plan.total_column)
        elif self.plan.operation == "filter":
            self.condition_value.setText(str(self.plan.filter_value))
        if self.plan.ambiguous_columns:
            QMessageBox.warning(self, "비슷한 열 확인 필요", f"선택한 열과 비슷한 후보가 있습니다: {', '.join(self.plan.ambiguous_columns)}\n아래 열 선택을 확인한 뒤 미리보기를 실행해 주세요.")
            return
        if self.plan.confidence < 0.65:
            QMessageBox.warning(self, "열 선택 필요", "요청과 열 제목을 안전하게 연결하지 못했습니다. 아래 기준 열과 계산 열을 직접 선택한 뒤 미리보기를 실행해 주세요.")
            return
        if self.plan.confidence < 0.8:
            QMessageBox.warning(self, "열 확인 필요", "요청과 열 제목의 일치도가 낮습니다. 기준 열을 확인한 뒤 미리보기를 실행해 주세요.")
        self.preview()

    def _apply_selected_columns(self) -> None:
        if not self.plan:
            return
        first, second = self.group_combo.currentText(), self.value_combo.currentText()
        if self.plan.operation == "group_sum":
            self.plan.group_column, self.plan.value_column = first, second
            selected = tuple(value.strip() for value in self.extra_columns.text().split(",") if value.strip())
            self.plan.group_columns = selected or (first,)
            self.plan.group_column = self.plan.group_columns[0]
        elif self.plan.operation == "filter":
            self.plan.filter_column = first
            raw = self.condition_value.text().strip()
            if self.plan.filter_operator in {"gt", "gte", "lt", "lte"}:
                from .core import coerce_number
                self.plan.filter_value = coerce_number(raw)
            else:
                self.plan.filter_value = raw
        elif self.plan.operation == "sort":
            self.plan.sort_column = first
        elif self.plan.operation == "duplicates":
            self.plan.group_column = first
        elif self.plan.operation == "sum_check":
            self.plan.component_columns = tuple(value.strip() for value in self.extra_columns.text().split(",") if value.strip())
            self.plan.total_column = second

    def preview(self) -> None:
        if not self.source_path or not self.plan:
            return
        self._apply_selected_columns()
        handle = tempfile.NamedTemporaryFile(prefix="excel_assistant_preview_", suffix=".xlsx", delete=False)
        temp_path = Path(handle.name)
        handle.close()
        info = self.workbook_info[self.sheet_combo.currentText()]
        self.status_label.setText("결과를 계산하고 있습니다…")
        self._start_job(
            _preview_job,
            (self.source_path, info.get("source_sheet", self.sheet_combo.currentText()), info["header_row"], info.get("end_row"), self.plan, temp_path),
            self._preview_ready, "미리보기 실패",
        )

    def _preview_ready(self, result) -> None:
            headers, rows, stats = result
            self._show_preview(headers, rows[:100])
            self.status_label.setText(f"원본 {stats['source_rows']:,}행 → 결과 {stats['result_rows']:,}행 · 제외 {stats['excluded_rows']:,}행")
            if stats.get("statistics") and "maximum" in stats["statistics"]:
                item = stats["statistics"]
                self.plan_label.setText(
                    f"{self.plan.explanation}  |  최댓값 {item['maximum_group']} {item['maximum']:,.0f} · "
                    f"최솟값 {item['minimum_group']} {item['minimum']:,.0f} · 평균 {item['average']:,.0f}"
                )
            elif self.plan.operation == "sum_check":
                self.plan_label.setText(f"{self.plan.explanation}  |  불일치 {stats['statistics']['mismatch_count']:,}건")
            self.save_button.setEnabled(True)

    def _show_preview(self, headers: list[str], rows: list[list]) -> None:
        self.preview_table.clear()
        self.preview_table.setColumnCount(len(headers))
        self.preview_table.setHorizontalHeaderLabels(headers)
        self.preview_table.setRowCount(len(rows))
        for row_index, row in enumerate(rows):
            for col_index, value in enumerate(row):
                display = f"{value:,.0f}" if isinstance(value, float) else ("" if value is None else str(value))
                item = QTableWidgetItem(display)
                if isinstance(value, (int, float)):
                    item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                self.preview_table.setItem(row_index, col_index, item)
        self.preview_table.resizeColumnsToContents()

    def save_result(self) -> None:
        if not self.source_path or not self.plan:
            return
        default = str(self.source_path.with_name(f"{self.source_path.stem}_작업결과.xlsx"))
        path, _ = QFileDialog.getSaveFileName(self, "결과 엑셀 저장", default, "Excel 통합문서 (*.xlsx)")
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        self._apply_selected_columns()
        info = self.workbook_info[self.sheet_combo.currentText()]
        self.status_label.setText("결과 엑셀을 저장하고 있습니다…")
        self._start_job(
            execute_plan,
            (self.source_path, info.get("source_sheet", self.sheet_combo.currentText()), info["header_row"], self.plan, path, info.get("end_row")),
            lambda result: self._save_ready(path, result), "저장 실패",
        )

    def _save_ready(self, path: str, result) -> None:
            _, _, stats = result
            QMessageBox.information(self, "저장 완료", f"결과 파일을 저장했습니다.\n\n{path}\n\n결과 {stats['result_rows']:,}행")

    def _start_job(self, function, args: tuple, on_success, error_title: str) -> None:
        job = BackgroundJob(function, *args)
        self._jobs.append(job)
        self.centralWidget().setEnabled(False)
        self.progress.show()
        job.signals.finished.connect(lambda result: self._job_finished(job, on_success, result))
        job.signals.failed.connect(lambda message: self._job_failed(job, error_title, message))
        self.thread_pool.start(job)

    def _job_finished(self, job: BackgroundJob, callback, result) -> None:
        if job in self._jobs:
            self._jobs.remove(job)
        if not self._jobs:
            self.progress.hide()
            self.centralWidget().setEnabled(True)
        callback(result)

    def _job_failed(self, job: BackgroundJob, title: str, message: str) -> None:
        if job in self._jobs:
            self._jobs.remove(job)
        if not self._jobs:
            self.progress.hide()
            self.centralWidget().setEnabled(True)
        QMessageBox.critical(self, title, message)


def run() -> None:
    app = QApplication.instance() or QApplication(sys.argv)
    window = ExcelAssistantApp()
    window.show()
    app.exec()
