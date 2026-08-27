from __future__ import annotations
import sys
from pathlib import Path
from PySide6.QtCore import QObject, QThread, Signal, Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (QAbstractItemView,QApplication,QDialog,QFileDialog,QFrame,QHBoxLayout,QHeaderView,QLabel,QListWidget,QMainWindow,QMessageBox,QPlainTextEdit,QPushButton,QRadioButton,QSplitter,QTableWidget,QTableWidgetItem,QVBoxLayout,QWidget)
from privacy_core import Candidate, collect_files, process_file, replacement_for, scan_file, unsupported_files, write_report
from preview_ui import ExcelPreviewDialog, VisualMaskDialog

APP_VERSION = "2026.08.26.2"

class Worker(QObject):
    progress=Signal(str); scanned=Signal(object,object); processed=Signal(object,object,object); failed=Signal(str)
    def scan(self,files,custom):
        try:
            candidates=[]; warnings={}
            for i,path in enumerate(files,1):
                try: found,notes=scan_file(path,custom); candidates+=found; warnings[path]=notes
                except Exception as exc: warnings[path]=[f"분석 실패: {exc}"]
                self.progress.emit(f"분석 중 {i}/{len(files)} · {path.name}")
            self.scanned.emit(candidates,warnings)
        except Exception as exc:
            # 여기서 신호를 놓치면 화면의 버튼이 영영 잠긴 채로 남는다.
            self.failed.emit(f"분석을 끝내지 못했습니다: {exc}")
    def process(self,files,candidates,warnings,mode,output):
        try:
            aliases={}; results=[]
            for i,path in enumerate(files,1):
                result=process_file(path,candidates,mode,output,aliases)
                if warnings.get(path): result.message += " / " + "; ".join(warnings[path])
                results.append(result); self.progress.emit(f"처리 중 {i}/{len(files)} · {path.name}")
            self.processed.emit(results,write_report(results,output),output)
        except Exception as exc:
            self.failed.emit(f"결과 파일을 만들지 못했습니다: {exc}")

class PrivacyApp(QMainWindow):
    def __init__(self):
        super().__init__(); self.setWindowTitle(f"개인정보 비식별 처리기 · 오프라인 · {APP_VERSION}"); self.resize(1360,840); self.setMinimumSize(1080,700)
        self.files=[]; self.candidates=[]; self.warnings={}; self.thread=None; self.failed_files=[]; self.last_output=None; self._build()
    def _build(self):
        self.setStyleSheet("""
            QMainWindow, QWidget#page { background:#F4F7FB; color:#172033; font-family:'Malgun Gothic'; font-size:13px; }
            QFrame#header { background:#173B67; border-radius:14px; }
            QFrame.card { background:white; border:1px solid #DCE4EE; border-radius:14px; }
            QLabel#title { color:white; font-size:26px; font-weight:700; }
            QLabel#subtitle { color:#D8E7F7; font-size:13px; }
            QLabel.step { color:#173B67; font-size:16px; font-weight:700; }
            QLabel.hint { color:#64748B; font-size:12px; }
            QLabel#safeBadge { color:#166534; background:#DCFCE7; border-radius:11px; padding:5px 10px; font-weight:700; }
            QPushButton { min-height:34px; padding:0 15px; border:1px solid #C8D4E3; border-radius:8px; background:white; color:#24364B; font-weight:600; }
            QPushButton:hover { background:#EFF6FF; border-color:#7BA7D9; }
            QPushButton#primary { background:#1769AA; color:white; border:0; min-height:42px; font-size:14px; }
            QPushButton#primary:hover { background:#0F5B96; }
            QPushButton#save { background:#0F766E; color:white; border:0; min-height:42px; font-size:14px; padding:0 24px; }
            QPushButton#save:disabled, QPushButton#primary:disabled { background:#AAB7C4; }
            QListWidget, QPlainTextEdit, QTableWidget { background:white; border:1px solid #D5DFEA; border-radius:9px; selection-background-color:#DCEEFF; selection-color:#172033; }
            QListWidget { padding:6px; }
            QPlainTextEdit { padding:8px; }
            QHeaderView::section { background:#EEF3F8; color:#334155; border:0; border-bottom:1px solid #CBD5E1; padding:9px; font-weight:700; }
            QTableWidget { gridline-color:#E7EDF4; alternate-background-color:#F8FAFC; }
            QRadioButton { spacing:7px; }
            QSplitter::handle { background:#E2E8F0; width:2px; margin:8px 5px; }
        """)
        main=QWidget(); main.setObjectName("page"); self.setCentralWidget(main); root=QVBoxLayout(main); root.setContentsMargins(20,18,20,18); root.setSpacing(14)
        header=QFrame(); header.setObjectName("header"); hh=QHBoxLayout(header); hh.setContentsMargins(22,16,22,16)
        head_text=QVBoxLayout(); title=QLabel("개인정보 비식별 처리기"); title.setObjectName("title"); head_text.addWidget(title)
        sub=QLabel("파일은 외부로 전송되지 않으며 원본은 그대로 보존됩니다."); sub.setObjectName("subtitle"); head_text.addWidget(sub); hh.addLayout(head_text); hh.addStretch()
        badge=QLabel("● 오프라인 안전 처리"); badge.setObjectName("safeBadge"); hh.addWidget(badge); root.addWidget(header)
        guide=QLabel("① 파일 선택     →     ② 마스킹 항목 입력     →     ③ 후보 검토     →     ④ 새 파일 저장")
        guide.setStyleSheet("color:#46627F; font-weight:600; padding:2px 8px"); root.addWidget(guide)
        split=QSplitter(); root.addWidget(split,1)
        left=QFrame(); left.setProperty("class","card"); ll=QVBoxLayout(left); ll.setContentsMargins(16,15,16,16); ll.setSpacing(9)
        file_head=QHBoxLayout(); step1=QLabel("1  입력 파일"); step1.setProperty("class","step"); file_head.addWidget(step1); file_head.addStretch(); self.file_count=QLabel("0개"); self.file_count.setProperty("class","hint"); file_head.addWidget(self.file_count); ll.addLayout(file_head)
        tools=QHBoxLayout()
        for text,fn in [("＋ 파일",self.add_files),("＋ 폴더",self.add_folder),("비우기",self.clear)]:
            b=QPushButton(text); b.clicked.connect(fn); tools.addWidget(b)
        tools.addStretch(); ll.addLayout(tools)
        self.file_list=QListWidget(); self.file_list.setWordWrap(False); self.file_list.setUniformItemSizes(True); self.file_list.setToolTip("분석할 파일 목록"); ll.addWidget(self.file_list,1)
        step2=QLabel("2  마스킹할 항목 또는 값"); step2.setProperty("class","step"); ll.addWidget(step2)
        self.custom=QPlainTextEdit(); self.custom.setPlaceholderText("한 줄에 하나씩 입력하세요.\n예) 성명\n     생년월일\n     계좌번호\n     홍길동"); self.custom.setMaximumHeight(145); ll.addWidget(self.custom)
        hint=QLabel("비워두면 지원하는 개인정보 유형을 모두 검사합니다. 입력하면 해당 항목·값만 표시합니다."); hint.setWordWrap(True); hint.setProperty("class","hint"); ll.addWidget(hint)
        self.scan_button=QPushButton("개인정보 후보 찾기"); self.scan_button.setObjectName("primary"); self.scan_button.clicked.connect(self.start_scan); ll.addWidget(self.scan_button); split.addWidget(left)
        right=QFrame(); right.setProperty("class","card"); rl=QVBoxLayout(right); rl.setContentsMargins(16,15,16,16); rl.setSpacing(10)
        review_head=QHBoxLayout(); step3=QLabel("3  마스킹 후보 검토"); step3.setProperty("class","step"); review_head.addWidget(step3); review_head.addStretch(); self.candidate_count=QLabel("후보 0건"); self.candidate_count.setProperty("class","hint"); review_head.addWidget(self.candidate_count); rl.addLayout(review_head)
        review_hint=QLabel("저신뢰도 후보는 자동 적용하지 않습니다. 행을 선택해 미리보기에서 확인한 뒤 승인하세요."); review_hint.setProperty("class","hint"); rl.addWidget(review_hint)
        preview_tools=QHBoxLayout(); preview=QPushButton("선택 항목 미리보기·영역 편집"); preview.clicked.connect(self.review_selected); preview_tools.addWidget(preview); approve=QPushButton("검토 완료·적용"); approve.clicked.connect(self.approve_selected); preview_tools.addWidget(approve); preview_tools.addStretch(); rl.addLayout(preview_tools)
        self.table=QTableWidget(0,7); self.table.setAlternatingRowColors(True); self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows); self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers); self.table.verticalHeader().setDefaultSectionSize(34)
        self.table.setHorizontalHeaderLabels(["적용","검토","유형","보호된 값","파일","위치","신뢰도"]); header_view=self.table.horizontalHeader(); header_view.setSectionResizeMode(QHeaderView.ResizeMode.Interactive); header_view.setStretchLastSection(False)
        for column,width in enumerate([62,76,105,135,175,165,70]): self.table.setColumnWidth(column,width)
        header_view.setSectionResizeMode(4,QHeaderView.ResizeMode.Stretch); self.table.cellDoubleClicked.connect(self.table_double_clicked); rl.addWidget(self.table,1)
        choices=QHBoxLayout(); on=QPushButton("전체 선택"); on.clicked.connect(lambda:self.set_all(True)); choices.addWidget(on); off=QPushButton("전체 해제"); off.clicked.connect(lambda:self.set_all(False)); choices.addWidget(off); choices.addSpacing(18)
        choices.addWidget(QLabel("처리 방식")); self.alias=QRadioButton("일관 가명"); self.alias.setChecked(True); choices.addWidget(self.alias); self.label_ko=QRadioButton("한글 라벨"); choices.addWidget(self.label_ko); self.label_en=QRadioButton("영문 라벨"); choices.addWidget(self.label_en); self.delete=QRadioButton("검정 블록 완전 삭제"); choices.addWidget(self.delete); choices.addStretch(); rl.addLayout(choices)
        split.addWidget(right); split.setSizes([390,930])
        footer=QFrame(); footer.setProperty("class","card"); fl=QHBoxLayout(footer); fl.setContentsMargins(16,10,12,10); self.status=QLabel("파일 또는 폴더를 추가해 시작하세요."); self.status.setStyleSheet("color:#475569"); fl.addWidget(self.status,1)
        self.retry_button=QPushButton("실패 파일만 재실행"); self.retry_button.setEnabled(False); self.retry_button.clicked.connect(self.retry_failed); fl.addWidget(self.retry_button)
        self.process_button=QPushButton("4  검증 후 새 파일로 저장"); self.process_button.setObjectName("save"); self.process_button.setEnabled(False); self.process_button.clicked.connect(self.start_process); fl.addWidget(self.process_button); root.addWidget(footer)
    def add_files(self):
        paths,_=QFileDialog.getOpenFileNames(self,"파일 추가","","지원 파일 (*.xlsx *.xls *.pdf *.png *.jpg *.jpeg *.hwpx);;모든 파일 (*.*)"); self._add(paths)
    def add_folder(self):
        path=QFileDialog.getExistingDirectory(self,"폴더 추가")
        if path:self._add([path])
    def _add(self,paths):
        self.files=sorted(set(self.files)|set(collect_files(paths))); self.file_list.clear(); self.file_list.addItems([str(p) for p in self.files])
        self.file_count.setText(f"{len(self.files)}개")
        skipped=unsupported_files(paths); self.status.setText(f"지원 파일 {len(self.files)}개가 선택되었습니다.")
        if skipped:
            # 지원하지 않는 파일이 조용히 빠지면 처리된 줄 알고 그대로 배포하게 된다.
            names="\n".join(f"• {p.name}" for p in skipped[:10]); more=f"\n… 외 {len(skipped)-10}개" if len(skipped)>10 else ""
            QMessageBox.information(self,"제외된 파일",f"지원하지 않는 형식이라 {len(skipped)}개 파일을 목록에서 제외했습니다.\n이 파일들은 비식별 처리되지 않습니다.\n\n{names}{more}")
    def clear(self):
        self.files=[]; self.candidates=[]; self.warnings={}; self.failed_files=[]; self.file_list.clear(); self.table.setRowCount(0); self.file_count.setText("0개"); self.candidate_count.setText("후보 0건"); self.process_button.setEnabled(False); self.retry_button.setEnabled(False); self.status.setText("목록을 비웠습니다.")
    def _run(self,action):
        self.thread=QThread(self); self.worker=Worker(); self.worker.moveToThread(self.thread); self.worker.progress.connect(self.status.setText); self.worker.scanned.connect(self.on_scanned); self.worker.processed.connect(self.on_processed); self.worker.failed.connect(self.on_failed); self.thread.started.connect(action); self.thread.start()
    def _stop(self):
        if self.thread:
            self.thread.quit(); self.thread.wait(); self.thread.deleteLater(); self.thread=None
        if getattr(self,"worker",None): self.worker.deleteLater(); self.worker=None
    def on_failed(self,message):
        self._stop(); self.status.setText(message); self.scan_button.setEnabled(True); self.process_button.setEnabled(bool(self.candidates)); QMessageBox.critical(self,"처리 중단",message)
    def start_scan(self):
        if not self.files: QMessageBox.warning(self,"파일 없음","먼저 파일 또는 폴더를 추가하세요."); return
        custom=[x.strip() for x in self.custom.toPlainText().splitlines() if x.strip()]; self.scan_button.setEnabled(False); self.process_button.setEnabled(False); self.status.setText("전체 파일 분석 중…")
        self._run(lambda:self.worker.scan(list(self.files),custom))
    def on_scanned(self,candidates,warnings):
        self.candidates,self.warnings=candidates,warnings
        for c in self.candidates:
            if c.confidence < 0.75: c.selected=False
        self.show_candidates(); self._stop(); wc=sum(len(x) for x in warnings.values()); low=sum(c.confidence<0.75 for c in candidates); self.status.setText(f"후보 {len(candidates)}건 · 검토 필요 {low}건 · 경고 {wc}건"); self.scan_button.setEnabled(True); self.process_button.setEnabled(True)
        if not candidates and wc:
            details="\n".join(f"• {path.name}: {'; '.join(notes)}" for path,notes in warnings.items() if notes)
            has_ocr_error=any("OCR" in note for notes in warnings.values() for note in notes)
            advice="Tesseract 한국어 OCR 설치 상태와 문서 화질을 확인하세요." if has_ocr_error else "파일이 손상되었거나 지원하지 않는 내부 구조인지 오류 내용을 확인하세요."
            QMessageBox.warning(self,"인식 결과 없음",f"후보를 찾지 못했습니다.\n\n{details}\n\n{advice}")
    def show_candidates(self):
        self.table.setRowCount(len(self.candidates))
        for row in range(len(self.candidates)): self._fill_row(row)
        selected=sum(c.selected for c in self.candidates); self.candidate_count.setText(f"후보 {len(self.candidates)}건 · 적용 {selected}건")
    def _fill_row(self,row):
        c=self.candidates[row]
        state="완료" if c.reviewed else ("필요" if c.confidence<0.75 else "자동")
        values=["적용" if c.selected else "제외",state,c.kind,c.safe_value,c.file.name,c.location,f"{c.confidence:.0%}"]
        for col,value in enumerate(values):
            item=QTableWidgetItem(value)
            if col in {0,1,6}: item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if col==0:
                item.setForeground(QColor("#166534" if c.selected else "#9F1239")); item.setBackground(QColor("#ECFDF5" if c.selected else "#FFF1F2"))
            self.table.setItem(row,col,item)
    def toggle_selected(self,row,_=None):
        # 표 전체를 다시 만들면 검수 중이던 스크롤 위치가 맨 위로 튄다.
        self.candidates[row].selected=not self.candidates[row].selected
        if self.candidates[row].confidence<0.75:self.candidates[row].reviewed=True
        self._fill_row(row); self.candidate_count.setText(f"후보 {len(self.candidates)}건 · 적용 {sum(c.selected for c in self.candidates)}건")
    def table_double_clicked(self,row,column):
        if column==0:self.toggle_selected(row)
        else:self.review_selected()
    def _mode(self):
        if self.delete.isChecked(): return "delete"
        if self.label_ko.isChecked(): return "label_ko"
        if self.label_en.isChecked(): return "label_en"
        return "alias"
    def review_selected(self):
        row=self.table.currentRow()
        if row<0 or row>=len(self.candidates): QMessageBox.information(self,"항목 선택","먼저 검토할 후보 행을 선택하세요."); return
        candidate=self.candidates[row]; ext=candidate.file.suffix.lower()
        if ext in {".xlsx",".xls"}:
            replacement=replacement_for(candidate,self._mode(),{})
            ExcelPreviewDialog(candidate.file,candidate,replacement,self).exec(); candidate.reviewed=True; candidate.selected=True; self._fill_row(row); return
        if ext not in {".pdf",".png",".jpg",".jpeg"}: QMessageBox.information(self,"미리보기 제한","이 형식은 후보 목록에서 검토해 주세요."); return
        same=[c for c in self.candidates if c.file==candidate.file and c.location==candidate.location and c.selected]
        dialog=VisualMaskDialog(candidate.file,same or [candidate],self)
        if dialog.exec()!=QDialog.DialogCode.Accepted:return
        present=set(); new=[]
        for linked,bbox in dialog.areas():
            if linked is not None:
                linked.bbox=bbox; linked.manual=True; linked.reviewed=True; linked.selected=True; present.add(id(linked))
            else:
                new.append(Candidate("수동 지정","수동 영역",candidate.file,candidate.location,1.0,True,bbox,True,True))
        for linked in same:
            if id(linked) not in present: linked.selected=False; linked.reviewed=True
        self.candidates.extend(new); self.show_candidates()
    def approve_selected(self):
        rows=sorted({index.row() for index in self.table.selectedIndexes()})
        if not rows: QMessageBox.information(self,"항목 선택","검토 완료로 표시할 행을 선택하세요."); return
        for row in rows: self.candidates[row].reviewed=True; self.candidates[row].selected=True; self._fill_row(row)
        self.candidate_count.setText(f"후보 {len(self.candidates)}건 · 적용 {sum(c.selected for c in self.candidates)}건")
    def set_all(self,selected):
        for c in self.candidates:c.selected=selected
        self.show_candidates()
    def start_process(self):
        raw=QFileDialog.getExistingDirectory(self,"새 파일을 저장할 폴더 선택")
        if not raw:return
        target=Path(raw)
        if any(path.parent==target for path in self.files):
            answer=QMessageBox.question(self,"저장 위치 확인","선택한 폴더에 원본 파일이 함께 있습니다.\n결과와 원본이 섞이면 원본을 배포할 위험이 있습니다.\n\n그래도 이 폴더에 저장할까요?")
            if answer!=QMessageBox.StandardButton.Yes:return
        if not any(c.selected for c in self.candidates):
            answer=QMessageBox.question(self,"선택 항목 없음","적용으로 표시된 후보가 없습니다.\n개인정보가 지워지지 않은 복사본만 생성됩니다.\n\n계속할까요?")
            if answer!=QMessageBox.StandardButton.Yes:return
        unreviewed=[c for c in self.candidates if c.confidence<0.75 and not c.reviewed]
        if unreviewed:
            QMessageBox.warning(self,"저신뢰도 검토 필요",f"신뢰도 75% 미만 후보 {len(unreviewed)}건이 아직 검토되지 않았습니다.\n\n후보를 선택해 미리보기로 확인하거나 ‘검토 완료·적용’을 누른 뒤 저장하세요.")
            return
        if self.delete.isChecked():
            unsafe_ocr=[f"{path.name}: {note}" for path,notes in self.warnings.items() for note in notes if "OCR" in note]
            if unsafe_ocr:
                QMessageBox.warning(
                    self,"완전 삭제형 처리 중단",
                    "OCR 실패 또는 신뢰도 부족 경고가 있어 완전 삭제를 보장할 수 없습니다.\n\n"
                    + "\n".join(unsafe_ocr[:8])
                    + "\n\n원본 화질과 OCR 설치 상태를 개선한 뒤 다시 분석하세요."
                )
                return
        mode=self._mode(); self.scan_button.setEnabled(False); self.process_button.setEnabled(False); self.status.setText("마스킹 후 원문 잔존 여부까지 검증 중…")
        self._run(lambda:self.worker.process(list(self.files),self.candidates,self.warnings,mode,Path(raw)))
    def on_processed(self,results,report,output):
        self._stop(); ok=sum(r.status=="성공" for r in results); failed=len(results)-ok; count=sum(r.processed for r in results); low=sum(r.low_confidence for r in results)
        self.failed_files=[r.source for r in results if r.status!="성공"]; self.last_output=Path(output); self.retry_button.setEnabled(bool(self.failed_files))
        self.status.setText(f"완료: 성공 {ok}, 실패 {failed}, 처리 {count}건 · 저장 후 검증 완료"); self.scan_button.setEnabled(True); self.process_button.setEnabled(True)
        # 실패한 파일은 결과물이 없으므로 목록 뒤로 밀려 가려지면 안 된다.
        ordered=sorted(results,key=lambda r:r.status=="성공")
        details="\n".join(f"• {r.source.name}: {r.status} / {r.processed}건 / {r.message}" for r in ordered[:12])
        more=f"\n… 외 {len(results)-12}개" if len(results)>12 else ""
        text=f"성공 {ok}개 · 실패 {failed}개\n처리 {count}건 · 신뢰도 부족 {low}건\n\n{details}{more}\n\n결과 폴더: {output}\n보고서: {report.name}"
        (QMessageBox.warning if failed else QMessageBox.information)(self,"처리 결과",text)
    def retry_failed(self):
        if not self.failed_files or not self.last_output:return
        files=list(self.failed_files); self.retry_button.setEnabled(False); self.scan_button.setEnabled(False); self.process_button.setEnabled(False); self.status.setText(f"실패 파일 {len(files)}개 재실행 및 재검증 중…")
        self._run(lambda:self.worker.process(files,self.candidates,self.warnings,self._mode(),self.last_output))

if __name__=="__main__":
    app=QApplication(sys.argv); app.setStyle("Fusion"); window=PrivacyApp(); window.show(); sys.exit(app.exec())
