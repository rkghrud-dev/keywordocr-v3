"""OCR 전용 프로그램 — PySide6 GUI."""
from __future__ import annotations

import os
import subprocess
import traceback

from PySide6 import QtCore, QtGui, QtWidgets

from app.gui.widgets import DropLineEdit
from app.services.ocr_pipeline import run_ocr_pipeline, OcrPipelineConfig


# ── 워커 ──────────────────────────────────────────────────────────────

class OcrWorker(QtCore.QObject):
    status = QtCore.Signal(str)
    progress = QtCore.Signal(int)
    finished = QtCore.Signal(str)   # output Excel path
    error = QtCore.Signal(str)

    def __init__(self, config: OcrPipelineConfig) -> None:
        super().__init__()
        self._config = config

    @QtCore.Slot()
    def run(self) -> None:
        try:
            out_path = run_ocr_pipeline(
                self._config,
                status_cb=self.status.emit,
                progress_cb=self.progress.emit,
            )
            self.finished.emit(out_path)
        except Exception:
            self.error.emit(traceback.format_exc())


# ── 메인 윈도우 ──────────────────────────────────────────────────────

class OcrWindow(QtWidgets.QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("OCR 전용 도구")
        self.resize(780, 650)
        self._thread: QtCore.QThread | None = None
        self._last_output: str = ""
        self._init_ui()

    # ── UI 구성 ───────────────────────────────────────────────────

    def _init_ui(self) -> None:
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        root = QtWidgets.QVBoxLayout(central)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # ── 입력 그룹 ────
        input_box = QtWidgets.QGroupBox("입력")
        ig = QtWidgets.QGridLayout(input_box)

        ig.addWidget(QtWidgets.QLabel("CSV 파일"), 0, 0)
        self.csv_edit = DropLineEdit(mode="file")
        self.csv_edit.setPlaceholderText("CSV / Excel 파일을 선택하거나 드래그 앤 드롭")
        ig.addWidget(self.csv_edit, 0, 1)
        csv_btn = QtWidgets.QPushButton("파일 선택")
        csv_btn.clicked.connect(self._browse_csv)
        ig.addWidget(csv_btn, 0, 2)

        ig.addWidget(QtWidgets.QLabel("이미지 폴더"), 1, 0)
        self.img_dir_edit = DropLineEdit(mode="dir")
        self.img_dir_edit.setPlaceholderText("로컬 이미지 루트 폴더")
        default_dir = r"D:\pp"
        if os.path.isdir(default_dir):
            self.img_dir_edit.setText(default_dir)
        ig.addWidget(self.img_dir_edit, 1, 1)
        dir_btn = QtWidgets.QPushButton("폴더 선택")
        dir_btn.clicked.connect(self._browse_dir)
        ig.addWidget(dir_btn, 1, 2)

        ig.addWidget(QtWidgets.QLabel("Tesseract 경로"), 2, 0)
        self.tess_edit = QtWidgets.QLineEdit()
        self.tess_edit.setPlaceholderText("자동 감지 (비워두면 기본 경로 검색)")
        ig.addWidget(self.tess_edit, 2, 1, 1, 2)

        root.addWidget(input_box)

        # ── OCR 엔진 선택 그룹 ────
        engine_box = QtWidgets.QGroupBox("OCR 엔진")
        eg = QtWidgets.QGridLayout(engine_box)

        self.use_google_vision = QtWidgets.QCheckBox("Google Cloud Vision API 사용 (고정확도)")
        self.use_google_vision.setChecked(False)
        self.use_google_vision.toggled.connect(self._on_engine_toggle)
        eg.addWidget(self.use_google_vision, 0, 0, 1, 4)

        eg.addWidget(QtWidgets.QLabel("서비스 계정 JSON"), 1, 0)
        self.gcp_cred_edit = DropLineEdit(mode="file")
        self.gcp_cred_edit.setPlaceholderText("Google Cloud 서비스 계정 키 파일 (.json)")
        # 기본 경로 자동 탐색
        default_cred = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(
            os.path.abspath(__file__)))), "google_vision_key.json")
        if os.path.isfile(default_cred):
            self.gcp_cred_edit.setText(default_cred)
        eg.addWidget(self.gcp_cred_edit, 1, 1, 1, 2)
        gcp_btn = QtWidgets.QPushButton("파일 선택")
        gcp_btn.clicked.connect(self._browse_gcp_cred)
        eg.addWidget(gcp_btn, 1, 3)

        self.gcp_cred_edit.setEnabled(False)

        root.addWidget(engine_box)

        # ── OCR 설정 그룹 ────
        ocr_box = QtWidgets.QGroupBox("OCR 설정")
        og = QtWidgets.QGridLayout(ocr_box)

        self.korean_only = QtWidgets.QCheckBox("한글만 유지")
        self.korean_only.setChecked(True)
        og.addWidget(self.korean_only, 0, 0, 1, 2)

        self.skip_last_image = QtWidgets.QCheckBox("마지막 이미지 제외 (주의사항/배송안내)")
        self.skip_last_image.setChecked(True)
        og.addWidget(self.skip_last_image, 0, 2, 1, 2)

        og.addWidget(QtWidgets.QLabel("PSM"), 1, 0)
        self.psm = QtWidgets.QSpinBox()
        self.psm.setRange(0, 13)
        self.psm.setValue(11)
        og.addWidget(self.psm, 1, 1)

        og.addWidget(QtWidgets.QLabel("OEM"), 1, 2)
        self.oem = QtWidgets.QSpinBox()
        self.oem.setRange(0, 3)
        self.oem.setValue(3)
        og.addWidget(self.oem, 1, 3)

        og.addWidget(QtWidgets.QLabel("스레드 수"), 2, 0)
        self.threads = QtWidgets.QSpinBox()
        self.threads.setRange(1, 16)
        self.threads.setValue(6)
        og.addWidget(self.threads, 2, 1)

        og.addWidget(QtWidgets.QLabel("폴더 탐색 깊이"), 2, 2)
        self.max_depth = QtWidgets.QSpinBox()
        self.max_depth.setRange(-1, 20)
        self.max_depth.setValue(-1)
        self.max_depth.setSpecialValueText("무제한")
        og.addWidget(self.max_depth, 2, 3)

        self.allow_folder_match = QtWidgets.QCheckBox("폴더명 GS코드 매칭 허용")
        self.allow_folder_match.setChecked(True)
        og.addWidget(self.allow_folder_match, 3, 0, 1, 2)

        self.filter_noise = QtWidgets.QCheckBox("반복 문구 자동 필터링 (배송안내, 주의사항 등)")
        self.filter_noise.setChecked(True)
        og.addWidget(self.filter_noise, 3, 2, 1, 2)

        root.addWidget(ocr_box)

        # ── 출력 그룹 ────
        out_box = QtWidgets.QGroupBox("출력")
        out_lay = QtWidgets.QHBoxLayout(out_box)
        out_lay.addWidget(QtWidgets.QLabel("저장 폴더"))
        self.out_dir_edit = DropLineEdit(mode="dir")
        self.out_dir_edit.setPlaceholderText("비워두면 CSV와 같은 폴더에 저장")
        out_lay.addWidget(self.out_dir_edit, 1)
        out_btn = QtWidgets.QPushButton("폴더 선택")
        out_btn.clicked.connect(self._browse_out_dir)
        out_lay.addWidget(out_btn)
        root.addWidget(out_box)

        # ── 실행 버튼 ────
        self.run_btn = QtWidgets.QPushButton("  OCR 실행  ")
        self.run_btn.setStyleSheet(
            "QPushButton{background:#4caf50;color:#fff;font-size:14px;"
            "font-weight:bold;padding:10px 30px;border-radius:5px}"
            "QPushButton:hover{background:#388e3c}"
            "QPushButton:disabled{background:#999}"
        )
        self.run_btn.clicked.connect(self._on_run)
        root.addWidget(self.run_btn, alignment=QtCore.Qt.AlignCenter)

        # ── 진행상황 ────
        prog_lay = QtWidgets.QHBoxLayout()
        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setRange(0, 100)
        prog_lay.addWidget(self.progress_bar, 1)
        self.open_folder_btn = QtWidgets.QPushButton("결과 폴더 열기")
        self.open_folder_btn.setEnabled(False)
        self.open_folder_btn.clicked.connect(self._open_result_folder)
        prog_lay.addWidget(self.open_folder_btn)
        root.addLayout(prog_lay)

        self.status_label = QtWidgets.QLabel("대기 중")
        self.status_label.setStyleSheet("color:#1565c0; font-weight:bold;")
        root.addWidget(self.status_label)

        # ── 로그 ────
        self.log_area = QtWidgets.QPlainTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMaximumBlockCount(2000)
        root.addWidget(self.log_area, 1)

    # ── 파일/폴더 탐색 ────────────────────────────────────────────

    def _browse_csv(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "CSV/Excel 파일 선택", "",
            "데이터 파일 (*.csv *.xls *.xlsx);;모든 파일 (*)")
        if path:
            self.csv_edit.setText(path)

    def _browse_dir(self) -> None:
        d = QtWidgets.QFileDialog.getExistingDirectory(self, "이미지 폴더 선택")
        if d:
            self.img_dir_edit.setText(d)

    def _browse_out_dir(self) -> None:
        d = QtWidgets.QFileDialog.getExistingDirectory(self, "출력 폴더 선택")
        if d:
            self.out_dir_edit.setText(d)

    def _browse_gcp_cred(self) -> None:
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Google Cloud 서비스 계정 키 선택", "", "JSON (*.json)")
        if path:
            self.gcp_cred_edit.setText(path)

    def _on_engine_toggle(self, checked: bool) -> None:
        self.gcp_cred_edit.setEnabled(checked)
        self.psm.setEnabled(not checked)
        self.oem.setEnabled(not checked)
        self.tess_edit.setEnabled(not checked)

    # ── 실행 ──────────────────────────────────────────────────────

    def _on_run(self) -> None:
        csv_path = self.csv_edit.text().strip()
        img_dir = self.img_dir_edit.text().strip()

        if not csv_path or not os.path.isfile(csv_path):
            QtWidgets.QMessageBox.warning(self, "입력 오류", "CSV 파일을 선택해 주세요.")
            return
        if not img_dir or not os.path.isdir(img_dir):
            QtWidgets.QMessageBox.warning(self, "입력 오류", "이미지 폴더를 선택해 주세요.")
            return

        config = OcrPipelineConfig(
            csv_path=csv_path,
            local_img_dir=img_dir,
            output_dir=self.out_dir_edit.text().strip(),
            tesseract_path=self.tess_edit.text().strip(),
            korean_only=self.korean_only.isChecked(),
            psm=self.psm.value(),
            oem=self.oem.value(),
            max_depth=self.max_depth.value(),
            allow_folder_match=self.allow_folder_match.isChecked(),
            threads=self.threads.value(),
            skip_last_image=self.skip_last_image.isChecked(),
            use_google_vision=self.use_google_vision.isChecked(),
            google_credentials_path=self.gcp_cred_edit.text().strip(),
            filter_noise=self.filter_noise.isChecked(),
        )

        self.run_btn.setEnabled(False)
        self.open_folder_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.log_area.clear()

        self._thread = QtCore.QThread()
        self._worker = OcrWorker(config)
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.status.connect(self._on_status)
        self._worker.progress.connect(self.progress_bar.setValue)
        self._worker.finished.connect(self._on_finished)
        self._worker.error.connect(self._on_error)
        self._worker.finished.connect(self._thread.quit)
        self._worker.error.connect(self._thread.quit)

        self._thread.start()

    def _on_status(self, msg: str) -> None:
        self.status_label.setText(msg)
        self.log_area.appendPlainText(msg)

    def _on_finished(self, out_path: str) -> None:
        self.run_btn.setEnabled(True)
        self._last_output = out_path
        self.open_folder_btn.setEnabled(True)
        QtWidgets.QMessageBox.information(
            self, "완료",
            f"OCR 결과가 저장되었습니다:\n{out_path}")

    def _on_error(self, tb: str) -> None:
        self.run_btn.setEnabled(True)
        self.log_area.appendPlainText(f"\n=== 오류 ===\n{tb}")
        self.status_label.setText("오류 발생")
        QtWidgets.QMessageBox.critical(self, "오류", f"처리 중 오류 발생:\n{tb[:500]}")

    def _open_result_folder(self) -> None:
        if self._last_output and os.path.isfile(self._last_output):
            folder = os.path.dirname(self._last_output)
            if os.name == "nt":
                os.startfile(folder)
            else:
                subprocess.Popen(["xdg-open", folder])
