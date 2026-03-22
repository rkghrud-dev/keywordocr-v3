# -*- coding: utf-8 -*-
"""이미지 선택 대화상자 - 업로드 전 대표/추가 이미지 선택"""

import os
from PySide6 import QtWidgets, QtCore, QtGui


class ImageSelectorDialog(QtWidgets.QDialog):
    """이미지 선택 대화상자

    여러 상품의 이미지를 순차적으로 보여주고
    각 상품별로 대표 이미지와 추가 이미지를 선택할 수 있음
    """

    def __init__(self, products_data: list[dict], parent=None, initial_selections: dict | None = None):
        """
        Args:
            products_data: [{
                "gs_code": "GS2100186",
                "image_folder": "path/to/GS2100186",
                "images": ["1.jpg", "2.jpg", ...]
            }, ...]
        """
        super().__init__(parent)
        self.products_data = products_data
        self.current_index = 0
        self.selections = dict(initial_selections or {})  # {gs_code: {"main": idx, "additional": [idx1, idx2, ...]}}
        self.max_additional = 10

        self.setWindowTitle("이미지 선택")
        self.setMinimumSize(900, 700)
        self.setModal(True)

        self._build_ui()
        self._build_shortcuts()
        self._load_current_product()

    def _build_ui(self):
        layout = QtWidgets.QVBoxLayout(self)

        # 상단: 진행 상황
        self.progress_label = QtWidgets.QLabel()
        self.progress_label.setStyleSheet("font-size:14px; font-weight:600;")
        layout.addWidget(self.progress_label)

        # 안내 문구
        info_label = QtWidgets.QLabel(
            "● = 대표 이미지 (1개만 선택)   |   ☑ = 추가 이미지 (최대 10개)   |   단축키: 1~9,0=대표선택 / Space=다음 / Backspace=이전"
        )
        info_label.setStyleSheet("color:#555; margin:5px;")
        layout.addWidget(info_label)

        # 스크롤 영역
        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)

        self.image_container = QtWidgets.QWidget()
        self.image_layout = QtWidgets.QGridLayout(self.image_container)
        self.image_layout.setSpacing(10)

        scroll.setWidget(self.image_container)
        layout.addWidget(scroll, 1)

        # 하단: 버튼
        btn_row = QtWidgets.QHBoxLayout()

        self.prev_btn = QtWidgets.QPushButton("◀ 이전")
        self.prev_btn.clicked.connect(self._go_prev)
        btn_row.addWidget(self.prev_btn)

        self.next_btn = QtWidgets.QPushButton("다음 상품 ▶")
        self.next_btn.clicked.connect(self._go_next)
        btn_row.addWidget(self.next_btn)

        btn_row.addStretch()

        cancel_btn = QtWidgets.QPushButton("취소")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)

        self.done_btn = QtWidgets.QPushButton("완료")
        self.done_btn.setStyleSheet("background:#2e7d32; color:#fff; padding:8px; font-weight:600;")
        self.done_btn.clicked.connect(self.accept)
        btn_row.addWidget(self.done_btn)

        layout.addLayout(btn_row)

    
    def _build_shortcuts(self):
        """숫자/이동 단축키: 1~9,0은 대표 선택, Space 다음, Backspace 이전."""
        self._digit_shortcuts = []
        for key_text in ["1", "2", "3", "4", "5", "6", "7", "8", "9", "0"]:
            sc = QtGui.QShortcut(QtGui.QKeySequence(key_text), self)
            sc.setContext(QtCore.Qt.WidgetWithChildrenShortcut)
            sc.activated.connect(lambda k=key_text: self._on_digit_shortcut(k))
            self._digit_shortcuts.append(sc)

        self._next_shortcut = QtGui.QShortcut(QtGui.QKeySequence("Space"), self)
        self._next_shortcut.setContext(QtCore.Qt.WidgetWithChildrenShortcut)
        self._next_shortcut.activated.connect(self._go_next)

        self._prev_shortcut = QtGui.QShortcut(QtGui.QKeySequence("Backspace"), self)
        self._prev_shortcut.setContext(QtCore.Qt.WidgetWithChildrenShortcut)
        self._prev_shortcut.activated.connect(self._go_prev)

    def _on_digit_shortcut(self, key_text: str):
        if not hasattr(self, "image_widgets"):
            return
        # 1~9 -> 1~9번째, 0 -> 10번째 이미지를 대표로 선택
        idx = 9 if key_text == "0" else (int(key_text) - 1)
        if idx < 0 or idx >= len(self.image_widgets):
            return
        self.image_widgets[idx].main_radio.setChecked(True)

    def _load_current_product(self):
        """현재 상품의 이미지 로드"""
        if self.current_index >= len(self.products_data):
            return

        product = self.products_data[self.current_index]
        gs_code = product["gs_code"]
        image_folder = product["image_folder"]
        images = product["images"]

        # 진행 상황 업데이트
        self.progress_label.setText(
            f"{gs_code} 이미지 선택 ({self.current_index + 1}/{len(self.products_data)})"
        )

        # 이전 위젯 제거
        for i in reversed(range(self.image_layout.count())):
            widget = self.image_layout.itemAt(i).widget()
            if widget:
                widget.deleteLater()

        # 이전 선택 복원
        prev_selection = self.selections.get(gs_code, {"main": None, "additional": []})

        # 이미지 위젯 생성
        self.image_widgets = []
        cols = 4
        for idx, img_name in enumerate(images):
            row = idx // cols
            col = idx % cols

            widget = self._create_image_widget(
                image_folder, img_name, idx,
                is_main=(idx == prev_selection["main"]),
                is_additional=(idx in prev_selection["additional"])
            )
            self.image_widgets.append(widget)
            self.image_layout.addWidget(widget, row, col)

        # 버튼 상태 업데이트
        self.prev_btn.setEnabled(self.current_index > 0)
        self.next_btn.setEnabled(self.current_index < len(self.products_data) - 1)

        # 추가 이미지 개수 확인
        self._check_additional_count()

    def _create_image_widget(self, folder: str, filename: str, idx: int,
                            is_main: bool, is_additional: bool) -> QtWidgets.QWidget:
        """개별 이미지 위젯 생성"""
        widget = QtWidgets.QWidget()
        widget.setFixedSize(200, 240)
        layout = QtWidgets.QVBoxLayout(widget)
        layout.setContentsMargins(5, 5, 5, 5)

        # 이미지 표시
        image_path = os.path.join(folder, filename)
        pixmap = QtGui.QPixmap(image_path)
        if not pixmap.isNull():
            pixmap = pixmap.scaled(180, 180, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)

        img_label = QtWidgets.QLabel()
        img_label.setPixmap(pixmap)
        img_label.setAlignment(QtCore.Qt.AlignCenter)
        img_label.setStyleSheet("border:1px solid #ccc; background:#f5f5f5;")
        img_label.setFixedSize(180, 180)
        layout.addWidget(img_label)

        # 파일명
        name_label = QtWidgets.QLabel(filename)
        name_label.setAlignment(QtCore.Qt.AlignCenter)
        name_label.setStyleSheet("font-size:11px;")
        layout.addWidget(name_label)

        # 선택 버튼
        btn_row = QtWidgets.QHBoxLayout()

        main_radio = QtWidgets.QRadioButton("대표")
        main_radio.setChecked(is_main)
        main_radio.toggled.connect(lambda checked: self._on_main_selected(idx, checked))
        btn_row.addWidget(main_radio)

        add_check = QtWidgets.QCheckBox("추가")
        add_check.setChecked(is_additional)
        add_check.toggled.connect(lambda checked: self._on_additional_toggled(idx, checked))
        btn_row.addWidget(add_check)

        layout.addLayout(btn_row)

        # 위젯에 참조 저장
        widget.main_radio = main_radio
        widget.add_check = add_check
        widget.img_idx = idx

        return widget

    def _on_main_selected(self, idx: int, checked: bool):
        """대표 이미지 선택"""
        if checked:
            # 다른 라디오 버튼 해제
            for w in self.image_widgets:
                if w.img_idx != idx:
                    w.main_radio.setChecked(False)

            # 대표 이미지를 제외한 나머지를 자동으로 추가 이미지로 체크
            for w in self.image_widgets:
                if w.img_idx != idx:
                    w.add_check.setChecked(True)

            # 추가 이미지 개수 확인
            self._check_additional_count()

    def _on_additional_toggled(self, idx: int, checked: bool):
        """추가 이미지 선택"""
        self._check_additional_count()

    def _check_additional_count(self):
        """추가 이미지 개수 확인"""
        checked_count = sum(1 for w in self.image_widgets if w.add_check.isChecked())
        over_count = checked_count - self.max_additional

        if over_count > 0:
            # 초과 - 버튼 비활성화 + 툴팁 설정
            self.next_btn.setEnabled(False)
            self.done_btn.setEnabled(False)
            tooltip_msg = f"추가 이미지 {over_count}개를 체크 해제해주세요 (현재: {checked_count}개)"
            self.next_btn.setToolTip(tooltip_msg)
            self.done_btn.setToolTip(tooltip_msg)
        else:
            # 정상 범위 - 버튼 활성화 + 툴팁 제거
            self.next_btn.setEnabled(self.current_index < len(self.products_data) - 1)
            self.done_btn.setEnabled(True)
            self.next_btn.setToolTip("")
            self.done_btn.setToolTip("")

    def _save_current_selection(self):
        """현재 선택 저장"""
        if self.current_index >= len(self.products_data):
            return

        product = self.products_data[self.current_index]
        gs_code = product["gs_code"]

        main_idx = None
        additional_indices = []

        for w in self.image_widgets:
            if w.main_radio.isChecked():
                main_idx = w.img_idx
            if w.add_check.isChecked():
                additional_indices.append(w.img_idx)

        self.selections[gs_code] = {
            "main": main_idx,
            "additional": sorted(additional_indices)
        }

    def _go_prev(self):
        """이전 상품으로"""
        self._save_current_selection()
        if self.current_index > 0:
            self.current_index -= 1
            self._load_current_product()

    def _go_next(self):
        """다음 상품으로"""
        self._save_current_selection()
        if self.current_index < len(self.products_data) - 1:
            self.current_index += 1
            self._load_current_product()

    def accept(self):
        """완료 - 모든 선택 저장"""
        self._save_current_selection()

        # 검증: 모든 상품에 대표 이미지가 선택되었는지
        missing = []
        for product in self.products_data:
            gs_code = product["gs_code"]
            if gs_code not in self.selections or self.selections[gs_code]["main"] is None:
                missing.append(gs_code)

        if missing:
            QtWidgets.QMessageBox.warning(
                self, "선택 미완료",
                f"다음 상품의 대표 이미지를 선택해주세요:\n{', '.join(missing[:5])}" +
                (f"\n... 외 {len(missing)-5}개" if len(missing) > 5 else "")
            )
            return

        super().accept()

    def get_selections(self) -> dict:
        """선택 결과 반환

        Returns:
            {gs_code: {"main": idx, "additional": [idx1, idx2, ...]}}
        """
        return self.selections


