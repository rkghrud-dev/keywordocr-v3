from __future__ import annotations

import traceback
from PySide6 import QtCore

from app.services.pipeline import run_pipeline, PipelineConfig, run_listing_only, ListingOnlyConfig
from app.services.coupang import run_coupang_upload, CoupangUploadConfig, CoupangUploadResult


class PipelineWorker(QtCore.QObject):
    status = QtCore.Signal(str)
    progress = QtCore.Signal(int)
    finished = QtCore.Signal(str, str)
    error = QtCore.Signal(str)

    def __init__(self, config: PipelineConfig) -> None:
        super().__init__()
        self._config = config

    @QtCore.Slot()
    def run(self) -> None:
        try:
            self.status.emit("⚙️ PipelineWorker.run() 시작!")
            out_root, out_file = run_pipeline(
                self._config,
                status_cb=self.status.emit,
                progress_cb=self.progress.emit,
            )
            self.status.emit("✅ run_pipeline 완료!")
            self.finished.emit(out_root, out_file)
        except Exception:
            self.error.emit(traceback.format_exc())


class ListingWorker(QtCore.QObject):
    """대표이미지만 생성하는 워커."""
    status = QtCore.Signal(str)
    progress = QtCore.Signal(int)
    finished_listing = QtCore.Signal(str)   # listing_out_root
    error = QtCore.Signal(str)

    def __init__(self, config: ListingOnlyConfig) -> None:
        super().__init__()
        self._config = config

    @QtCore.Slot()
    def run(self) -> None:
        try:
            out_root = run_listing_only(
                self._config,
                status_cb=self.status.emit,
                progress_cb=self.progress.emit,
            )
            self.finished_listing.emit(out_root)
        except Exception:
            self.error.emit(traceback.format_exc())


class CoupangUploadWorker(QtCore.QObject):
    """쿠팡 업로드 워커 (QThread에서 실행)."""
    status = QtCore.Signal(str)
    progress = QtCore.Signal(int)
    finished = QtCore.Signal(list)   # list[CoupangUploadResult]
    error = QtCore.Signal(str)

    def __init__(self, config: CoupangUploadConfig) -> None:
        super().__init__()
        self._config = config

    @QtCore.Slot()
    def run(self) -> None:
        try:
            results = run_coupang_upload(
                self._config,
                status_cb=self.status.emit,
                progress_cb=self.progress.emit,
            )
            self.finished.emit(results)
        except Exception:
            self.error.emit(traceback.format_exc())
