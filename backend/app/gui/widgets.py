from __future__ import annotations

import os
from PySide6 import QtCore, QtGui, QtWidgets


class DropLineEdit(QtWidgets.QLineEdit):
    path_dropped = QtCore.Signal(str)

    def __init__(self, mode: str = "file", parent: QtWidgets.QWidget | None = None) -> None:
        super().__init__(parent)
        self._mode = mode  # "file" or "dir"
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QtGui.QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QtGui.QDropEvent) -> None:
        urls = event.mimeData().urls()
        if not urls:
            return
        for url in urls:
            path = url.toLocalFile()
            if not path:
                continue
            if self._mode == "dir" and os.path.isdir(path):
                self.setText(path)
                self.path_dropped.emit(path)
                return
            if self._mode == "file" and os.path.isfile(path):
                self.setText(path)
                self.path_dropped.emit(path)
                return
