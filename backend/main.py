import os
import sys
import traceback
from pathlib import Path

from PySide6.QtWidgets import QApplication, QMessageBox


PROJECT_ROOT = Path(__file__).resolve().parent


def _prepare_runtime() -> None:
    # Ensure relative paths always resolve from the project root.
    os.chdir(PROJECT_ROOT)
    project_root_str = str(PROJECT_ROOT)
    if project_root_str not in sys.path:
        sys.path.insert(0, project_root_str)


def _show_startup_error(exc: BaseException) -> None:
    error_text = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    log_path = PROJECT_ROOT / "startup_error.log"
    try:
        log_path.write_text(error_text, encoding="utf-8")
    except Exception:
        pass

    app = QApplication.instance() or QApplication(sys.argv)
    QMessageBox.critical(
        None,
        "앱 시작 오류",
        f"프로그램 시작 중 오류가 발생했습니다.\\n\\n로그: {log_path}\\n\\n{exc}",
    )


def main() -> None:
    _prepare_runtime()

    from app.gui.main_window import MainWindow

    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        _show_startup_error(exc)
        sys.exit(1)
