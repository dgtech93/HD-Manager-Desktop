import sys

from pathlib import Path

from PyQt6.QtCore import QMessageLogContext, QtMsgType, qInstallMessageHandler
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication

from app.controllers.app_controller import AppController
from app.logging_config import setup_logging

import app.excel_io  # noqa: F401 — registra filtri warning openpyxl (Data Validation)


def _qt_message_handler(mode: QtMsgType, context: QMessageLogContext, message: str) -> None:
    """Evita rumore in console per messaggi Qt noti e innocui (font tema / stile)."""
    if "QFont::setPointSize" in message and "must be greater than 0" in message:
        return
    sys.stderr.write(message)
    if not message.endswith("\n"):
        sys.stderr.write("\n")


def main() -> int:
    setup_logging()
    # excel_io già importato sopra: filtri attivi prima di qualsiasi lettura Excel
    qInstallMessageHandler(_qt_message_handler)
    qt_app = QApplication(sys.argv)

    # Imposta l'icona dell'app (taskbar + eventuale titlebar).
    try:
        icon_path = (
            Path(__file__).resolve().parent / "app" / "assets" / "image.ico"
        )
        if icon_path.exists():
            qt_app.setWindowIcon(QIcon(str(icon_path)))
    except Exception:
        # In caso di problemi di path/formato, l'app deve comunque partire.
        pass

    controller = AppController()
    controller.start(qt_app)
    return qt_app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
