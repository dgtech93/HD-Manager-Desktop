import sys

from pathlib import Path

from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication

from app.controllers.app_controller import AppController
from app.logging_config import setup_logging


def main() -> int:
    setup_logging()
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
