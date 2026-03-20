from __future__ import annotations

from PyQt6.QtWidgets import QApplication

from app.services.db_service import DbService
from app.services.repository_service import RepositoryService
from app.services.system_service import SystemService
from app.services.clients_service import ClientsService
from app.services.archive_service import ArchiveService
from app.services.credentials_service import CredentialsService
from app.controllers.main_controller import MainController
from app.views.main_window import MainWindow


class AppController:
    """Composition root: wires services, controllers and views."""

    def __init__(self) -> None:
        self.db = DbService()
        self.repository = RepositoryService()
        self.system = SystemService()

        self.main_controller = MainController(
            clients=ClientsService(self.repository),
            archive=ArchiveService(self.repository),
            credentials=CredentialsService(self.repository),
            system=self.system,
        )

        self._main_window: MainWindow | None = None

    def start(self, qt_app: QApplication) -> None:
        self.db.init_db()
        self._main_window = MainWindow(controller=self.main_controller)
        self._main_window.show()

