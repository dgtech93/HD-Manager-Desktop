from __future__ import annotations

from pathlib import Path

from PyQt6.QtCore import QEvent, Qt
from PyQt6.QtGui import QColor, QIcon, QPainter, QPainterPath, QPen, QPixmap
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLayout,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from app.ui.settings_window import SettingsWindow
from app.ui.ui_constants import SIDE_MENU_TAB_CONTENT_MARGINS, SIDE_MENU_WIDTH_PX
from app.ui.clients_mixin import ClientsMixin
from app.ui.archive_mixin import ArchiveMixin
from app.ui.calculator_widget import CalculatorWorkspaceWidget
from app.ui.agenda_widget import AgendaUpcomingHeaderWidget, AgendaWorkspaceWidget
from app.ui.sql_archive_widget import SqlArchiveWorkspaceWidget
from app.ui.excel_management_widget import ExcelManagementWidget
from app.ui.sticky_notes_widget import NotesWorkspaceWidget, StickyNotesHeaderWidget
from app.version import __version__


def _pixmap_circular(source: QPixmap, diameter: int) -> QPixmap:
    """Ridimensiona l'immagine e la ritaglia in un cerchio (antialias)."""
    if source.isNull():
        return source
    side = diameter
    scaled = source.scaled(
        side,
        side,
        Qt.AspectRatioMode.KeepAspectRatioByExpanding,
        Qt.TransformationMode.SmoothTransformation,
    )
    sx = max(0, (scaled.width() - side) // 2)
    sy = max(0, (scaled.height() - side) // 2)
    out = QPixmap(side, side)
    out.fill(Qt.GlobalColor.transparent)
    painter = QPainter(out)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    path = QPainterPath()
    path.addEllipse(0, 0, side, side)
    painter.setClipPath(path)
    painter.drawPixmap(0, 0, side, side, scaled, sx, sy, side, side)
    painter.setClipping(False)
    pen = QPen(QColor(255, 255, 255, 230))
    pen.setWidth(2)
    painter.setPen(pen)
    painter.setBrush(Qt.BrushStyle.NoBrush)
    painter.drawEllipse(1, 1, side - 2, side - 2)
    painter.end()
    return out


class MainWindow(ClientsMixin, ArchiveMixin, QMainWindow):
    def __init__(self, repository) -> None:
        super().__init__()
        self.repository = repository
        self.settings_window: SettingsWindow | None = None
        self._agenda_window: QMainWindow | None = None
        self._notes_window: QMainWindow | None = None

        # Icona principale della finestra (taskbar + titlebar).
        # Se non trova il file, continua senza bloccare l'app.
        try:
            icon_path = Path(__file__).resolve().parent.parent / "assets" / "image.ico"
            if icon_path.exists():
                self.setWindowIcon(QIcon(str(icon_path)))
        except Exception:
            pass

        self.clients_cache: list[dict] = []
        self.products_cache: list[dict] = []
        self.resources_cache: list[dict] = []
        self.vpns_cache: list[dict] = []
        self.roles_cache: list[dict] = []
        self.role_multi_map: dict[str, bool] = {}
        self.role_order_map: dict[str, int] = {}
        self.clients_by_id: dict[int, dict] = {}
        self.selected_client_id: int | None = None
        self.selected_product_id: int | None = None

        self.setWindowTitle(f"HD Manager Desktop  v{__version__}")
        self.resize(1380, 880)
        self._build_ui()
        self._apply_style()
        self.refresh_views()

    # UI
    def _build_ui(self) -> None:
        self.setMenuBar(None)

        root = QWidget()
        self.setCentralWidget(root)

        layout = QVBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        header_frame = QFrame()
        header_frame.setObjectName("headerFrame")
        # Margine layout + riga (~142px card + eventuale scrollbar orizzontale)
        header_frame.setMinimumHeight(200)
        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(24, 22, 24, 20)
        header_layout.setSpacing(8)

        top_row = QHBoxLayout()
        top_row.setSpacing(16)

        # Contenitore centrale: scroll orizzontale se la finestra è stretta (evita tagli).
        self._header_middle_inner = QWidget()
        self._header_middle_inner.setObjectName("headerMiddleArea")
        self._header_middle_inner.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        mid_l = QHBoxLayout(self._header_middle_inner)
        mid_l.setContentsMargins(0, 0, 0, 0)
        mid_l.setSpacing(12)
        self.sticky_notes_header = StickyNotesHeaderWidget(self)
        self.sticky_notes_header.setObjectName("stickyNotesHeader")
        mid_l.addWidget(self.sticky_notes_header, 0, Qt.AlignmentFlag.AlignTop)
        self.agenda_upcoming_header = AgendaUpcomingHeaderWidget(self)
        self.agenda_upcoming_header.setObjectName("agendaUpcomingHeader")
        mid_l.addWidget(self.agenda_upcoming_header, 0, Qt.AlignmentFlag.AlignTop)
        mid_l.addStretch(1)

        self._header_middle = QScrollArea()
        self._header_middle.setObjectName("headerMiddleScroll")
        self._header_middle.setWidget(self._header_middle_inner)
        # True comprime il widget al viewport e taglia verticalmente le card (142px fisse).
        self._header_middle.setWidgetResizable(False)
        self._header_middle.setAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop
        )
        self._header_middle.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self._header_middle.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._header_middle.setFrameShape(QFrame.Shape.NoFrame)
        self._header_middle.setSizePolicy(
            QSizePolicy.Policy.Expanding,
            QSizePolicy.Policy.Minimum,
        )
        self._header_middle.setMinimumHeight(142)
        self._header_middle.viewport().setAutoFillBackground(False)
        top_row.addWidget(self._header_middle, 1, Qt.AlignmentFlag.AlignTop)
        self._apply_sticky_notes_visibility()
        self._apply_agenda_header_visibility()
        self._sync_header_middle_min_width()

        self.header_logo_label = QLabel()
        self.header_logo_label.setObjectName("headerLogoLabel")
        self.header_logo_label.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.header_logo_label.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        icon_path = Path(__file__).resolve().parent.parent / "assets" / "image.ico"
        if icon_path.exists():
            pm = QPixmap(str(icon_path))
            if not pm.isNull():
                self.header_logo_label.setPixmap(_pixmap_circular(pm, 88))

        self.agenda_header_btn = QPushButton("Agenda")
        self.agenda_header_btn.setObjectName("headerAgendaButton")
        self.agenda_header_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.agenda_header_btn.setToolTip(
            "Apri l’agenda in una finestra massimizzata (barra del titolo e pulsanti Chiudi / Riduci)"
        )
        self.agenda_header_btn.clicked.connect(self.open_agenda_window)

        self.notes_header_btn = QPushButton("Note")
        self.notes_header_btn.setObjectName("headerNotesButton")
        self.notes_header_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.notes_header_btn.setToolTip("Apri la gestione note in una finestra separata")
        self.notes_header_btn.clicked.connect(self.open_notes_window)

        self.settings_header_btn = QPushButton("Impostazioni…")
        self.settings_header_btn.setObjectName("headerSettingsButton")
        self.settings_header_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.settings_header_btn.setToolTip("Apri la finestra delle impostazioni dell’applicazione")
        self.settings_header_btn.clicked.connect(self.open_settings)

        header_right = QWidget()
        header_right.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        header_right_l = QVBoxLayout(header_right)
        header_right_l.setContentsMargins(0, 0, 0, 0)
        header_right_l.setSpacing(10)
        header_btn_row = QHBoxLayout()
        header_btn_row.setContentsMargins(0, 0, 0, 0)
        header_btn_row.setSpacing(8)
        header_btn_row.addWidget(self.agenda_header_btn, 0)
        header_btn_row.addWidget(self.notes_header_btn, 0)
        header_btn_row.addWidget(self.settings_header_btn, 0)
        header_btn_row.addStretch(1)
        header_right_l.addLayout(header_btn_row)
        header_right_l.addWidget(self.header_logo_label, 0, Qt.AlignmentFlag.AlignRight)
        top_row.addWidget(header_right, 0, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)

        header_layout.addLayout(top_row)
        layout.addWidget(header_frame)

        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(18, 16, 18, 18)
        content_layout.setSpacing(14)

        self.main_tabs = QTabWidget()
        self.main_tabs.setObjectName("mainTabs")
        self.main_tabs.addTab(self._build_clients_tab(), "Clienti")
        self.main_tabs.addTab(self._build_archive_tab(), "Archivio")
        self.main_tabs.addTab(self._build_tools_tab(), "Strumenti")
        self.main_tabs.addTab(ExcelManagementWidget(main_window=self), "Strumenti Excel")
        self._agenda_workspace_widget = AgendaWorkspaceWidget(main_window=self)
        self._notes_workspace_widget = NotesWorkspaceWidget(main_window=self)
        self._sql_archive_widget = SqlArchiveWorkspaceWidget(main_window=self)
        self.main_tabs.addTab(self._sql_archive_widget, "SQL")
        self.main_tabs.currentChanged.connect(self._on_main_tab_changed)
        content_layout.addWidget(self.main_tabs, 1)
        layout.addLayout(content_layout)

    def _build_clients_tab(self) -> QWidget:
        return self._build_client_workspace_page()

    def _build_tools_placeholder_page(self, title: str, hint: str) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(12)
        title_lbl = QLabel(title)
        title_lbl.setObjectName("sectionTitle")
        hint_lbl = QLabel(hint)
        hint_lbl.setObjectName("subText")
        hint_lbl.setWordWrap(True)
        layout.addWidget(title_lbl)
        layout.addWidget(hint_lbl)
        layout.addStretch(1)
        return page

    def _build_tools_tab(self) -> QWidget:
        tab = QWidget()
        tab.setObjectName("toolsTabPage")
        tab.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        root = QHBoxLayout(tab)
        root.setContentsMargins(*SIDE_MENU_TAB_CONTENT_MARGINS)
        root.setSpacing(0)

        self.tools_menu = QListWidget()
        self.tools_menu.setObjectName("toolsSideMenu")
        self.tools_menu.setFixedWidth(SIDE_MENU_WIDTH_PX)
        self.tools_menu.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.tools_menu.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)

        self.tools_stack = QStackedWidget()
        self.tools_stack.setObjectName("toolsStack")

        self.tools_menu.addItem(QListWidgetItem("Calcolatrice"))
        self.tools_stack.addWidget(CalculatorWorkspaceWidget(main_window=self))

        self.tools_menu.currentRowChanged.connect(self.tools_stack.setCurrentIndex)
        self.tools_menu.setCurrentRow(0)

        root.addWidget(self.tools_menu)
        root.addWidget(self.tools_stack, 1)
        return tab

    def open_settings(self) -> None:
        if self.settings_window is None:
            self.settings_window = SettingsWindow(self.repository, None)
            self.settings_window.data_changed.connect(self.refresh_views)
            self.settings_window.destroyed.connect(self._clear_settings_window)
        self.settings_window.show()
        self.settings_window.raise_()
        self.settings_window.activateWindow()

    def _ensure_agenda_window(self) -> None:
        if self._agenda_window is None:
            w = QMainWindow(self)
            w.setWindowTitle("Agenda")
            w.setWindowIcon(self.windowIcon())
            w.setWindowModality(Qt.WindowModality.NonModal)
            w.setStyleSheet(self.styleSheet())
            w.setCentralWidget(self._agenda_workspace_widget)
            self._agenda_window = w

    def open_agenda_window(self) -> None:
        self._ensure_agenda_window()
        assert self._agenda_window is not None
        # Massimizza lo spazio utile restando una finestra normale con cornice di sistema
        # (chiudi / riduci / ingrandisci). showFullScreen() toglierebbe la barra titolo.
        self._agenda_window.showMaximized()
        self._agenda_window.raise_()
        self._agenda_window.activateWindow()
        self.refresh_views()

    def _ensure_notes_window(self) -> None:
        if self._notes_window is None:
            w = QMainWindow(self)
            w.setWindowTitle("Note")
            w.setWindowIcon(self.windowIcon())
            w.setWindowModality(Qt.WindowModality.NonModal)
            w.setStyleSheet(self.styleSheet())
            w.resize(640, 520)
            w.setMinimumSize(420, 380)
            w.setCentralWidget(self._notes_workspace_widget)
            self._notes_window = w

    def open_notes_window(self) -> None:
        self._ensure_notes_window()
        assert self._notes_window is not None
        self._notes_window.show()
        self._notes_window.raise_()
        self._notes_window.activateWindow()
        self.refresh_views()

    def _clear_settings_window(self) -> None:
        self.settings_window = None

    def _on_main_tab_changed(self, _index: int) -> None:
        self.refresh_views()

    def closeEvent(self, event) -> None:
        if self.settings_window is not None:
            self.settings_window.close()
        if self._agenda_window is not None:
            self._agenda_window.close()
        if self._notes_window is not None:
            self._notes_window.close()
        super().closeEvent(event)

    def changeEvent(self, event: QEvent) -> None:
        super().changeEvent(event)
        if event.type() == QEvent.Type.ActivationChange and self.isActiveWindow():
            if hasattr(self, "agenda_upcoming_header"):
                self.agenda_upcoming_header.refresh()
            if hasattr(self, "sticky_notes_header"):
                self.sticky_notes_header.refresh()

    def _apply_sticky_notes_visibility(self) -> None:
        show = True
        repo = self.repository
        if hasattr(repo, "settings") and repo.settings is not None:
            show = repo.settings.get_notes_widget_enabled()
        if hasattr(self, "sticky_notes_header"):
            self.sticky_notes_header.setVisible(show)

    def _apply_agenda_header_visibility(self) -> None:
        show = True
        repo = self.repository
        if hasattr(repo, "settings") and repo.settings is not None:
            show = repo.settings.get_agenda_header_widget_enabled()
        if hasattr(self, "agenda_upcoming_header"):
            self.agenda_upcoming_header.setVisible(show)

    def _sync_header_middle_min_width(self) -> None:
        inner = getattr(self, "_header_middle_inner", None)
        if inner is None:
            return
        w = 0
        if hasattr(self, "sticky_notes_header") and self.sticky_notes_header.isVisible():
            w += self.sticky_notes_header.minimumWidth()
        if hasattr(self, "agenda_upcoming_header") and self.agenda_upcoming_header.isVisible():
            w += self.agenda_upcoming_header.minimumWidth()
        if (
            hasattr(self, "sticky_notes_header")
            and hasattr(self, "agenda_upcoming_header")
            and self.sticky_notes_header.isVisible()
            and self.agenda_upcoming_header.isVisible()
        ):
            w += 12
        inner.setMinimumWidth(max(w + 24, 160))
        inner.setMinimumHeight(142)
        inner.adjustSize()
        self._header_middle.updateGeometry()

    def refresh_views(self) -> None:
        self._apply_sticky_notes_visibility()
        self._apply_agenda_header_visibility()
        self._sync_header_middle_min_width()
        if hasattr(self, "sticky_notes_header"):
            self.sticky_notes_header.refresh()
        if hasattr(self, "agenda_upcoming_header"):
            self.agenda_upcoming_header.refresh()
        aw = getattr(self, "_agenda_workspace_widget", None)
        if aw is not None:
            aw.refresh()
        nw = getattr(self, "_notes_workspace_widget", None)
        if nw is not None:
            nw.refresh()
        sql_w = getattr(self, "_sql_archive_widget", None)
        if sql_w is not None:
            sql_w.refresh()
        selected_client_id, selected_product_id = self._selected_archive_selection()
        self._load_data()
        self._render_favorites()
        self._render_tags()
        self._render_archive_overview()
        self._render_archive_client_list(selected_client_id, selected_product_id)

    def _load_data(self) -> None:
        # MVC: if a controller is injected as `repository`, use its cache.
        if hasattr(self.repository, "refresh_cache") and hasattr(self.repository, "cache"):
            self.repository.refresh_cache()
            cache = getattr(self.repository, "cache")
            self.clients_cache = list(getattr(cache, "clients", []) or [])
            self.products_cache = list(getattr(cache, "products", []) or [])
            self.resources_cache = list(getattr(cache, "resources", []) or [])
            self.vpns_cache = list(getattr(cache, "vpns", []) or [])
            self.roles_cache = list(getattr(cache, "roles", []) or [])
            self.tags_cache = list(getattr(cache, "tags", []) or [])
            self.archive_folders_cache = list(getattr(cache, "archive_folders", []) or [])
            # Files are loaded on-demand by folder; keep links cache for global searches
            self.archive_links_cache = list(getattr(cache, "archive_links", []) or [])
        else:
            # Legacy fallback path (pre-MVC)
            self.clients_cache = self.repository.list_clients()
            self.products_cache = self.repository.list_products()
            self.resources_cache = self.repository.list_resources()
            self.vpns_cache = self.repository.list_vpns()
            self.roles_cache = self.repository.list_roles()
            self.tags_cache = self.repository.list_tags()
            self.archive_folders_cache = self.repository.list_archive_folders()
            self.archive_links_cache = self.repository.list_archive_links()
        self.role_multi_map = {
            str(role.get("name", "")).strip().lower(): self._is_truthy(role.get("multi_clients"))
            for role in self.roles_cache
        }
        self.role_order_map = {}
        for role in self.roles_cache:
            role_name = str(role.get("name", "")).strip().lower()
            raw_order = role.get("display_order")
            if not role_name:
                continue
            try:
                self.role_order_map[role_name] = int(raw_order) if raw_order is not None else 999
            except (TypeError, ValueError):
                self.role_order_map[role_name] = 999
        self.clients_by_id = {
            int(row["id"]): row for row in self.clients_cache if row.get("id") is not None
        }

    def _current_archive_folder_id(self) -> int | None:
        if not hasattr(self, "archive_folder_tree"):
            return None
        item = self.archive_folder_tree.currentItem()
        if item is None:
            return None
        folder_id = item.data(0, Qt.ItemDataRole.UserRole)
        return int(folder_id) if folder_id is not None else None

    @staticmethod
    def _format_size(value: object) -> str:
        try:
            size = int(value)
        except (TypeError, ValueError):
            return ""
        if size < 1024:
            return f"{size} B"
        if size < 1024 * 1024:
            return f"{size / 1024:.1f} KB"
        if size < 1024 * 1024 * 1024:
            return f"{size / (1024 * 1024):.1f} MB"
        return f"{size / (1024 * 1024 * 1024):.1f} GB"

    def _clear_archive_filters(self) -> None:
        if hasattr(self, "archive_filter_name"):
            self.archive_filter_name.setText("")
        if hasattr(self, "archive_filter_ext"):
            self.archive_filter_ext.setCurrentText("Tutte")

    def _clear_layout(self, layout: QLayout) -> None:
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            child_layout = item.layout()
            if widget is not None:
                widget.deleteLater()
            elif child_layout is not None:
                self._clear_layout(child_layout)

    def _product_by_id(self, product_id: int | None) -> dict | None:
        if product_id is None:
            return None
        for product in self.products_cache:
            if int(product["id"]) == int(product_id):
                return product
        return None

    @staticmethod
    def _csv_values(value: str | None) -> list[str]:
        return [chunk.strip() for chunk in (value or "").split(",") if chunk.strip()]

    @staticmethod
    def _is_truthy(value: object) -> bool:
        if isinstance(value, bool):
            return value
        if isinstance(value, int):
            return value == 1
        text = ("" if value is None else str(value)).strip().lower()
        return text in {"1", "true", "si", "yes", "y"}

    def _csv_contains(self, csv_value: str | None, value: str | None) -> bool:
        target = (value or "").strip().lower()
        if not target:
            return False
        return target in {chunk.lower() for chunk in self._csv_values(csv_value)}

    @staticmethod
    def _set_table_item(table: QTableWidget, row: int, col: int, value: object) -> None:
        item = QTableWidgetItem("" if value is None else str(value))
        flags = item.flags()
        flags &= ~Qt.ItemFlag.ItemIsEditable
        item.setFlags(flags)
        table.setItem(row, col, item)

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow, QWidget {
                background: #f3f7fa;
                color: #1e293b;
                font-family: "Segoe UI";
                font-size: 12px;
            }
            /* Agenda: teal come voci selezionate nei menu laterali (#0f766e) */
            #headerAgendaButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #2dd4bf, stop:1 #0d9488);
                color: #ffffff;
                border: 2px solid #99f6e4;
                border-radius: 10px;
                padding: 8px 18px;
                font-weight: 700;
                font-size: 12px;
            }
            #headerAgendaButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #5eead4, stop:1 #14b8a6);
                border-color: #ccfbf1;
            }
            #headerAgendaButton:pressed {
                background: #0f766e;
                border-color: #5eead4;
            }
            /* Note: giallo “post-it”, coerente con le note in evidenza */
            #headerNotesButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #fffbeb, stop:1 #fef3c7);
                color: #78350f;
                border: 2px solid #fcd34d;
                border-radius: 10px;
                padding: 8px 18px;
                font-weight: 700;
                font-size: 12px;
            }
            #headerNotesButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #fff7ed, stop:1 #fde68a);
                border-color: #fbbf24;
                color: #451a03;
            }
            #headerNotesButton:pressed {
                background: #fde68a;
                border-color: #d97706;
            }
            /* Impostazioni: blu primario come pulsanti principali dell’app */
            #headerSettingsButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #60a5fa, stop:1 #2563eb);
                color: #ffffff;
                border: 2px solid #bfdbfe;
                border-radius: 10px;
                padding: 8px 18px;
                font-weight: 700;
                font-size: 12px;
            }
            #headerSettingsButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #3b82f6, stop:1 #1d4ed8);
                border-color: #dbeafe;
            }
            #headerSettingsButton:pressed {
                background: #1e40af;
                border-color: #93c5fd;
            }
            #headerFrame {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #2563eb, stop:0.45 #3b82f6, stop:1 #dbeafe);
                border-bottom: 1px solid #2563eb;
                padding-top: 2px;
            }
            /* Evita la fascia chiara del QWidget globale sopra il gradiente (#f3f7fa). */
            #headerMiddleScroll {
                background: transparent;
                border: none;
            }
            #headerMiddleArea {
                background: transparent;
            }
            #headerLogoLabel {
                background: transparent;
            }
            #stickyNotesHeader {
                background: transparent;
            }
            #stickyNotesCard {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
            }
            #stickyNotesNewButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #3b82f6, stop:1 #2563eb);
                color: #ffffff;
                border: none;
                border-radius: 6px;
                padding: 4px 10px;
                font-weight: 600;
                font-size: 11px;
            }
            #stickyNotesNewButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #2563eb, stop:1 #1d4ed8);
            }
            #stickyNotesCardSeparator {
                background: #e2e8f0;
                border: none;
                max-height: 1px;
            }
            #stickyNotesScroll {
                background: transparent;
            }
            #stickyNotesScrollInner {
                background: transparent;
            }
            #stickyNoteCardMuted {
                color: #64748b;
                font-size: 12px;
            }
            #stickyNoteItemCard {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
            }
            #stickyNoteItemCard:hover {
                border-color: #cbd5e1;
                background: #f8fafc;
            }
            #stickyNoteItemTitle {
                color: #0f172a;
                font-size: 11px;
                font-weight: 700;
            }
            #stickyNoteItemDate {
                color: #64748b;
                font-size: 9px;
            }
            #agendaUpcomingHeader {
                background: transparent;
            }
            #agendaHeaderCard {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
            }
            #agendaHeaderCardTitle {
                color: #0f172a;
                font-size: 11px;
                font-weight: 700;
            }
            #agendaHeaderNewButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #3b82f6, stop:1 #2563eb);
                color: #ffffff;
                border: none;
                border-radius: 6px;
                padding: 3px 8px;
                font-weight: 600;
                font-size: 10px;
            }
            #agendaHeaderNewButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #2563eb, stop:1 #1d4ed8);
            }
            #agendaHeaderCardSeparator {
                background: #e2e8f0;
                border: none;
                max-height: 1px;
            }
            #agendaHeaderCardMuted {
                color: #64748b;
                font-size: 11px;
            }
            #agendaHeaderScroll {
                background: transparent;
            }
            #agendaHeaderScrollInner {
                background: transparent;
            }
            /* Sfondo per tipo impegno */
            #agendaHeaderItemCard[kind="appointment"] {
                background: #eff6ff;
                border: 1px solid #93c5fd;
                border-radius: 8px;
            }
            #agendaHeaderItemCard[kind="task"] {
                background: #fffbeb;
                border: 1px solid #fcd34d;
                border-radius: 8px;
            }
            #agendaHeaderItemCard[kind="vacation"] {
                background: #f0fdf4;
                border: 1px solid #86efac;
                border-radius: 8px;
            }
            #agendaHeaderItemCard[kind="leave"] {
                background: #fdf2f8;
                border: 1px solid #f9a8d4;
                border-radius: 8px;
            }
            QLabel#agendaHeaderItemKind[kind="appointment"] {
                color: #1e40af;
                font-size: 8px;
                font-weight: 800;
                letter-spacing: 0.06em;
            }
            QLabel#agendaHeaderItemKind[kind="task"] {
                color: #b45309;
                font-size: 8px;
                font-weight: 800;
                letter-spacing: 0.06em;
            }
            QLabel#agendaHeaderItemKind[kind="vacation"] {
                color: #166534;
                font-size: 8px;
                font-weight: 800;
                letter-spacing: 0.06em;
            }
            QLabel#agendaHeaderItemKind[kind="leave"] {
                color: #9d174d;
                font-size: 8px;
                font-weight: 800;
                letter-spacing: 0.06em;
            }
            QLabel#agendaHeaderItemTitle[kind="appointment"],
            QLabel#agendaHeaderItemTitle[kind="task"],
            QLabel#agendaHeaderItemTitle[kind="vacation"],
            QLabel#agendaHeaderItemTitle[kind="leave"] {
                font-size: 11px;
                font-weight: 700;
            }
            QLabel#agendaHeaderItemTitle[kind="appointment"] { color: #1e3a8a; }
            QLabel#agendaHeaderItemTitle[kind="task"] { color: #78350f; }
            QLabel#agendaHeaderItemTitle[kind="vacation"] { color: #14532d; }
            QLabel#agendaHeaderItemTitle[kind="leave"] { color: #831843; }
            QLabel#agendaHeaderItemMeta[kind="appointment"] { color: #475569; font-size: 9px; }
            QLabel#agendaHeaderItemMeta[kind="task"] { color: #92400e; font-size: 9px; }
            QLabel#agendaHeaderItemMeta[kind="vacation"] { color: #166534; font-size: 9px; }
            QLabel#agendaHeaderItemMeta[kind="leave"] { color: #9d174d; font-size: 9px; }
            /* Stato: in arrivo / in corso / terminato */
            QLabel#agendaHeaderItemStatus[status="upcoming"] {
                color: #2563eb;
                font-size: 9px;
                font-weight: 600;
            }
            QLabel#agendaHeaderItemStatus[status="ongoing"] {
                color: #16a34a;
                font-size: 9px;
                font-weight: 600;
            }
            QLabel#agendaHeaderItemStatus[status="ended"] {
                color: #94a3b8;
                font-size: 9px;
                font-weight: 600;
            }
            #refreshButton {
                background: #ffffff;
                color: #0f172a;
                border: 1px solid #94a3b8;
                border-radius: 10px;
                padding: 8px 16px;
                font-weight: 600;
            }
            #refreshButton:hover {
                background: #e2e8f0;
            }
            #mainTitle {
                color: #0f172a;
            }
            #subTitle {
                color: #334155;
                font-size: 14px;
            }
            #hintCard {
                background: #e0f2fe;
                border: 1px solid #bae6fd;
                border-radius: 10px;
                padding: 10px;
                color: #0c4a6e;
            }
            #clientAccessTabPage, #clientArchiveTabPage, #clientContactsTabPage, #clientNotesTabPage, #archiveTabPage,
            #toolsTabPage, #excelTabPage, #sqlWorkspacePage, #notesWindowPage {
                background: #f4f7f9;
            }
            #accessVpnCard {
                background: #ffffff;
                border: 1px solid #dbe5ee;
                border-radius: 12px;
            }
            #accessVpnBody {
                background: #f8fafc;
                border-top: 1px solid #e8eef5;
            }
            #accessVpnFieldLabel {
                color: #64748b;
                font-size: 11px;
                font-weight: 600;
            }
            #accessVpnCopyField {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 6px 8px;
                color: #001a41;
                font-size: 12px;
            }
            #vpnConnectButton, #accessCredBtnRdpIp, #accessCredBtnRdpHost, #accessCredBtnUrl, #accessCredBtnRdpConf, #accessEditCredentialButton, #clientFormSecondaryButton, #accessDeleteCredentialButton, #contactsDeleteButton, #notesDeleteButton, #notesCellClearButton {
                min-height: 26px;
                padding: 4px 12px;
                font-size: 11px;
                font-weight: 600;
                border: none;
                border-radius: 8px;
            }
            #clientCompactButton {
                min-height: 26px;
                padding: 4px 12px;
                font-size: 11px;
                font-weight: 600;
                border-radius: 8px;
                background: #ffffff;
                border: 1px solid #cbd5e1;
                color: #001a41;
            }
            #clientCompactButton:hover:enabled {
                background: #f8fafc;
                border-color: #94a3b8;
                color: #0f172a;
            }
            #clientCompactButton:disabled {
                background: #f1f5f9;
                color: #94a3b8;
                border-color: #e2e8f0;
            }
            #vpnConnectButton {
                background: #ea580c;
                color: #ffffff;
            }
            #vpnConnectButton:hover:enabled {
                background: #c2410c;
            }
            #vpnConnectButton:disabled {
                background: #e2e8f0;
                color: #94a3b8;
            }
            #accessCredentialCard, #clientDashboardCard {
                background: #ffffff;
                border: 1px solid #dbe5ee;
                border-radius: 12px;
            }
            #accessCredentialBody {
                background: transparent;
            }
            #accessProductHint {
                color: #475569;
                font-size: 12px;
                background: #f1f5f9;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 10px 12px;
            }
            #accessCredActionBar {
                background: #f8fafc;
                border: 1px solid #e8eef5;
                border-radius: 10px;
            }
            #accessCredBtnRdpIp {
                background: #0f766e;
                color: #ffffff;
            }
            #accessCredBtnRdpIp:hover:enabled {
                background: #115e59;
            }
            #accessCredBtnRdpIp:disabled {
                background: #e2e8f0;
                color: #94a3b8;
            }
            #accessCredBtnRdpHost {
                background: #047857;
                color: #ffffff;
            }
            #accessCredBtnRdpHost:hover:enabled {
                background: #065f46;
            }
            #accessCredBtnRdpHost:disabled {
                background: #e2e8f0;
                color: #94a3b8;
            }
            #accessCredBtnUrl {
                background: #7c3aed;
                color: #ffffff;
            }
            #accessCredBtnUrl:hover:enabled {
                background: #6d28d9;
            }
            #accessCredBtnUrl:disabled {
                background: #e2e8f0;
                color: #94a3b8;
            }
            #accessCredBtnRdpConf {
                background: #ca8a04;
                color: #ffffff;
            }
            #accessCredBtnRdpConf:hover:enabled {
                background: #a16207;
            }
            #accessCredBtnRdpConf:disabled {
                background: #e2e8f0;
                color: #94a3b8;
            }
            #accessEditCredentialButton, #clientFormSecondaryButton {
                background: #334155;
                color: #ffffff;
            }
            #accessEditCredentialButton:hover:enabled, #clientFormSecondaryButton:hover:enabled {
                background: #1e293b;
            }
            #accessEditCredentialButton:disabled, #clientFormSecondaryButton:disabled {
                background: #e2e8f0;
                color: #94a3b8;
            }
            #accessDeleteCredentialButton, #contactsDeleteButton, #notesDeleteButton, #notesCellClearButton {
                background: #dc2626;
                color: #ffffff;
            }
            #accessDeleteCredentialButton:hover:enabled, #contactsDeleteButton:hover:enabled, #notesDeleteButton:hover:enabled, #notesCellClearButton:hover:enabled {
                background: #b91c1c;
            }
            #accessDeleteCredentialButton:disabled, #contactsDeleteButton:disabled, #notesDeleteButton:disabled, #notesCellClearButton:disabled {
                background: #e2e8f0;
                color: #94a3b8;
            }
            #accessCredTableWrap, #clientDashboardTableWrap {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 10px;
            }
            #accessCredentialsTable, #clientArchiveFilesTable, #clientArchiveLinksTable, #contactsTable, #notesSpreadsheetTable, #notesWorkspaceTable, #archiveFavoritesFilesTable, #archiveFavoritesLinksTable, #archiveTagsListTable, #archiveTaggedFilesTable, #archiveTaggedLinksTable, #archiveBrowseFilesTable, #archiveBrowseLinksTable {
                background: #ffffff;
                border: none;
                border-radius: 10px;
                gridline-color: transparent;
                outline: none;
            }
            #accessCredentialsTable::item, #clientArchiveFilesTable::item, #clientArchiveLinksTable::item, #contactsTable::item, #notesSpreadsheetTable::item, #notesWorkspaceTable::item, #archiveFavoritesFilesTable::item, #archiveFavoritesLinksTable::item, #archiveTagsListTable::item, #archiveTaggedFilesTable::item, #archiveTaggedLinksTable::item, #archiveBrowseFilesTable::item, #archiveBrowseLinksTable::item {
                padding: 8px 10px;
                border-bottom: 1px solid #f1f5f9;
            }
            #accessCredentialsTable::item:selected, #clientArchiveFilesTable::item:selected, #clientArchiveLinksTable::item:selected, #contactsTable::item:selected, #notesSpreadsheetTable::item:selected, #notesWorkspaceTable::item:selected, #archiveFavoritesFilesTable::item:selected, #archiveFavoritesLinksTable::item:selected, #archiveTagsListTable::item:selected, #archiveTaggedFilesTable::item:selected, #archiveTaggedLinksTable::item:selected, #archiveBrowseFilesTable::item:selected, #archiveBrowseLinksTable::item:selected {
                background: #ebf3ff;
                color: #001a41;
            }
            #accessCredentialsTable QHeaderView::section, #clientArchiveFilesTable QHeaderView::section, #clientArchiveLinksTable QHeaderView::section, #contactsTable QHeaderView::section, #notesSpreadsheetTable QHeaderView::section, #notesWorkspaceTable QHeaderView::section, #archiveFavoritesFilesTable QHeaderView::section, #archiveFavoritesLinksTable QHeaderView::section, #archiveTagsListTable QHeaderView::section, #archiveTaggedFilesTable QHeaderView::section, #archiveTaggedLinksTable QHeaderView::section, #archiveBrowseFilesTable QHeaderView::section, #archiveBrowseLinksTable QHeaderView::section {
                background: #f1f5f9;
                color: #475569;
                font-size: 11px;
                font-weight: 700;
                border: none;
                border-bottom: 1px solid #e2e8f0;
                padding: 10px 10px;
            }
            #accessCredentialsTable QHeaderView::section:first, #clientArchiveFilesTable QHeaderView::section:first, #clientArchiveLinksTable QHeaderView::section:first, #contactsTable QHeaderView::section:first, #notesSpreadsheetTable QHeaderView::section:first, #notesWorkspaceTable QHeaderView::section:first, #archiveFavoritesFilesTable QHeaderView::section:first, #archiveFavoritesLinksTable QHeaderView::section:first, #archiveTagsListTable QHeaderView::section:first, #archiveTaggedFilesTable QHeaderView::section:first, #archiveTaggedLinksTable QHeaderView::section:first, #archiveBrowseFilesTable QHeaderView::section:first, #archiveBrowseLinksTable QHeaderView::section:first {
                border-top-left-radius: 9px;
            }
            #accessCredentialsTable QHeaderView::section:last, #clientArchiveFilesTable QHeaderView::section:last, #clientArchiveLinksTable QHeaderView::section:last, #contactsTable QHeaderView::section:last, #notesSpreadsheetTable QHeaderView::section:last, #notesWorkspaceTable QHeaderView::section:last, #archiveFavoritesFilesTable QHeaderView::section:last, #archiveFavoritesLinksTable QHeaderView::section:last, #archiveTagsListTable QHeaderView::section:last, #archiveTaggedFilesTable QHeaderView::section:last, #archiveTaggedLinksTable QHeaderView::section:last, #archiveBrowseFilesTable QHeaderView::section:last, #archiveBrowseLinksTable QHeaderView::section:last {
                border-top-right-radius: 9px;
            }
            #archiveBrowserPane {
                background: #f8fafc;
                border: 1px solid #e8eef5;
                border-radius: 10px;
                min-width: 220px;
            }
            #clientNotesSidebar {
                background: #f8fafc;
                border: 1px solid #e8eef5;
                border-radius: 10px;
            }
            #notesTitleEdit {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 8px 10px;
                font-size: 13px;
            }
            QTabWidget#clientNotesTabWidget::pane {
                border: 1px solid #e2e8f0;
                border-radius: 10px;
                background: #ffffff;
                top: -1px;
            }
            QTabWidget#clientNotesTabWidget QTabBar::tab {
                background: #f1f5f9;
                border: 1px solid #e2e8f0;
                border-bottom: none;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                padding: 8px 16px;
                margin-right: 4px;
                font-weight: 600;
                color: #475569;
            }
            QTabWidget#clientNotesTabWidget QTabBar::tab:selected {
                background: #ffffff;
                color: #001a41;
                border-bottom: 1px solid #ffffff;
            }
            QTabWidget#archiveOverviewTabWidget::pane {
                border: 1px solid #e2e8f0;
                border-radius: 10px;
                background: #ffffff;
                top: -1px;
            }
            QTabWidget#archiveOverviewTabWidget QTabBar::tab {
                background: #f1f5f9;
                border: 1px solid #e2e8f0;
                border-bottom: none;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                padding: 8px 16px;
                margin-right: 4px;
                font-weight: 600;
                color: #475569;
            }
            QTabWidget#archiveOverviewTabWidget QTabBar::tab:selected {
                background: #ffffff;
                color: #001a41;
                border-bottom: 1px solid #ffffff;
            }
            #accessCopyFeedback {
                color: #0f766e;
                font-size: 12px;
                font-weight: 600;
            }
            #sectionTitle {
                color: #0f172a;
                font-size: 16px;
                font-weight: 700;
            }
            #subSectionTitle {
                color: #1f2937;
                font-size: 13px;
                font-weight: 700;
            }
            #subText {
                color: #475569;
            }
            #smallBadge {
                background: #e2ecf5;
                border: 1px solid #cfdeec;
                border-radius: 999px;
                padding: 4px 10px;
                color: #1e3a5f;
            }
            #accessActionButton {
                background: #ffffff;
                color: #0f172a;
                border: 1px solid #94a3b8;
                border-radius: 8px;
                padding: 6px 12px;
                font-weight: 600;
            }
            #accessActionButton:hover {
                background: #e0f2fe;
                border-color: #0284c7;
            }
            #accessActionButton:enabled {
                background: #0f766e;
                color: #ffffff;
                border: 1px solid #0f766e;
            }
            #accessActionButton:enabled:hover {
                background: #0b5f58;
            }
            #primaryActionButton {
                background: #2563eb;
                color: #ffffff;
                border: 1px solid #1d4ed8;
                border-radius: 10px;
                padding: 6px 14px;
                font-weight: 700;
            }
            #primaryActionButton:hover {
                background: #1d4ed8;
                border-color: #1e40af;
            }
            #primaryActionButton:disabled {
                background: #bfdbfe;
                color: #ffffff;
                border-color: #bfdbfe;
            }
            #secondaryActionButton {
                background: #ffffff;
                color: #0f172a;
                border: 1px solid #94a3b8;
                border-radius: 10px;
                padding: 6px 14px;
                font-weight: 600;
            }
            #secondaryActionButton:hover:enabled {
                background: #e0f2fe;
                border-color: #0284c7;
            }
            #dangerActionButton {
                background: #dc2626;
                color: #ffffff;
                border: 1px solid #b91c1c;
                border-radius: 10px;
                padding: 6px 14px;
                font-weight: 700;
            }
            #dangerActionButton:hover {
                background: #b91c1c;
                border-color: #991b1b;
            }
            #dangerActionButton:disabled {
                background: #fecaca;
                color: #ffffff;
                border-color: #fecaca;
            }
            #mainTabs::pane {
                border: 1px solid #cbd5e1;
                border-radius: 10px;
                background: #ffffff;
            }
            #mainTabs QTabBar::tab {
                background: #ffffff;
                border: 1px solid #cbd5e1;
                border-bottom: none;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                padding: 8px 16px;
                margin-right: 3px;
            }
            #mainTabs QTabBar::tab:hover {
                background: #e0f2fe;
                border-color: #0284c7;
            }
            #mainTabs QTabBar::tab:selected {
                background: #2563eb;
                border-color: #2563eb;
                color: #ffffff;
                font-weight: 700;
            }
            /* Menu laterali e aree contenuto: stessi valori della scheda Archivio. */
            #archiveSideMenu, #toolsSideMenu, #excelSideMenu, #sqlModeSideMenu, #clientsSideMenu {
                background: #0f172a;
                color: #e2e8f0;
                border: none;
                padding: 8px;
            }
            #archiveSideMenu::item, #toolsSideMenu::item, #excelSideMenu::item, #sqlModeSideMenu::item, #clientsSideMenu::item {
                padding: 10px 12px;
                margin: 4px 6px;
                border-radius: 8px;
            }
            #archiveSideMenu::item:selected, #toolsSideMenu::item:selected, #excelSideMenu::item:selected, #sqlModeSideMenu::item:selected, #clientsSideMenu::item:selected {
                background: #0f766e;
                color: #ffffff;
            }
            #archiveStack {
                background: #f4f7f9;
                border-left: 1px solid #dbe5ee;
            }
            /* Stesso sfondo della scheda Archivio (area contenuto a destra del menu). */
            #toolsStack, #excelToolsStack, #sqlEditorStack {
                background: #f4f7f9;
                border-left: 1px solid #dbe5ee;
            }
            #clientStack {
                background: #ffffff;
                border-left: 1px solid #dbe5ee;
            }
            #sqlArchiveListPane, #sqlArchiveEditorPane {
                background: #ffffff;
                border: 1px solid #dbe5ee;
                border-radius: 12px;
            }
            #sqlArchiveQueryList {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 10px;
                padding: 6px;
            }
            #sqlArchiveQueryList::item {
                padding: 8px 10px;
                border-radius: 8px;
            }
            #sqlArchiveQueryList::item:selected {
                background: #2563eb;
                color: #ffffff;
            }
            #sqlModeColumn {
                background: transparent;
            }
            #excelImportPreviewGroup {
                background: #ffffff;
                border: 1px solid #dbe5ee;
                border-radius: 12px;
                margin-top: 12px;
                padding-top: 14px;
                font-weight: 600;
            }
            #excelImportPreviewGroup::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
                color: #001a41;
            }
            #archiveActionButton {
                background: #ffffff;
                color: #001a41;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                padding: 4px 12px;
                font-weight: 600;
                font-size: 11px;
                min-height: 26px;
            }
            #archiveActionButton:hover:enabled {
                background: #f8fafc;
                border-color: #94a3b8;
            }
            #archiveFolderTree {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 10px;
                padding: 6px;
            }
            #archiveFolderTree::item {
                padding: 6px 8px;
                border-radius: 6px;
            }
            #archiveFolderTree::item:selected {
                background: #2563eb;
                color: #ffffff;
            }
            #archiveFilter {
                background: #ffffff;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                padding: 4px 8px;
            }
            #clientsTree {
                background: #ffffff;
                border: 1px solid #94a3b8;
                border-radius: 10px;
                padding: 6px;
            }
            #clientsTree::item {
                padding: 8px 10px;
                border-radius: 8px;
            }
            #clientsTree::item:selected {
                /* Selection colors are applied programmatically to differentiate client vs product */
                background: transparent;
                color: #0f172a;
            }
            #infoKey {
                color: #475569;
                font-weight: 600;
            }
            #infoValue {
                color: #0f172a;
                font-weight: 600;
            }
            #roleCard {
                background: #f8fbfe;
                border: 1px solid #dbe5ee;
                border-radius: 10px;
            }
            #roleCardTitle {
                color: #0f172a;
                font-size: 13px;
                font-weight: 700;
            }
            #roleCountBadge {
                background: #e0f2fe;
                border: 1px solid #7dd3fc;
                border-radius: 999px;
                color: #075985;
                font-size: 11px;
                font-weight: 700;
                padding: 2px 8px;
            }
            #clientInfoTabPage {
                background: #f4f7f9;
            }
            #clientInfoDataCard, #clientSitePreviewCard, #clientProductsOuterCard {
                background: #ffffff;
                border: 1px solid #dbe5ee;
                border-radius: 12px;
            }
            #clientInfoCardHeader {
                background: #ebf3ff;
                border-top-left-radius: 11px;
                border-top-right-radius: 11px;
                border-bottom: 1px solid #dbe5ee;
            }
            #clientInfoCardHeaderIcon {
                color: #1d63d2;
                font-size: 16px;
            }
            #clientInfoCardHeaderTitle {
                color: #001a41;
                font-size: 13px;
                font-weight: 700;
            }
            #clientInfoDataKey {
                color: #64748b;
                font-size: 12px;
                font-weight: 600;
                min-width: 72px;
            }
            #clientInfoFieldValue {
                color: #001a41;
                font-size: 12px;
                font-weight: 700;
            }
            #clientInfoLinkValue {
                color: #001a41;
                font-size: 12px;
            }
            #clientInfoLinkValue a {
                color: #1565c0;
            }
            #clientInfoDataRow {
                border-bottom: 1px solid #e8eef5;
            }
            #clientInfoDataRowLast {
                border-bottom: none;
            }
            #clientSitePreview {
                background: #f8fafc;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
                min-width: 140px;
            }
            #clientSiteLogo {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
                color: #94a3b8;
                font-size: 11px;
                font-weight: 700;
            }
            #clientSiteUrlBar {
                background: #eef2f6;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
            }
            #clientSiteHost {
                color: #475569;
                font-size: 11px;
                font-weight: 600;
            }
            #clientSiteOpenButton {
                background: #1d63d2;
                color: #ffffff;
                border: none;
                border-radius: 8px;
                padding: 8px 16px;
                font-weight: 600;
                font-size: 12px;
            }
            #clientSiteOpenButton:hover {
                background: #1557b8;
            }
            #clientSiteOpenButton:disabled {
                background: #e2e8f0;
                color: #94a3b8;
            }
            #clientProductsScroll {
                background: transparent;
            }
            #resourcePersonRowScroll {
                background: transparent;
                border: none;
            }
            #resourcePersonRowScroll QScrollBar:horizontal {
                height: 10px;
                background: #e2e8f0;
                border-radius: 5px;
            }
            #resourcePersonRowScroll QScrollBar::handle:horizontal {
                background: #94a3b8;
                border-radius: 5px;
                min-width: 40px;
            }
            #resourcePersonCardH {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #4fc3f7, stop:0.4 #1e88e5, stop:1 #0d47a1);
                border: none;
                border-radius: 14px;
            }
            #resourcePersonCardH QLabel {
                background-color: transparent;
                border: none;
            }
            #resourcePersonHRight {
                background: transparent;
                border: none;
            }
            #resourcePersonPhotoH {
                background: transparent;
                border: none;
            }
            #resourcePersonHName {
                color: #ffffff;
                font-size: 15px;
                font-weight: 800;
            }
            #resourcePersonHCompetence {
                color: #bae6fd;
                font-size: 12px;
                font-weight: 600;
            }
            #resourcePersonHEmail {
                color: #ffffff;
                font-size: 12px;
            }
            #resourcePersonHEmail a {
                color: #ffffff;
                text-decoration: none;
            }
            #resourcePersonHEmail a:hover {
                color: #e3f2fd;
                text-decoration: underline;
            }
            #resourcePersonHPhoneIcon {
                color: #ffffff;
                font-size: 12px;
            }
            #resourcePersonHPhone {
                color: #ffffff;
                font-size: 12px;
            }
            #resourcePersonHPhone a {
                color: #ffffff;
                text-decoration: none;
            }
            #resourcePersonHPhone a:hover {
                color: #e3f2fd;
                text-decoration: underline;
            }
            #resourcePersonHRule {
                background: rgba(255, 255, 255, 0.28);
                border: none;
                max-height: 1px;
            }
            #resourcePersonHInBadge {
                color: #ffffff;
                font-size: 10px;
                font-weight: 900;
                background: rgba(255, 255, 255, 0.22);
                border-radius: 3px;
                padding: 1px 5px;
            }
            #resourcePersonHLinkedin {
                color: #ffffff;
                font-size: 11px;
            }
            #resourcePersonHLinkedin a {
                color: #ffffff;
                text-decoration: none;
            }
            #resourcePersonHLinkedin a:hover {
                color: #e3f2fd;
                text-decoration: underline;
            }
            #resourcePersonHNote {
                color: rgba(255, 255, 255, 0.78);
                font-size: 10px;
                font-style: italic;
            }
            #productLinkCard {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
            }
            #productLinkCardSelected {
                background: #f8fbff;
                border: 2px solid #1d63d2;
                border-radius: 12px;
            }
            #productLinkCardTop {
                background: #f1f5f9;
                border-radius: 11px;
            }
            #productLinkTitle {
                color: #001a41;
                font-size: 13px;
                font-weight: 700;
            }
            #productLinkSubtitle {
                color: #475569;
                font-size: 12px;
            }
            #productSelectedBadge {
                background: #e3f0ff;
                border: 1px solid #93c5fd;
                border-radius: 999px;
                color: #1d63d2;
                font-size: 10px;
                font-weight: 700;
                padding: 2px 8px;
            }
            QTableWidget {
                background: #ffffff;
                border: 1px solid #cbd5e1;
                border-radius: 10px;
                gridline-color: #e7edf3;
                selection-background-color: #2563eb;
                selection-color: #ffffff;
            }
            QHeaderView::section {
                background: #0f172a;
                border: 1px solid #0f172a;
                padding: 6px;
                color: #ffffff;
                font-weight: 700;
            }
            QTabWidget::pane {
                border: 1px solid #cbd5e1;
                border-radius: 10px;
                background: #ffffff;
            }
            QTabBar::tab {
                background: #ffffff;
                border: 1px solid #cbd5e1;
                border-bottom: none;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                padding: 7px 12px;
                margin-right: 3px;
            }
            QTabBar::tab:hover {
                background: #e0f2fe;
                border-color: #0284c7;
            }
            QTabBar::tab:selected {
                background: #2563eb;
                border-color: #2563eb;
                color: #ffffff;
                font-weight: 700;
            }
            QGroupBox {
                border: 1px solid #dbe5ee;
                border-radius: 10px;
                margin-top: 8px;
                background: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
                color: #0f172a;
                font-weight: 600;
            }
            """
        )
