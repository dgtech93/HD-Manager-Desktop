from __future__ import annotations

from pathlib import Path

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QIcon, QPainter, QPainterPath, QPen, QPixmap
from PyQt6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QLayout,
    QMainWindow,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QSizePolicy,
    QWidget,
)

from app.ui.settings_window import SettingsWindow
from app.ui.clients_mixin import ClientsMixin
from app.ui.archive_mixin import ArchiveMixin


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

        self.setWindowTitle("HD Manager Desktop")
        self.resize(1380, 880)
        self._build_ui()
        self._apply_style()
        self.refresh_views()

    # UI
    def _build_ui(self) -> None:
        root = QWidget()
        self.setCentralWidget(root)

        layout = QVBoxLayout(root)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        header_frame = QFrame()
        header_frame.setObjectName("headerFrame")
        header_frame.setMinimumHeight(100)
        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(24, 20, 24, 20)
        header_layout.setSpacing(8)

        top_row = QHBoxLayout()
        top_row.setSpacing(16)
        self.settings_btn = QPushButton("Impostazioni")
        self.settings_btn.setObjectName("settingsButton")
        self.settings_btn.clicked.connect(self.open_settings)
        top_row.addWidget(self.settings_btn)

        top_row.addStretch()
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
        top_row.addWidget(self.header_logo_label, 0, Qt.AlignmentFlag.AlignRight)

        header_layout.addLayout(top_row)
        layout.addWidget(header_frame)

        content_layout = QVBoxLayout()
        content_layout.setContentsMargins(18, 16, 18, 18)
        content_layout.setSpacing(14)

        self.main_tabs = QTabWidget()
        self.main_tabs.setObjectName("mainTabs")
        self.main_tabs.addTab(self._build_clients_tab(), "Clienti")
        self.main_tabs.addTab(self._build_archive_tab(), "Archivio")
        self.main_tabs.currentChanged.connect(self._on_main_tab_changed)
        content_layout.addWidget(self.main_tabs, 1)
        layout.addLayout(content_layout)

    def _build_clients_tab(self) -> QWidget:
        return self._build_client_workspace_page("Pagina Cliente")

    def open_settings(self) -> None:
        if self.settings_window is None:
            self.settings_window = SettingsWindow(self.repository, None)
            self.settings_window.data_changed.connect(self.refresh_views)
            self.settings_window.destroyed.connect(self._clear_settings_window)
        self.settings_window.show()
        self.settings_window.raise_()
        self.settings_window.activateWindow()

    def _clear_settings_window(self) -> None:
        self.settings_window = None

    def _on_main_tab_changed(self, _index: int) -> None:
        self.refresh_views()

    def closeEvent(self, event) -> None:
        if self.settings_window is not None:
            self.settings_window.close()
        super().closeEvent(event)

    def refresh_views(self) -> None:
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
            #settingsButton {
                background: #0f766e;
                color: #ffffff;
                border: none;
                border-radius: 10px;
                padding: 8px 16px;
                font-weight: 600;
            }
            #settingsButton:hover {
                background: #0b5f58;
            }
            #headerFrame {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #2563eb, stop:0.45 #3b82f6, stop:1 #dbeafe);
                border-bottom: 1px solid #2563eb;
            }
            #headerLogoLabel {
                background: transparent;
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
            #vpnCard {
                background: #ffffff;
                border: 1px solid #cbd5e1;
                border-radius: 12px;
            }
            #vpnToggle {
                background: #ffffff;
                border: 1px solid #94a3b8;
                border-radius: 10px;
                padding: 6px 10px;
                color: #0f172a;
                font-weight: 700;
            }
            #vpnToggle:hover {
                background: #e0f2fe;
                border-color: #0284c7;
            }
            #vpnToggle:checked {
                background: #0f766e;
                border-color: #0f766e;
                color: #ffffff;
            }
            #copyField {
                background: #ffffff;
                border: 1px solid #d7e2ee;
                border-radius: 8px;
                padding: 5px 8px;
                color: #0f172a;
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
            #archiveSideMenu {
                background: #0f172a;
                color: #e2e8f0;
                border: none;
                padding: 8px;
            }
            #archiveSideMenu::item {
                padding: 10px 12px;
                margin: 4px 6px;
                border-radius: 8px;
            }
            #archiveSideMenu::item:selected {
                background: #0f766e;
                color: #ffffff;
            }
            #archiveStack {
                background: #ffffff;
                border-left: 1px solid #dbe5ee;
            }
            #archiveActionButton {
                background: #ffffff;
                color: #0f172a;
                border: 1px solid #94a3b8;
                border-radius: 8px;
                padding: 6px 12px;
                font-weight: 600;
            }
            #archiveActionButton:hover {
                background: #e0f2fe;
                border-color: #0284c7;
            }
            #archiveFolderTree {
                background: #ffffff;
                border: 1px solid #94a3b8;
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
