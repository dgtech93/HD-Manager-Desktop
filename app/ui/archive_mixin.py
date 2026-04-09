from __future__ import annotations

import os

from PyQt6.QtCore import QFileInfo, QSize, Qt, QUrl
from PyQt6.QtGui import QColor, QDesktopServices, QIcon, QPixmap
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFileIconProvider,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QMessageBox,
    QInputDialog,
    QPushButton,
    QSplitter,
    QStyle,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QStackedWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from app.ui.dialogs import LinkDialog
from app.ui.ui_constants import SIDE_MENU_TAB_CONTENT_MARGINS, SIDE_MENU_WIDTH_PX


class ArchiveMixin:
    # Archive UI + behaviors

    def _archive_file_icon_provider(self) -> QFileIconProvider:
        if not hasattr(self, "_archive_fip"):
            self._archive_fip = QFileIconProvider()
        return self._archive_fip

    def _archive_icon_for_file(self, path: str | None, extension: str | None = None) -> QIcon:
        p = (path or "").strip()
        if p and os.path.isfile(p):
            return self._archive_file_icon_provider().icon(QFileInfo(p))
        ext = (extension or "").strip().lstrip(".").lower()
        if ext:
            dummy = os.path.join(
                os.environ.get("TEMP", os.path.expanduser("~")),
                f"_hdm_icon.{ext}",
            )
            return self._archive_file_icon_provider().icon(QFileInfo(dummy))
        return self._archive_file_icon_provider().icon(QFileIconProvider.IconType.File)

    def _archive_icon_for_link(self, url: str | None) -> QIcon:
        raw = (url or "").strip()
        low = raw.lower()
        style = self.style()
        if low.startswith("mailto:"):
            ic = QIcon.fromTheme("mail-message")
            if not ic.isNull():
                return ic
            if style is not None:
                return style.standardIcon(QStyle.StandardPixmap.SP_DirLinkIcon)
        if low.startswith("ftp"):
            ic = QIcon.fromTheme("folder-remote")
            if not ic.isNull():
                return ic
        ic = QIcon.fromTheme("internet-web-browser")
        if not ic.isNull():
            return ic
        if style is not None:
            return style.standardIcon(QStyle.StandardPixmap.SP_DriveNetIcon)
        return QIcon()

    def _apply_archive_row_icon_file(
        self,
        table: QTableWidget,
        row: int,
        path: str | None,
        extension: str | None,
    ) -> None:
        it = table.item(row, 0)
        if it is not None:
            it.setIcon(self._archive_icon_for_file(path, extension))

    def _apply_archive_row_icon_link(self, table: QTableWidget, row: int, url: str | None) -> None:
        it = table.item(row, 0)
        if it is not None:
            it.setIcon(self._archive_icon_for_link(url))

    @staticmethod
    def _archive_style_table_icons(table: QTableWidget) -> None:
        table.setIconSize(QSize(20, 20))

    def _build_archive_tab(self) -> QWidget:
        page = QWidget()
        layout = QHBoxLayout(page)
        layout.setContentsMargins(*SIDE_MENU_TAB_CONTENT_MARGINS)
        layout.setSpacing(0)
    
        self.archive_menu = QListWidget()
        self.archive_menu.setObjectName("archiveSideMenu")
        self.archive_menu.setFixedWidth(SIDE_MENU_WIDTH_PX)
        self.archive_menu.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.archive_menu.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        for label in ("Archivio", "Tag", "Preferiti"):
            self.archive_menu.addItem(QListWidgetItem(label))
        self.archive_menu.currentRowChanged.connect(self._on_archive_menu_changed)

        self.archive_stack = QStackedWidget()
        self.archive_stack.setObjectName("archiveStack")
        self.archive_stack.addWidget(self._build_archive_overview_page())
        self.archive_stack.addWidget(self._build_archive_tags_page())
        self.archive_stack.addWidget(self._build_archive_favorites_page())

        layout.addWidget(self.archive_menu)
        layout.addWidget(self.archive_stack, 1)

        self.archive_menu.setCurrentRow(0)
        return page
    

    def _build_archive_favorites_page(self) -> QWidget:
        page = QWidget()
        page.setObjectName("archiveTabPage")
        page.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        layout = QVBoxLayout(page)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(16)

        card = QFrame()
        card.setObjectName("clientDashboardCard")
        card.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        card_l = QVBoxLayout(card)
        card_l.setContentsMargins(0, 0, 0, 0)
        card_l.setSpacing(0)
        card_l.addWidget(self._client_info_card_header("⭐", "Preferiti"))

        body = QWidget()
        body_l = QVBoxLayout(body)
        body_l.setContentsMargins(14, 14, 14, 14)
        body_l.setSpacing(12)

        self.favorites_hint_lbl = QLabel(
            "File e link contrassegnati come preferiti (tasto destro nell’archivio)."
        )
        self.favorites_hint_lbl.setObjectName("accessProductHint")
        self.favorites_hint_lbl.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.favorites_hint_lbl.setWordWrap(True)
        body_l.addWidget(self.favorites_hint_lbl)

        lf = QLabel("File")
        lf.setObjectName("subText")
        body_l.addWidget(lf)

        self.favorites_table = QTableWidget(0, 2)
        self._archive_style_table_icons(self.favorites_table)
        self.favorites_table.setObjectName("archiveFavoritesFilesTable")
        self.favorites_table.setHorizontalHeaderLabels(["Nome", "Percorso/URL"])
        self.favorites_table.verticalHeader().setVisible(False)
        self.favorites_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.favorites_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.favorites_table.setAlternatingRowColors(True)
        self.favorites_table.setShowGrid(False)
        self.favorites_table.setColumnWidth(0, 260)
        self.favorites_table.horizontalHeader().setStretchLastSection(True)
        self.favorites_table.cellDoubleClicked.connect(self._open_favorite_file)
        self.favorites_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.favorites_table.customContextMenuRequested.connect(
            self._on_favorites_files_menu
        )
        ftw = QFrame()
        ftw.setObjectName("clientDashboardTableWrap")
        ftl = QVBoxLayout(ftw)
        ftl.setContentsMargins(0, 0, 0, 0)
        ftl.addWidget(self.favorites_table, 1)
        body_l.addWidget(ftw, 1)

        ll = QLabel("Link")
        ll.setObjectName("subText")
        body_l.addWidget(ll)

        self.favorites_links_table = QTableWidget(0, 2)
        self._archive_style_table_icons(self.favorites_links_table)
        self.favorites_links_table.setObjectName("archiveFavoritesLinksTable")
        self.favorites_links_table.setHorizontalHeaderLabels(["Nome", "URL"])
        self.favorites_links_table.verticalHeader().setVisible(False)
        self.favorites_links_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.favorites_links_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.favorites_links_table.setAlternatingRowColors(True)
        self.favorites_links_table.setShowGrid(False)
        self.favorites_links_table.setColumnWidth(0, 260)
        self.favorites_links_table.horizontalHeader().setStretchLastSection(True)
        self.favorites_links_table.cellDoubleClicked.connect(self._open_favorite_link)
        self.favorites_links_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.favorites_links_table.customContextMenuRequested.connect(
            self._on_favorites_links_menu
        )
        ltw = QFrame()
        ltw.setObjectName("clientDashboardTableWrap")
        ltl = QVBoxLayout(ltw)
        ltl.setContentsMargins(0, 0, 0, 0)
        ltl.addWidget(self.favorites_links_table, 1)
        body_l.addWidget(ltw, 1)

        card_l.addWidget(body, 1)
        layout.addWidget(card, 1)
        return page
    

    def _build_archive_tags_page(self) -> QWidget:
        page = QWidget()
        page.setObjectName("archiveTabPage")
        page.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        outer = QVBoxLayout(page)
        outer.setContentsMargins(16, 16, 16, 16)
        outer.setSpacing(0)

        card = QFrame()
        card.setObjectName("clientDashboardCard")
        card.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        card_l = QVBoxLayout(card)
        card_l.setContentsMargins(0, 0, 0, 0)
        card_l.setSpacing(0)
        card_l.addWidget(self._client_info_card_header("🏷️", "Tag"))

        body = QWidget()
        body_l = QVBoxLayout(body)
        body_l.setContentsMargins(14, 14, 14, 14)
        body_l.setSpacing(12)

        hint = QLabel(
            "Tag definiti in Impostazioni. Seleziona un tag a sinistra per filtrare file e link."
        )
        hint.setObjectName("accessProductHint")
        hint.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        hint.setWordWrap(True)
        body_l.addWidget(hint)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)

        left = QFrame()
        left.setObjectName("archiveBrowserPane")
        left.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(10, 10, 10, 10)
        left_layout.setSpacing(8)
        tl = QLabel("Tag disponibili")
        tl.setObjectName("subText")
        left_layout.addWidget(tl)

        self.tags_table = QTableWidget(0, 1)
        self.tags_table.setObjectName("archiveTagsListTable")
        self.tags_table.setHorizontalHeaderLabels(["Tag"])
        self.tags_table.verticalHeader().setVisible(False)
        self.tags_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tags_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tags_table.setAlternatingRowColors(True)
        self.tags_table.setShowGrid(False)
        self.tags_table.setColumnWidth(0, 220)
        self.tags_table.horizontalHeader().setStretchLastSection(True)
        self.tags_table.cellClicked.connect(self._on_archive_tag_selected)
        tgw = QFrame()
        tgw.setObjectName("clientDashboardTableWrap")
        tgl = QVBoxLayout(tgw)
        tgl.setContentsMargins(0, 0, 0, 0)
        tgl.addWidget(self.tags_table, 1)
        left_layout.addWidget(tgw, 1)
        splitter.addWidget(left)

        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(10)

        files_title = QLabel("File con questo tag")
        files_title.setObjectName("subSectionTitle")
        right_layout.addWidget(files_title)

        self.tags_files_table = QTableWidget(0, 3)
        self._archive_style_table_icons(self.tags_files_table)
        self.tags_files_table.setObjectName("archiveTaggedFilesTable")
        self.tags_files_table.setHorizontalHeaderLabels(["Nome", "Percorso", "Cartella"])
        self.tags_files_table.verticalHeader().setVisible(False)
        self.tags_files_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tags_files_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tags_files_table.setAlternatingRowColors(True)
        self.tags_files_table.setShowGrid(False)
        self.tags_files_table.setColumnWidth(0, 240)
        self.tags_files_table.setColumnWidth(1, 360)
        self.tags_files_table.setColumnWidth(2, 200)
        self.tags_files_table.horizontalHeader().setStretchLastSection(True)
        self.tags_files_table.cellDoubleClicked.connect(self._on_tags_files_page_double_click)
        tfw = QFrame()
        tfw.setObjectName("clientDashboardTableWrap")
        tfl = QVBoxLayout(tfw)
        tfl.setContentsMargins(0, 0, 0, 0)
        tfl.addWidget(self.tags_files_table, 1)
        right_layout.addWidget(tfw, 1)

        links_title = QLabel("Link con questo tag")
        links_title.setObjectName("subSectionTitle")
        right_layout.addWidget(links_title)

        self.tags_links_table = QTableWidget(0, 3)
        self._archive_style_table_icons(self.tags_links_table)
        self.tags_links_table.setObjectName("archiveTaggedLinksTable")
        self.tags_links_table.setHorizontalHeaderLabels(["Nome", "URL", "Cartella"])
        self.tags_links_table.verticalHeader().setVisible(False)
        self.tags_links_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tags_links_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tags_links_table.setAlternatingRowColors(True)
        self.tags_links_table.setShowGrid(False)
        self.tags_links_table.setColumnWidth(0, 240)
        self.tags_links_table.setColumnWidth(1, 360)
        self.tags_links_table.setColumnWidth(2, 200)
        self.tags_links_table.horizontalHeader().setStretchLastSection(True)
        self.tags_links_table.cellDoubleClicked.connect(self._on_tags_links_page_double_click)
        tlw = QFrame()
        tlw.setObjectName("clientDashboardTableWrap")
        tll = QVBoxLayout(tlw)
        tll.setContentsMargins(0, 0, 0, 0)
        tll.addWidget(self.tags_links_table, 1)
        right_layout.addWidget(tlw, 1)

        splitter.addWidget(right)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([280, 780])
        body_l.addWidget(splitter, 1)

        card_l.addWidget(body, 1)
        outer.addWidget(card, 1)
        return page
    

    def _build_archive_overview_page(self) -> QWidget:
        page = QWidget()
        page.setObjectName("archiveTabPage")
        page.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        outer = QVBoxLayout(page)
        outer.setContentsMargins(16, 16, 16, 16)
        outer.setSpacing(0)

        card = QFrame()
        card.setObjectName("clientDashboardCard")
        card.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        card_l = QVBoxLayout(card)
        card_l.setContentsMargins(0, 0, 0, 0)
        card_l.setSpacing(0)
        card_l.addWidget(self._client_info_card_header("🗂️", "Esplora archivio"))

        body = QWidget()
        body_l = QVBoxLayout(body)
        body_l.setContentsMargins(14, 14, 14, 14)
        body_l.setSpacing(0)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)

        left = QFrame()
        left.setObjectName("archiveBrowserPane")
        left.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(10, 10, 10, 10)
        left_layout.setSpacing(8)
        cl = QLabel("Cartelle")
        cl.setObjectName("subText")
        left_layout.addWidget(cl)

        self.archive_folder_tree = QTreeWidget()
        self.archive_folder_tree.setObjectName("archiveFolderTree")
        self.archive_folder_tree.setHeaderHidden(True)
        self.archive_folder_tree.currentItemChanged.connect(self._on_archive_folder_selected)
        self.archive_folder_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.archive_folder_tree.customContextMenuRequested.connect(
            self._on_archive_folder_menu
        )
        left_layout.addWidget(self.archive_folder_tree, 1)
        splitter.addWidget(left)

        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(10)

        tabs = QTabWidget()
        tabs.setObjectName("archiveOverviewTabWidget")
        tabs.addTab(self._build_archive_files_page(), "File")
        tabs.addTab(self._build_archive_links_page(), "Link")
        right_layout.addWidget(tabs, 1)
        splitter.addWidget(right)

        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([260, 820])
        body_l.addWidget(splitter, 1)

        card_l.addWidget(body, 1)
        outer.addWidget(card, 1)
        return page
    

    def _build_archive_files_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        toolbar_wrap = QFrame()
        toolbar_wrap.setObjectName("accessCredActionBar")
        toolbar = QHBoxLayout(toolbar_wrap)
        toolbar.setContentsMargins(10, 8, 10, 8)
        toolbar.setSpacing(8)
        self.archive_new_folder_btn = QPushButton("Nuova cartella")
        self.archive_subfolder_btn = QPushButton("Sottocartella")
        self.archive_add_file_btn = QPushButton("Aggiungi file")
        self.archive_delete_file_btn = QPushButton("Elimina file")
        self.archive_refresh_btn = QPushButton("Aggiorna")

        for btn in (
            self.archive_new_folder_btn,
            self.archive_subfolder_btn,
            self.archive_add_file_btn,
            self.archive_refresh_btn,
        ):
            btn.setObjectName("archiveActionButton")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            toolbar.addWidget(btn)
        self.archive_delete_file_btn.setObjectName("dangerActionButton")
        self.archive_delete_file_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        toolbar.addWidget(self.archive_delete_file_btn)
        toolbar.addStretch()
        layout.addWidget(toolbar_wrap)

        filters = QHBoxLayout()
        filters.setSpacing(8)
        filters.addWidget(QLabel("Estensione"))
        self.archive_filter_ext = QComboBox()
        self.archive_filter_ext.setObjectName("archiveFilter")
        self.archive_filter_ext.addItem("Tutte")
        filters.addWidget(self.archive_filter_ext)
        filters.addWidget(QLabel("Nome"))
        self.archive_filter_name = QLineEdit()
        self.archive_filter_name.setObjectName("archiveFilter")
        self.archive_filter_name.setPlaceholderText("Cerca per nome file...")
        filters.addWidget(self.archive_filter_name, 1)
        clear_btn = QPushButton("Pulisci filtri")
        clear_btn.setObjectName("archiveActionButton")
        clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        filters.addWidget(clear_btn)
        layout.addLayout(filters)

        right_layout = QVBoxLayout()
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(8)
        fl = QLabel("Contenuto cartella")
        fl.setObjectName("subText")
        right_layout.addWidget(fl)

        self.archive_files_table = QTableWidget(0, 7)
        self._archive_style_table_icons(self.archive_files_table)
        self.archive_files_table.setObjectName("archiveBrowseFilesTable")
        self.archive_files_table.setHorizontalHeaderLabels(
            ["Nome", "Tipo", "Ultima modifica", "Peso", "Estensione", "Percorso", "Tag"]
        )
        self.archive_files_table.verticalHeader().setVisible(False)
        self.archive_files_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.archive_files_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.archive_files_table.setAlternatingRowColors(True)
        self.archive_files_table.setShowGrid(False)
        self.archive_files_table.setColumnWidth(0, 220)
        self.archive_files_table.setColumnWidth(1, 120)
        self.archive_files_table.setColumnWidth(2, 150)
        self.archive_files_table.setColumnWidth(3, 110)
        self.archive_files_table.setColumnWidth(4, 110)
        self.archive_files_table.setColumnWidth(5, 320)
        self.archive_files_table.horizontalHeader().setStretchLastSection(True)
        self.archive_files_table.cellDoubleClicked.connect(
            self._on_archive_file_double_click
        )
        self.archive_files_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.archive_files_table.customContextMenuRequested.connect(
            self._on_archive_files_menu
        )
        ft_wrap = QFrame()
        ft_wrap.setObjectName("clientDashboardTableWrap")
        ftl = QVBoxLayout(ft_wrap)
        ftl.setContentsMargins(0, 0, 0, 0)
        ftl.addWidget(self.archive_files_table, 1)
        right_layout.addWidget(ft_wrap, 1)
        layout.addLayout(right_layout, 1)
    
        self.archive_new_folder_btn.clicked.connect(self._create_archive_root_folder)
        self.archive_subfolder_btn.clicked.connect(self._create_archive_subfolder)
        self.archive_add_file_btn.clicked.connect(self._add_archive_files)
        self.archive_delete_file_btn.clicked.connect(self._delete_selected_archive_file)
        self.archive_refresh_btn.clicked.connect(self._render_archive_overview)
        self.archive_filter_ext.currentTextChanged.connect(self._render_archive_files)
        self.archive_filter_name.textChanged.connect(self._render_archive_files)
        clear_btn.clicked.connect(self._clear_archive_filters)
        return page
    

    def _build_archive_links_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)

        toolbar_wrap = QFrame()
        toolbar_wrap.setObjectName("accessCredActionBar")
        toolbar = QHBoxLayout(toolbar_wrap)
        toolbar.setContentsMargins(10, 8, 10, 8)
        toolbar.setSpacing(8)
        self.archive_new_link_btn = QPushButton("Nuovo link")
        self.archive_edit_link_btn = QPushButton("Modifica")
        self.archive_delete_link_btn = QPushButton("Elimina")
        self.archive_open_link_btn = QPushButton("Apri link")

        for btn in (
            self.archive_new_link_btn,
            self.archive_edit_link_btn,
            self.archive_open_link_btn,
        ):
            btn.setObjectName("archiveActionButton")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            toolbar.addWidget(btn)
        self.archive_delete_link_btn.setObjectName("dangerActionButton")
        self.archive_delete_link_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        toolbar.addWidget(self.archive_delete_link_btn)
        toolbar.addStretch()
        layout.addWidget(toolbar_wrap)

        filters = QHBoxLayout()
        filters.setSpacing(8)
        filters.addWidget(QLabel("Tag"))
        self.archive_link_filter_tag = QComboBox()
        self.archive_link_filter_tag.setObjectName("archiveFilter")
        self.archive_link_filter_tag.addItem("Tutti")
        filters.addWidget(self.archive_link_filter_tag)
        filters.addWidget(QLabel("Nome"))
        self.archive_link_filter_name = QLineEdit()
        self.archive_link_filter_name.setObjectName("archiveFilter")
        self.archive_link_filter_name.setPlaceholderText("Cerca per nome link...")
        filters.addWidget(self.archive_link_filter_name, 1)
        clear_btn = QPushButton("Pulisci filtri")
        clear_btn.setObjectName("archiveActionButton")
        clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        filters.addWidget(clear_btn)
        layout.addLayout(filters)

        ll = QLabel("Link nella cartella")
        ll.setObjectName("subText")
        layout.addWidget(ll)

        self.archive_links_table = QTableWidget(0, 3)
        self._archive_style_table_icons(self.archive_links_table)
        self.archive_links_table.setObjectName("archiveBrowseLinksTable")
        self.archive_links_table.setHorizontalHeaderLabels(["Nome", "URL", "Tag"])
        self.archive_links_table.verticalHeader().setVisible(False)
        self.archive_links_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.archive_links_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.archive_links_table.setAlternatingRowColors(True)
        self.archive_links_table.setShowGrid(False)
        self.archive_links_table.setColumnWidth(0, 260)
        self.archive_links_table.setColumnWidth(1, 380)
        self.archive_links_table.horizontalHeader().setStretchLastSection(True)
        self.archive_links_table.cellDoubleClicked.connect(self._on_archive_link_double_click)
        self.archive_links_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.archive_links_table.customContextMenuRequested.connect(
            self._on_archive_links_menu
        )
        lw_wrap = QFrame()
        lw_wrap.setObjectName("clientDashboardTableWrap")
        lwl = QVBoxLayout(lw_wrap)
        lwl.setContentsMargins(0, 0, 0, 0)
        lwl.addWidget(self.archive_links_table, 1)
        layout.addWidget(lw_wrap, 1)
    
        self.archive_new_link_btn.clicked.connect(self._create_archive_link)
        self.archive_edit_link_btn.clicked.connect(self._edit_selected_link)
        self.archive_delete_link_btn.clicked.connect(self._delete_selected_link)
        self.archive_open_link_btn.clicked.connect(self._open_selected_link)
        self.archive_link_filter_tag.currentTextChanged.connect(self._render_archive_links)
        self.archive_link_filter_name.textChanged.connect(self._render_archive_links)
        clear_btn.clicked.connect(self._clear_archive_link_filters)
        return page
    

    def _on_archive_menu_changed(self, index: int) -> None:
        if index < 0:
            return
        self.archive_stack.setCurrentIndex(index)
    

    def _render_favorites(self) -> None:
        self.favorites_table.setRowCount(0)
        if hasattr(self, "favorites_links_table"):
            self.favorites_links_table.setRowCount(0)
        favorites = self.repository.archive.list_favorites() if hasattr(self.repository, "archive") else self.repository.list_archive_favorites()
        if not favorites:
            self.favorites_hint_lbl.setText(
                "Nessun preferito ancora disponibile: usa il tasto destro su un file o link."
            )
            return
        self.favorites_hint_lbl.setText("Elementi contrassegnati come preferiti.")
        for fav in favorites:
            item_type = fav.get("item_type")
            if item_type == "file":
                row = self.favorites_table.rowCount()
                self.favorites_table.insertRow(row)
                name = str(fav.get("name", "") or "")
                path = str(fav.get("location", "") or "")
                ext = os.path.splitext(name)[1].lstrip(".").lower() or None
                self._set_table_item(self.favorites_table, row, 0, name)
                self._set_table_item(self.favorites_table, row, 1, path)
                self._apply_archive_row_icon_file(self.favorites_table, row, path, ext)
                id_item = self.favorites_table.item(row, 0)
                if id_item is not None:
                    id_item.setData(Qt.ItemDataRole.UserRole, fav.get("item_id"))
            else:
                row = self.favorites_links_table.rowCount()
                self.favorites_links_table.insertRow(row)
                self._set_table_item(self.favorites_links_table, row, 0, fav.get("name", ""))
                loc = str(fav.get("location", "") or "")
                self._set_table_item(self.favorites_links_table, row, 1, loc)
                self._apply_archive_row_icon_link(self.favorites_links_table, row, loc)
                id_item = self.favorites_links_table.item(row, 0)
                if id_item is not None:
                    id_item.setData(Qt.ItemDataRole.UserRole, fav.get("item_id"))
    

    def _render_archive_overview(self) -> None:
        if hasattr(self, "archive_folder_tree"):
            self._render_archive_folders()
        if hasattr(self, "archive_files_table"):
            self._render_archive_files()
        if hasattr(self, "archive_links_table"):
            self._render_archive_links()
    

    def _render_tags(self) -> None:
        self.tags_table.setRowCount(0)
        if not getattr(self, "tags_cache", None):
            return
        for tag in self.tags_cache:
            row = self.tags_table.rowCount()
            self.tags_table.insertRow(row)
            self._set_table_item(self.tags_table, row, 0, tag.get("name", ""))
            color = str(tag.get("color") or "").strip()
            if color:
                pixmap = QPixmap(10, 10)
                pixmap.fill(QColor(color))
                name_item = self.tags_table.item(row, 0)
                if name_item is not None:
                    name_item.setData(Qt.ItemDataRole.DecorationRole, pixmap)
        if self.tags_table.rowCount() > 0 and self.tags_table.currentRow() < 0:
            self.tags_table.selectRow(0)
            self._on_archive_tag_selected(0, 0)
    

    def _on_archive_tag_selected(self, row: int, column: int) -> None:
        if not hasattr(self, "tags_table"):
            return
        tag_item = self.tags_table.item(row, 0)
        if tag_item is None:
            return
        tag_name = tag_item.text().strip()
        self._render_tagged_files(tag_name)
        self._render_tagged_links(tag_name)

    def _on_tags_files_page_double_click(self, row: int, column: int) -> None:
        """Apre il file con l’applicazione predefinita (doppio clic su «File con Tag»)."""
        if not hasattr(self, "tags_files_table"):
            return
        path_item = self.tags_files_table.item(row, 1)
        if path_item is None:
            return
        path = path_item.text().strip()
        if not path:
            return
        if not os.path.isfile(path):
            QMessageBox.warning(
                self,
                "Archivio",
                "File non trovato sul disco. Potrebbe essere stato spostato o eliminato.",
            )
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def _on_tags_links_page_double_click(self, row: int, column: int) -> None:
        """Apre l’URL del link (doppio clic su «Link con Tag»)."""
        if not hasattr(self, "tags_links_table"):
            return
        url_item = self.tags_links_table.item(row, 1)
        if url_item is None:
            return
        url = url_item.text().strip()
        if not url:
            return
        q = QUrl.fromUserInput(url)
        if not q.isValid():
            QMessageBox.warning(self, "Archivio", "URL non valido.")
            return
        QDesktopServices.openUrl(q)
    

    def _render_tagged_files(self, tag_name: str) -> None:
        if not hasattr(self, "tags_files_table"):
            return
        rows = self.repository.archive.list_files_all() if hasattr(self.repository, "archive") else self.repository.list_archive_files_all()
        folder_map = self._archive_folder_path_map()
        filtered = [row for row in rows if str(row.get("tag_name") or "") == tag_name]
        self.tags_files_table.setRowCount(0)
        for row_data in filtered:
            row = self.tags_files_table.rowCount()
            self.tags_files_table.insertRow(row)
            self._set_table_item(self.tags_files_table, row, 0, row_data.get("name", ""))
            self._set_table_item(self.tags_files_table, row, 1, row_data.get("path", ""))
            folder_label = folder_map.get(row_data.get("folder_id"))
            self._set_table_item(self.tags_files_table, row, 2, folder_label or "Archivio")
            self._apply_archive_row_icon_file(
                self.tags_files_table,
                row,
                str(row_data.get("path") or ""),
                str(row_data.get("extension") or ""),
            )
    

    def _render_tagged_links(self, tag_name: str) -> None:
        if not hasattr(self, "tags_links_table"):
            return
        rows = self.repository.archive.list_links_all() if hasattr(self.repository, "archive") else self.repository.list_archive_links_all()
        folder_map = self._archive_folder_path_map()
        filtered = [row for row in rows if str(row.get("tag_name") or "") == tag_name]
        self.tags_links_table.setRowCount(0)
        for row_data in filtered:
            row = self.tags_links_table.rowCount()
            self.tags_links_table.insertRow(row)
            self._set_table_item(self.tags_links_table, row, 0, row_data.get("name", ""))
            url = str(row_data.get("url") or "")
            self._set_table_item(self.tags_links_table, row, 1, url)
            folder_label = folder_map.get(row_data.get("folder_id"))
            self._set_table_item(self.tags_links_table, row, 2, folder_label or "Archivio")
            self._apply_archive_row_icon_link(self.tags_links_table, row, url)
    

    def _render_archive_folders(self) -> None:
        if not hasattr(self, "archive_folder_tree"):
            return
    
        selected_id = self._current_archive_folder_id()
        self.archive_folder_tree.blockSignals(True)
        self.archive_folder_tree.clear()
    
        root = QTreeWidgetItem(["Archivio"])
        root.setData(0, Qt.ItemDataRole.UserRole, None)
        self.archive_folder_tree.addTopLevelItem(root)
    
        folders = self.repository.archive.list_folders() if hasattr(self.repository, "archive") else self.repository.list_archive_folders()
        items: dict[int, QTreeWidgetItem] = {}
        for folder in folders:
            item = QTreeWidgetItem([folder.get("name", "")])
            folder_id = int(folder["id"])
            item.setData(0, Qt.ItemDataRole.UserRole, folder_id)
            items[folder_id] = item
    
        for folder in folders:
            folder_id = int(folder["id"])
            parent_id = folder.get("parent_id")
            parent_item = items.get(int(parent_id)) if parent_id is not None else root
            if parent_item is None:
                parent_item = root
            parent_item.addChild(items[folder_id])
    
        root.setExpanded(True)
        self.archive_folder_tree.blockSignals(False)
    
        target_item = root
        if selected_id is not None and selected_id in items:
            target_item = items[selected_id]
        self.archive_folder_tree.setCurrentItem(target_item)
    

    def _render_archive_files(self) -> None:
        if not hasattr(self, "archive_files_table"):
            return
        folder_id = self._current_archive_folder_id()
        archive = self.repository.archive if hasattr(self.repository, "archive") else self.repository

        # Extensions combo via SQL DISTINCT
        if hasattr(archive, "list_file_extensions") and hasattr(self, "archive_filter_ext"):
            current = self.archive_filter_ext.currentText()
            extensions = archive.list_file_extensions(folder_id)
            self.archive_filter_ext.blockSignals(True)
            self.archive_filter_ext.clear()
            self.archive_filter_ext.addItem("Tutte")
            for ext in extensions:
                self.archive_filter_ext.addItem(ext)
            if current and current in {"Tutte", *extensions}:
                self.archive_filter_ext.setCurrentText(current)
            self.archive_filter_ext.blockSignals(False)
    
        selected_ext = (
            self.archive_filter_ext.currentText()
            if hasattr(self, "archive_filter_ext")
            else "Tutte"
        )
        name_filter = (
            self.archive_filter_name.text().strip().lower()
            if hasattr(self, "archive_filter_name")
            else ""
        )
        # Main rows via SQL filtering
        if hasattr(archive, "list_files_filtered"):
            filtered_rows = archive.list_files_filtered(
                folder_id,
                extension=selected_ext,
                name_contains=name_filter,
            )
        else:
            rows = archive.list_archive_files(folder_id)
            filtered_rows = []
            for row in rows:
                ext = str(row.get("extension") or "").strip().lower()
                if selected_ext and selected_ext != "Tutte" and ext != selected_ext:
                    continue
                if name_filter and name_filter not in str(row.get("name") or "").lower():
                    continue
                filtered_rows.append(row)
    
        self.archive_files_table.setRowCount(0)
        for row_data in filtered_rows:
            row = self.archive_files_table.rowCount()
            self.archive_files_table.insertRow(row)
            self._set_table_item(self.archive_files_table, row, 0, row_data.get("name", ""))
            self._set_table_item(
                self.archive_files_table, row, 1, row_data.get("file_type", "")
            )
            self._set_table_item(
                self.archive_files_table, row, 2, row_data.get("last_modified", "")
            )
            self._set_table_item(
                self.archive_files_table,
                row,
                3,
                self._format_size(row_data.get("file_size")),
            )
            self._set_table_item(
                self.archive_files_table, row, 4, row_data.get("extension", "")
            )
            self._set_table_item(
                self.archive_files_table, row, 5, row_data.get("path", "")
            )
            self._set_table_item(
                self.archive_files_table, row, 6, row_data.get("tag_name", "")
            )
            self._apply_archive_row_icon_file(
                self.archive_files_table,
                row,
                str(row_data.get("path") or ""),
                str(row_data.get("extension") or ""),
            )
            id_item = self.archive_files_table.item(row, 0)
            if id_item is not None:
                id_item.setData(Qt.ItemDataRole.UserRole, row_data.get("id"))
    
            tag_color = str(row_data.get("tag_color") or "").strip()
            if tag_color:
                tag_item = self.archive_files_table.item(row, 6)
                if tag_item is not None:
                    tag_item.setBackground(QColor(tag_color))
    

    def _render_archive_links(self) -> None:
        if not hasattr(self, "archive_links_table"):
            return
        folder_id = self._current_archive_folder_id()
        archive = self.repository.archive if hasattr(self.repository, "archive") else self.repository
        rows = archive.list_links(folder_id) if hasattr(archive, "list_links") else archive.list_archive_links(folder_id)
    
        if hasattr(self, "archive_link_filter_tag"):
            current = self.archive_link_filter_tag.currentText()
            tags = sorted(
                {str(row.get("tag_name") or "").strip() for row in rows if row.get("tag_name")}
            )
            self.archive_link_filter_tag.blockSignals(True)
            self.archive_link_filter_tag.clear()
            self.archive_link_filter_tag.addItem("Tutti")
            for tag in tags:
                self.archive_link_filter_tag.addItem(tag)
            if current and current in {"Tutti", *tags}:
                self.archive_link_filter_tag.setCurrentText(current)
            self.archive_link_filter_tag.blockSignals(False)
    
        selected_tag = (
            self.archive_link_filter_tag.currentText()
            if hasattr(self, "archive_link_filter_tag")
            else "Tutti"
        )
        name_filter = (
            self.archive_link_filter_name.text().strip().lower()
            if hasattr(self, "archive_link_filter_name")
            else ""
        )
    
        if hasattr(archive, "list_links_filtered"):
            filtered_rows = archive.list_links_filtered(
                folder_id,
                tag_name=selected_tag,
                name_contains=name_filter,
            )
        else:
            filtered_rows = []
            for row in rows:
                tag_name = str(row.get("tag_name") or "").strip()
                if selected_tag and selected_tag != "Tutti" and tag_name != selected_tag:
                    continue
                if name_filter and name_filter not in str(row.get("name") or "").lower():
                    continue
                filtered_rows.append(row)
    
        self.archive_links_table.setRowCount(0)
        for row_data in filtered_rows:
            row = self.archive_links_table.rowCount()
            self.archive_links_table.insertRow(row)
            self._set_table_item(self.archive_links_table, row, 0, row_data.get("name", ""))
            link_url = str(row_data.get("url") or "")
            self._set_table_item(self.archive_links_table, row, 1, link_url)
            self._set_table_item(self.archive_links_table, row, 2, row_data.get("tag_name", ""))
            self._apply_archive_row_icon_link(self.archive_links_table, row, link_url)
            id_item = self.archive_links_table.item(row, 0)
            if id_item is not None:
                id_item.setData(Qt.ItemDataRole.UserRole, row_data.get("id"))
            tag_color = str(row_data.get("tag_color") or "").strip()
            if tag_color:
                tag_item = self.archive_links_table.item(row, 2)
                if tag_item is not None:
                    tag_item.setBackground(QColor(tag_color))
    

    def _archive_folder_path_map(self) -> dict[object, str]:
        folders = self.repository.archive.list_folders() if hasattr(self.repository, "archive") else self.repository.list_archive_folders()
        by_id = {int(row["id"]): row for row in folders}
        path_map: dict[object, str] = {}
    
        def build_path(folder_id: int) -> str:
            parts: list[str] = []
            current = by_id.get(folder_id)
            while current is not None:
                name = str(current.get("name") or "").strip()
                if name:
                    parts.append(name)
                parent_id = current.get("parent_id")
                current = by_id.get(int(parent_id)) if parent_id is not None else None
            return "/".join(reversed(parts)) if parts else "Archivio"
    
        for folder_id in by_id:
            path_map[folder_id] = build_path(folder_id)
        return path_map
    
    def _on_archive_folder_selected(
        self, current: QTreeWidgetItem | None, previous: QTreeWidgetItem | None = None
    ) -> None:
        if current is None:
            return
        self._render_archive_files()
        self._render_archive_links()
    

    def _on_archive_folder_menu(self, pos) -> None:
        item = self.archive_folder_tree.itemAt(pos)
        menu = QMenu(self)
        menu.addAction("Nuova cartella", self._create_archive_root_folder)
        if item is not None:
            menu.addAction("Sottocartella", self._create_archive_subfolder)
            menu.addAction("Elimina cartella", self._delete_selected_archive_folder)
        menu.exec(self.archive_folder_tree.mapToGlobal(pos))
    

    def _create_archive_root_folder(self) -> None:
        name, ok = QInputDialog.getText(self, "Nuova cartella", "Nome cartella:")
        if not ok:
            return
        try:
            if hasattr(self.repository, "archive"):
                self.repository.archive.add_folder(name, None)
            else:
                self.repository.add_archive_folder(name, None)
            self._render_archive_folders()
        except ValueError as exc:
            QMessageBox.warning(self, "Cartella", str(exc))
    

    def _create_archive_subfolder(self) -> None:
        parent_id = self._current_archive_folder_id()
        if parent_id is None:
            QMessageBox.information(self, "Sottocartella", "Seleziona una cartella.")
            return
        name, ok = QInputDialog.getText(self, "Sottocartella", "Nome sottocartella:")
        if not ok:
            return
        try:
            if hasattr(self.repository, "archive"):
                self.repository.archive.add_folder(name, parent_id)
            else:
                self.repository.add_archive_folder(name, parent_id)
            self._render_archive_folders()
        except ValueError as exc:
            QMessageBox.warning(self, "Cartella", str(exc))
    

    def _delete_selected_archive_folder(self) -> None:
        folder_id = self._current_archive_folder_id()
        if folder_id is None:
            QMessageBox.information(self, "Elimina cartella", "Seleziona una cartella.")
            return
        if (
            QMessageBox.question(
                self,
                "Elimina cartella",
                "Eliminare la cartella selezionata e le eventuali sottocartelle?",
            )
            != QMessageBox.StandardButton.Yes
        ):
            return
        try:
            if hasattr(self.repository, "archive"):
                self.repository.archive.delete_folder(folder_id)
            else:
                self.repository.delete_archive_folder(folder_id)
            self._render_archive_folders()
            self._render_archive_files()
        except ValueError as exc:
            QMessageBox.warning(self, "Cartella", str(exc))
    

    def _list_tags_lookup_rows(self) -> list[dict]:
        """Tag lookup: con MainController va usato `settings`; con repo legacy è sul repo."""
        if hasattr(self.repository, "settings"):
            return self.repository.settings.list_tags_lookup()
        return self.repository.list_tags_lookup()

    def _select_tag_dialog(self, current: str = "") -> str:
        dialog = QDialog(self)
        dialog.setWindowTitle("Seleziona tag")
        dialog.resize(320, 160)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(8)
    
        combo = QComboBox()
        combo.addItem("")
        for row in self._list_tags_lookup_rows():
            combo.addItem(row["label"])
        if current:
            combo.setCurrentText(current)
        layout.addWidget(QLabel("Tag"))
        layout.addWidget(combo)
    
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
    
        if dialog.exec() == QDialog.DialogCode.Accepted:
            return combo.currentText().strip()
        return current
    

    def _add_archive_files(self) -> None:
        folder_id = self._current_archive_folder_id()
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Seleziona file",
            "",
            "All files (*.*)",
        )
        if not file_paths:
            return
        tag_name = self._select_tag_dialog("")
        errors: list[str] = []
        for file_path in file_paths:
            try:
                if hasattr(self.repository, "archive"):
                    self.repository.archive.add_file(folder_id, file_path, tag_name)
                else:
                    self.repository.add_archive_file(folder_id, file_path, tag_name)
            except ValueError as exc:
                errors.append(str(exc))
        self._render_archive_files()
        if errors:
            QMessageBox.warning(self, "File", "\n".join(errors))
    

    def _delete_selected_archive_file(self) -> None:
        if not hasattr(self, "archive_files_table"):
            return
        selected_rows = sorted(
            {idx.row() for idx in self.archive_files_table.selectedIndexes()},
            reverse=True,
        )
        if not selected_rows:
            QMessageBox.information(self, "Elimina file", "Seleziona almeno un file.")
            return
        if (
            QMessageBox.question(
                self,
                "Elimina file",
                f"Eliminare {len(selected_rows)} file selezionati?",
            )
            != QMessageBox.StandardButton.Yes
        ):
            return
        for row in selected_rows:
            id_item = self.archive_files_table.item(row, 0)
            file_id = id_item.data(Qt.ItemDataRole.UserRole) if id_item else None
            if file_id is not None:
                if hasattr(self.repository, "archive"):
                    self.repository.archive.delete_file(int(file_id))
                else:
                    self.repository.delete_archive_file(int(file_id))
        self._render_archive_files()
    

    def _on_archive_file_double_click(self, row: int, column: int) -> None:
        if not hasattr(self, "archive_files_table"):
            return
        if column == 6:
            id_item = self.archive_files_table.item(row, 0)
            file_id = id_item.data(Qt.ItemDataRole.UserRole) if id_item else None
            current_tag = self.archive_files_table.item(row, 6).text()
            if file_id is not None:
                selected = self._select_tag_dialog(current_tag)
                if hasattr(self.repository, "archive"):
                    self.repository.archive.update_file_tag(int(file_id), selected)
                else:
                    self.repository.update_archive_file_tag(int(file_id), selected)
                self._render_archive_files()
            return
    
        path_item = self.archive_files_table.item(row, 5)
        if path_item is None:
            return
        path = path_item.text().strip()
        if not path:
            return
        url = QUrl.fromLocalFile(path)
        QDesktopServices.openUrl(url)
    

    def _create_archive_link(self) -> None:
        tag_options = [row["label"] for row in self._list_tags_lookup_rows()]
        dialog = LinkDialog(tag_options, parent=self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        name, url, tag = dialog.values()
        try:
            folder_id = self._current_archive_folder_id()
            if hasattr(self.repository, "archive"):
                self.repository.archive.upsert_link(None, name, url, folder_id, tag)
            else:
                self.repository.upsert_archive_link(None, name, url, folder_id, tag)
            self._render_archive_links()
        except ValueError as exc:
            QMessageBox.warning(self, "Link", str(exc))
    

    def _edit_selected_link(self) -> None:
        row = self._selected_link_row()
        if row is None:
            QMessageBox.information(self, "Link", "Seleziona un link.")
            return
        tag_options = [row["label"] for row in self._list_tags_lookup_rows()]
        link = row.copy()
        dialog = LinkDialog(tag_options, link, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        name, url, tag = dialog.values()
        try:
            folder_id = self._current_archive_folder_id()
            if hasattr(self.repository, "archive"):
                self.repository.archive.upsert_link(int(link["id"]), name, url, folder_id, tag)
            else:
                self.repository.upsert_archive_link(int(link["id"]), name, url, folder_id, tag)
            self._render_archive_links()
        except ValueError as exc:
            QMessageBox.warning(self, "Link", str(exc))
    

    def _delete_selected_link(self) -> None:
        row = self._selected_link_row()
        if row is None:
            QMessageBox.information(self, "Link", "Seleziona un link.")
            return
        if (
            QMessageBox.question(
                self,
                "Elimina link",
                "Eliminare il link selezionato?",
            )
            != QMessageBox.StandardButton.Yes
        ):
            return
        if hasattr(self.repository, "archive"):
            self.repository.archive.delete_link(int(row["id"]))
        else:
            self.repository.delete_archive_link(int(row["id"]))
        self._render_archive_links()
    

    def _selected_link_row(self) -> dict | None:
        if not hasattr(self, "archive_links_table"):
            return None
        selected = self.archive_links_table.selectedIndexes()
        if not selected:
            return None
        row = selected[0].row()
        id_item = self.archive_links_table.item(row, 0)
        link_id = id_item.data(Qt.ItemDataRole.UserRole) if id_item else None
        if link_id is None:
            return None
        name = self.archive_links_table.item(row, 0).text()
        url = self.archive_links_table.item(row, 1).text()
        tag_name = self.archive_links_table.item(row, 2).text()
        return {"id": link_id, "name": name, "url": url, "tag_name": tag_name}
    

    def _open_selected_link(self, row: int | None = None, column: int | None = None) -> None:
        link = self._selected_link_row()
        if link is None:
            return
        url = str(link.get("url") or "").strip()
        if not url:
            return
        QDesktopServices.openUrl(QUrl(url))
    

    def _open_favorite_file(self, row: int, column: int) -> None:
        id_item = self.favorites_table.item(row, 0)
        if id_item is None:
            return
        item_id = id_item.data(Qt.ItemDataRole.UserRole)
        rows = self.repository.archive.list_files_all() if hasattr(self.repository, "archive") else self.repository.list_archive_files_all()
        target = next((r for r in rows if int(r["id"]) == int(item_id)), None)
        if target:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(target.get("path") or "")))
    

    def _open_favorite_link(self, row: int, column: int) -> None:
        id_item = self.favorites_links_table.item(row, 0)
        if id_item is None:
            return
        item_id = id_item.data(Qt.ItemDataRole.UserRole)
        rows = self.repository.archive.list_links_all() if hasattr(self.repository, "archive") else self.repository.list_archive_links_all()
        target = next((r for r in rows if int(r["id"]) == int(item_id)), None)
        if target:
            QDesktopServices.openUrl(QUrl(str(target.get("url") or "")))
    

    def _on_archive_files_menu(self, pos) -> None:
        item = self.archive_files_table.itemAt(pos)
        if item is None:
            return
        row = item.row()
        id_item = self.archive_files_table.item(row, 0)
        file_id = id_item.data(Qt.ItemDataRole.UserRole) if id_item else None
        if file_id is None:
            return
        menu = QMenu(self)
        menu.addAction("Aggiungi ai preferiti", lambda: self._add_favorite("file", int(file_id)))
        menu.addAction("Sposta in cartella...", lambda: self._move_archive_file_to_folder(int(file_id)))
        menu.exec(self.archive_files_table.mapToGlobal(pos))
    

    def _on_archive_links_menu(self, pos) -> None:
        item = self.archive_links_table.itemAt(pos)
        if item is None:
            return
        row = item.row()
        id_item = self.archive_links_table.item(row, 0)
        link_id = id_item.data(Qt.ItemDataRole.UserRole) if id_item else None
        if link_id is None:
            return
        menu = QMenu(self)
        menu.addAction("Aggiungi ai preferiti", lambda: self._add_favorite("link", int(link_id)))
        menu.exec(self.archive_links_table.mapToGlobal(pos))
    

    def _add_favorite(self, item_type: str, item_id: int) -> None:
        try:
            if hasattr(self.repository, "archive"):
                self.repository.archive.add_favorite(item_type, item_id)
            else:
                self.repository.add_archive_favorite(item_type, item_id)
            self._render_favorites()
            QMessageBox.information(self, "Preferiti", "Elemento aggiunto ai preferiti.")
        except ValueError:
            QMessageBox.information(self, "Preferiti", "Elemento gia presente nei preferiti.")
    

    def _remove_favorite(self, item_type: str, item_id: int) -> None:
        if hasattr(self.repository, "archive"):
            self.repository.archive.remove_favorite(item_type, item_id)
        else:
            self.repository.remove_archive_favorite(item_type, item_id)
        self._render_favorites()
        QMessageBox.information(self, "Preferiti", "Elemento rimosso dai preferiti.")
    

    def _on_favorites_files_menu(self, pos) -> None:
        item = self.favorites_table.itemAt(pos)
        if item is None:
            return
        row = item.row()
        id_item = self.favorites_table.item(row, 0)
        file_id = id_item.data(Qt.ItemDataRole.UserRole) if id_item else None
        if file_id is None:
            return
        menu = QMenu(self)
        menu.addAction(
            "Rimuovi dai preferiti",
            lambda: self._remove_favorite("file", int(file_id)),
        )
        menu.exec(self.favorites_table.mapToGlobal(pos))
    

    def _on_favorites_links_menu(self, pos) -> None:
        item = self.favorites_links_table.itemAt(pos)
        if item is None:
            return
        row = item.row()
        id_item = self.favorites_links_table.item(row, 0)
        link_id = id_item.data(Qt.ItemDataRole.UserRole) if id_item else None
        if link_id is None:
            return
        menu = QMenu(self)
        menu.addAction(
            "Rimuovi dai preferiti",
            lambda: self._remove_favorite("link", int(link_id)),
        )
        menu.exec(self.favorites_links_table.mapToGlobal(pos))
    

    def _clear_archive_link_filters(self) -> None:
        if hasattr(self, "archive_link_filter_name"):
            self.archive_link_filter_name.setText("")
        if hasattr(self, "archive_link_filter_tag"):
            self.archive_link_filter_tag.setCurrentText("Tutti")
    

    def _on_archive_link_double_click(self, row: int, column: int) -> None:
        if column == 2:
            id_item = self.archive_links_table.item(row, 0)
            link_id = id_item.data(Qt.ItemDataRole.UserRole) if id_item else None
            current_tag = self.archive_links_table.item(row, 2).text()
            if link_id is not None:
                selected = self._select_tag_dialog(current_tag)
                link = self._selected_link_row()
                if link:
                    if hasattr(self.repository, "archive"):
                        self.repository.archive.upsert_link(
                            int(link_id),
                            link.get("name", ""),
                            link.get("url", ""),
                            self._current_archive_folder_id(),
                            selected,
                        )
                    else:
                        self.repository.upsert_archive_link(
                            int(link_id),
                            link.get("name", ""),
                            link.get("url", ""),
                            self._current_archive_folder_id(),
                            selected,
                        )
                self._render_archive_links()
            return
        self._open_selected_link()

    def _move_archive_file_to_folder(self, file_id: int) -> None:
        destination_folder_id = self._select_archive_destination_folder_dialog()
        if destination_folder_id is None:
            return

        try:
            if hasattr(self.repository, "archive"):
                self.repository.archive.move_file(int(file_id), int(destination_folder_id))
            else:
                self.repository.move_archive_file(int(file_id), int(destination_folder_id))
            self._render_archive_files()
            QMessageBox.information(self, "Sposta file", "File spostato correttamente.")
        except ValueError as exc:
            QMessageBox.warning(self, "Sposta file", str(exc))

    def _select_archive_destination_folder_dialog(self) -> int | None:
        if hasattr(self.repository, "archive"):
            folders = self.repository.archive.list_folders()
        else:
            folders = self.repository.list_archive_folders()

        if not folders:
            QMessageBox.information(
                self,
                "Sposta file",
                "Nessuna cartella disponibile. Crea prima una cartella.",
            )
            return None

        dialog = QDialog(self)
        dialog.setWindowTitle("Seleziona cartella destinazione")
        dialog.resize(420, 460)
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
        layout.addWidget(QLabel("Scegli la cartella in cui spostare il file:"))

        tree = QTreeWidget()
        tree.setHeaderHidden(True)
        layout.addWidget(tree, 1)

        by_parent: dict[int | None, list[dict]] = {}
        for row in folders:
            parent_id = row.get("parent_id")
            parent_key = int(parent_id) if parent_id is not None else None
            by_parent.setdefault(parent_key, []).append(row)

        def _add_children(parent_item: QTreeWidgetItem | None, parent_id: int | None) -> None:
            for row in sorted(
                by_parent.get(parent_id, []),
                key=lambda x: str(x.get("name") or "").lower(),
            ):
                item = QTreeWidgetItem([str(row.get("name") or "")])
                item.setData(0, Qt.ItemDataRole.UserRole, int(row["id"]))
                if parent_item is None:
                    tree.addTopLevelItem(item)
                else:
                    parent_item.addChild(item)
                _add_children(item, int(row["id"]))

        _add_children(None, None)
        tree.expandAll()

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)

        if dialog.exec() != QDialog.DialogCode.Accepted:
            return None

        current = tree.currentItem()
        if current is None:
            QMessageBox.information(self, "Sposta file", "Seleziona una cartella.")
            return None
        target = current.data(0, Qt.ItemDataRole.UserRole)
        return int(target) if target is not None else None
