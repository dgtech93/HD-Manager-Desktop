from __future__ import annotations

from functools import partial
from html import escape

import os
import subprocess
import sys
import tempfile
from collections import defaultdict

_SUBPROCESS_FLAGS = (
    subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
)

from PyQt6.QtCore import Qt, QUrl, QDate, QEvent, QThread, pyqtSignal, QMimeData
from PyQt6.QtGui import QDesktopServices, QGuiApplication, QKeySequence, QColor, QPalette
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialogButtonBox,
    QDoubleSpinBox,
    QLineEdit,
    QDateEdit,
    QDialog,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSplitter,
    QSpinBox,
    QStackedWidget,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTableWidgetSelectionRange,
    QToolButton,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
    QStyledItemDelegate,
    QStyle,
    QStyleOptionViewItem,
)

from app.views.dialogs import CredentialDialog, ContactDialog


class _RdpConnectWorker(QThread):
    """Esegue cmdkey + mstsc in background per non bloccare l'UI."""

    error_occurred = pyqtSignal(str)

    def __init__(
        self,
        target: str,
        port: str,
        username: str,
        password: str,
        system=None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self._target = target
        self._port = port
        self._username = username
        self._password = password
        self._system = system

    def run(self) -> None:
        try:
            if self._system is not None:
                self._system.connect_rdp(
                    self._target, self._port, self._username, self._password
                )
                return
            # Fallback legacy
            rdp_target = _normalize_rdp_target_static(self._target, self._port)
            target_variants = _rdp_target_variants_static(self._target, rdp_target)
            cmdkey_targets = [f"TERMSRV/{v}" for v in target_variants] + target_variants
            for key in cmdkey_targets:
                subprocess.run(
                    ["cmdkey", f"/delete:{key}"],
                    check=False,
                    capture_output=True,
                    text=True,
                    creationflags=_SUBPROCESS_FLAGS,
                )
            for key in cmdkey_targets:
                subprocess.run(
                    [
                        "cmdkey",
                        f"/generic:{key}",
                        f"/user:{self._username}",
                        f"/pass:{self._password}",
                    ],
                    check=True,
                    capture_output=True,
                    text=True,
                    creationflags=_SUBPROCESS_FLAGS,
                )
            for target_key in target_variants:
                subprocess.run(
                    [
                        "cmdkey",
                        f"/add:{target_key}",
                        f"/user:{self._username}",
                        f"/pass:{self._password}",
                    ],
                    check=False,
                    capture_output=True,
                    text=True,
                    creationflags=_SUBPROCESS_FLAGS,
                )
            rdp_file = _build_rdp_launch_file_static(rdp_target, self._username)
            subprocess.Popen(["mstsc", rdp_file], creationflags=_SUBPROCESS_FLAGS)
        except FileNotFoundError:
            self.error_occurred.emit("Comandi di sistema RDP non trovati (cmdkey/mstsc).")
        except subprocess.CalledProcessError as exc:
            details = (exc.stderr or exc.stdout or "").strip()
            self.error_occurred.emit(
                f"Impossibile preparare la connessione RDP.\n{details or 'Errore cmdkey.'}"
            )
        except (ValueError, OSError) as exc:
            self.error_occurred.emit(str(exc))


def _normalize_rdp_target_static(target: str, port: str) -> str:
    clean_target = str(target or "").strip()
    clean_port = str(port or "").strip()
    if not clean_port:
        return clean_target
    try:
        port_num = int(clean_port)
    except ValueError:
        return clean_target
    if port_num < 1 or port_num > 65535:
        return clean_target
    if ":" in clean_target:
        return clean_target
    return f"{clean_target}:{port_num}"


def _rdp_target_variants_static(clean_target: str, rdp_target: str) -> list[str]:
    variants: list[str] = []
    seen: set[str] = set()
    for value in (clean_target, rdp_target):
        current = str(value or "").strip()
        if not current:
            continue
        for candidate in (current, current.split(":")[0]):
            key = candidate.lower()
            if candidate and key not in seen:
                seen.add(key)
                variants.append(candidate)
    return variants


def _build_rdp_launch_file_static(target: str, username: str) -> str:
    lines = [
        f"full address:s:{target}",
        f"username:s:{username}",
        "prompt for credentials:i:0",
        "promptcredentialonce:i:1",
        "enablecredsspsupport:i:1",
        "authentication level:i:2",
        "negotiate security layer:i:1",
        "redirectclipboard:i:1",
        "screen mode id:i:2",
    ]
    with tempfile.NamedTemporaryFile(
        mode="w",
        encoding="utf-8",
        suffix=".rdp",
        prefix="hdm_rdp_",
        delete=False,
    ) as tmp_file:
        tmp_file.write("\n".join(lines))
        return tmp_file.name


class _NotesTableCellDelegate(QStyledItemDelegate):
    """Delegate per la tabella note: frecce a inizio/fine testo spostano alla cella adiacente."""

    def __init__(self, table: QTableWidget, on_edit_start=None) -> None:
        super().__init__(table)
        self._table = table
        self._on_edit_start = on_edit_start

    def createEditor(self, parent, option, index):
        if callable(self._on_edit_start):
            self._on_edit_start()
        editor = QLineEdit(parent)
        editor.installEventFilter(self)
        self._editing_index = index
        return editor

    def eventFilter(self, obj, event):
        if event.type() != QEvent.Type.KeyPress:
            return super().eventFilter(obj, event)
        editor = obj
        if not isinstance(editor, QLineEdit):
            return super().eventFilter(obj, event)
        text = editor.text()
        pos = editor.cursorPosition()
        key = event.key()
        if key == Qt.Key.Key_Right and pos == len(text):
            self.commitData.emit(editor)
            self.closeEditor.emit(editor, QStyledItemDelegate.EndEditHint.EditNextItem)
            return True
        if key == Qt.Key.Key_Left and pos == 0:
            self.commitData.emit(editor)
            self.closeEditor.emit(editor, QStyledItemDelegate.EndEditHint.EditPreviousItem)
            return True
        return super().eventFilter(obj, event)


def _notes_letter_to_col(s: str) -> int:
    """A->0, B->1, ..., Z->25, AA->26, ..."""
    result = 0
    for c in s.upper():
        if "A" <= c <= "Z":
            result = result * 26 + (ord(c) - ord("A") + 1)
    return result - 1 if result else 0


def _notes_parse_cell_ref(ref: str) -> tuple[int, int] | None:
    """Parse 'A1' -> (0, 0), 'B2' -> (1, 1). Returns (row, col) or None."""
    ref = ref.strip().upper()
    if not ref:
        return None
    i = 0
    while i < len(ref) and ref[i].isalpha():
        i += 1
    if i == 0 or i >= len(ref):
        return None
    letters, digits = ref[:i], ref[i:]
    try:
        row = int(digits) - 1
    except ValueError:
        return None
    col = _notes_letter_to_col(letters)
    return (row, col)


def _notes_parse_range(ref: str, table_rows: int, table_cols: int) -> list[tuple[int, int]]:
    """Parse 'A1:B3' or 'A1,B2,C3' into list of (row,col)."""
    cells: list[tuple[int, int]] = []
    ref = ref.strip()
    if ":" in ref:
        parts = ref.split(":", 1)
        tl = _notes_parse_cell_ref(parts[0].strip())
        br = _notes_parse_cell_ref(parts[1].strip())
        if tl is None or br is None:
            return []
        r1, c1 = tl
        r2, c2 = br
        for r in range(min(r1, r2), max(r1, r2) + 1):
            for c in range(min(c1, c2), max(c1, c2) + 1):
                if 0 <= r < table_rows and 0 <= c < table_cols:
                    cells.append((r, c))
    else:
        for part in ref.split(","):
            cell = _notes_parse_cell_ref(part.strip())
            if cell and 0 <= cell[0] < table_rows and 0 <= cell[1] < table_cols:
                cells.append(cell)
    return cells


class _NotesFormulasDialog(QDialog):
    OPERATIONS = [
        ("Somma", "sum"),
        ("Sottrazione", "sub"),
        ("Moltiplicazione", "mul"),
        ("Divisione", "div"),
        ("Percentuale (A % di B)", "pct"),
        ("Radice quadrata", "sqrt"),
        ("Radice cubica", "cbrt"),
        ("Radice n-esima", "nroot"),
        ("Concatenazione righe", "concat_row"),
        ("Concatenazione colonne", "concat_col"),
        ("Media", "avg"),
    ]

    def __init__(self, table: QTableWidget, parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Formule")
        self.resize(420, 380)
        self._table = table
        layout = QVBoxLayout(self)

        layout.addWidget(QLabel("Operazione:"))
        self.op_combo = QComboBox()
        for label, _key in self.OPERATIONS:
            self.op_combo.addItem(label)
        self.op_combo.currentIndexChanged.connect(self._on_op_changed)
        layout.addWidget(self.op_combo)

        src_h = QHBoxLayout()
        src_h.addWidget(QLabel("Celle sorgente:"))
        self.src_edit = QLineEdit()
        self.src_edit.setPlaceholderText("es. A1:B3 (intervallo) o A1,B2,C3 (celle sparse)")
        src_h.addWidget(self.src_edit, 1)
        use_sel_btn = QPushButton("Usa selezione")
        use_sel_btn.clicked.connect(self._fill_from_selection)
        src_h.addWidget(use_sel_btn)
        layout.addLayout(src_h)

        self.n_root_row = QHBoxLayout()
        self.n_root_row.addWidget(QLabel("n (per radice n-esima):"))
        self.n_spin = QSpinBox()
        self.n_spin.setMinimum(2)
        self.n_spin.setMaximum(100)
        self.n_spin.setValue(2)
        self.n_root_row.addWidget(self.n_spin)
        self.n_root_row.addStretch()
        self._n_root_widget = QWidget()
        self._n_root_widget.setLayout(self.n_root_row)
        layout.addWidget(self._n_root_widget)
        self._n_root_widget.setVisible(False)

        self.sep_row = QHBoxLayout()
        self.sep_row.addWidget(QLabel("Separatore:"))
        self.sep_edit = QLineEdit()
        self.sep_edit.setPlaceholderText("es. spazio, virgola, |")
        self.sep_edit.setMaxLength(5)
        self.sep_row.addWidget(self.sep_edit, 1)
        self._sep_widget = QWidget()
        self._sep_widget.setLayout(self.sep_row)
        layout.addWidget(self._sep_widget)
        self._sep_widget.setVisible(False)

        res_h = QHBoxLayout()
        res_h.addWidget(QLabel("Cella risultato:"))
        self.res_edit = QLineEdit()
        self.res_edit.setPlaceholderText("es. C1")
        res_h.addWidget(self.res_edit, 1)
        layout.addLayout(res_h)

        self.apply_all_cb = QCheckBox("Applica a tutte le righe (come in Excel)")
        layout.addWidget(self.apply_all_cb)

        layout.addStretch()
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self._on_op_changed(0)
        self._fill_from_selection()

    def _on_op_changed(self, index: int) -> None:
        _, key = self.OPERATIONS[index]
        self._n_root_widget.setVisible(key == "nroot")
        self._sep_widget.setVisible(key in ("concat_row", "concat_col"))

    def _fill_from_selection(self) -> None:
        """Riempie le celle sorgente dalla selezione (supporta celle non contigue)."""
        ranges = self._table.selectedRanges()
        if not ranges:
            return
        parts: list[str] = []
        for rng in ranges:
            for r in range(rng.topRow(), rng.bottomRow() + 1):
                for c in range(rng.leftColumn(), rng.rightColumn() + 1):
                    parts.append(f"{_notes_col_index_to_letter(c)}{r + 1}")
        if not parts:
            return
        if len(parts) == 1:
            self.src_edit.setText(parts[0])
        elif len(ranges) == 1 and (ranges[0].rowCount() > 1 or ranges[0].columnCount() > 1):
            rng = ranges[0]
            self.src_edit.setText(
                f"{_notes_col_index_to_letter(rng.leftColumn())}{rng.topRow() + 1}:"
                f"{_notes_col_index_to_letter(rng.rightColumn())}{rng.bottomRow() + 1}"
            )
        else:
            self.src_edit.setText(",".join(parts))
        if not self.res_edit.text().strip():
            rng = ranges[0]
            res_col = min(rng.rightColumn() + 1, self._table.columnCount() - 1)
            self.res_edit.setText(f"{_notes_col_index_to_letter(res_col)}{rng.topRow() + 1}")

    def _get_op_key(self) -> str:
        return self.OPERATIONS[self.op_combo.currentIndex()][1]

    def _get_src_cells(self) -> list[tuple[int, int]]:
        return _notes_parse_range(
            self.src_edit.text(), self._table.rowCount(), self._table.columnCount()
        )

    def _get_res_cell(self) -> tuple[int, int] | None:
        return _notes_parse_cell_ref(self.res_edit.text().strip())

    def _get_cell_value(self, r: int, c: int) -> str:
        item = self._table.item(r, c)
        return (item.text() or "").strip()

    def _get_cell_numbers(self, cells: list[tuple[int, int]]) -> list[float]:
        vals: list[float] = []
        for r, c in cells:
            txt = self._get_cell_value(r, c)
            if txt:
                try:
                    vals.append(float(txt.replace(",", ".")))
                except ValueError:
                    pass
        return vals

    def _set_result(self, row: int, col: int, value: str) -> None:
        if row >= self._table.rowCount() or col >= self._table.columnCount():
            return
        item = self._table.item(row, col)
        if item:
            item.setText(value)
        else:
            self._table.setItem(row, col, QTableWidgetItem(value))

    def _compute(self, src_cells: list[tuple[int, int]], res_row: int, res_col: int) -> bool:
        key = self._get_op_key()
        nums = self._get_cell_numbers(src_cells)
        texts = [self._get_cell_value(r, c) for r, c in src_cells]

        import math
        if key == "sum":
            if not nums:
                return False
            self._set_result(res_row, res_col, str(sum(nums)))
        elif key == "sub":
            if len(nums) < 2:
                return False
            r = nums[0] - sum(nums[1:])
            self._set_result(res_row, res_col, str(r))
        elif key == "mul":
            if not nums:
                return False
            p = 1
            for n in nums:
                p *= n
            self._set_result(res_row, res_col, str(p))
        elif key == "div":
            if len(nums) < 2 or nums[1] == 0:
                return False
            self._set_result(res_row, res_col, str(nums[0] / nums[1]))
        elif key == "pct":
            if len(nums) < 2 or nums[1] == 0:
                return False
            self._set_result(res_row, res_col, str(round(100 * nums[0] / nums[1], 2)))
        elif key == "sqrt":
            if len(nums) < 1 or nums[0] < 0:
                return False
            self._set_result(res_row, res_col, str(round(math.sqrt(nums[0]), 6)))
        elif key == "cbrt":
            if len(nums) < 1:
                return False
            self._set_result(res_row, res_col, str(round(nums[0] ** (1 / 3), 6)))
        elif key == "nroot":
            n = self.n_spin.value()
            if len(nums) < 1 or n <= 0:
                return False
            if nums[0] < 0 and n % 2 == 0:
                return False
            self._set_result(res_row, res_col, str(round(nums[0] ** (1 / n), 6)))
        elif key == "avg":
            if not nums:
                return False
            self._set_result(res_row, res_col, str(round(sum(nums) / len(nums), 6)))
        elif key == "concat_row":
            sep = self.sep_edit.text() or " "
            ordered = sorted(src_cells, key=lambda x: (x[0], x[1]))
            texts = [self._get_cell_value(r, c) for r, c in ordered]
            self._set_result(res_row, res_col, sep.join(texts))
        elif key == "concat_col":
            sep = self.sep_edit.text() or " "
            ordered = sorted(src_cells, key=lambda x: (x[1], x[0]))
            texts = [self._get_cell_value(r, c) for r, c in ordered]
            self._set_result(res_row, res_col, sep.join(texts))
        else:
            return False
        return True

    def accept(self) -> None:
        src = self._get_src_cells()
        res = self._get_res_cell()
        if not src:
            QMessageBox.warning(self, "Formule", "Specifica almeno una cella sorgente valida.")
            return
        if res is None:
            QMessageBox.warning(self, "Formule", "Specifica una cella risultato valida (es. C1).")
            return
        rows, cols = self._table.rowCount(), self._table.columnCount()
        if res[0] >= rows or res[1] >= cols:
            QMessageBox.warning(self, "Formule", "La cella risultato è fuori dai limiti della tabella.")
            return

        if self.apply_all_cb.isChecked():
            min_src_c = min(c for _, c in src)
            max_src_c = max(c for _, c in src)
            res_r, res_c = res
            for row in range(rows):
                src_cells = [(row, c) for c in range(min_src_c, max_src_c + 1)]
                if not self._compute(src_cells, row, res_c):
                    pass
        else:
            if not self._compute(src, res[0], res[1]):
                QMessageBox.warning(self, "Formule", "Impossibile calcolare con le celle specificate.")
                return

        super().accept()


def _notes_col_index_to_letter(idx: int) -> str:
    """0->A, 1->B, ..., 26->AA, ..."""
    result = ""
    idx += 1
    while idx > 0:
        idx -= 1
        result = chr(65 + idx % 26) + result
        idx //= 26
    return result or "A"


class _ClientsTreeDelegate(QStyledItemDelegate):
    def __init__(self, tree: QTreeWidget) -> None:
        super().__init__(tree)
        self._tree = tree
        self._client_bg = QColor("#2563eb")
        self._client_fg = QColor("#ffffff")
        self._product_bg = QColor("#93c5fd")
        self._product_fg = QColor("#0f172a")

    def paint(self, painter, option, index) -> None:  # type: ignore[override]
        item = self._tree.itemFromIndex(index)
        current = self._tree.currentItem()

        is_current = current is not None and item is current
        current_type = (
            current.data(0, Qt.ItemDataRole.UserRole) if current is not None else None
        )
        item_type = item.data(0, Qt.ItemDataRole.UserRole) if item is not None else None

        highlight_client = False
        highlight_product = False
        if is_current and item_type == "product":
            highlight_product = True
        elif is_current and item_type != "product":
            highlight_client = True
        elif current is not None and current_type == "product" and item is current.parent():
            highlight_client = True

        # Remove default selection state so our colors are used consistently.
        opt = QStyleOptionViewItem(option)
        opt.state &= ~QStyle.StateFlag.State_Selected

        if highlight_client:
            painter.save()
            painter.fillRect(opt.rect, self._client_bg)
            pal = QPalette(opt.palette)
            pal.setColor(QPalette.ColorRole.Text, self._client_fg)
            pal.setColor(QPalette.ColorRole.WindowText, self._client_fg)
            opt.palette = pal
            super().paint(painter, opt, index)
            painter.restore()
            return

        if highlight_product:
            painter.save()
            painter.fillRect(opt.rect, self._product_bg)
            pal = QPalette(opt.palette)
            pal.setColor(QPalette.ColorRole.Text, self._product_fg)
            pal.setColor(QPalette.ColorRole.WindowText, self._product_fg)
            opt.palette = pal
            super().paint(painter, opt, index)
            painter.restore()
            return

        super().paint(painter, opt, index)


class ClientsMixin:
    # Clienti, accessi, rubrica
    def _build_client_workspace_page(self, title_text: str) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
    
        title = QLabel(title_text)
        title.setObjectName("sectionTitle")
        layout.addWidget(title)
    
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
    
        left = QWidget()
        left_layout = QVBoxLayout(left)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)
        left_title = QLabel("Lista Clienti")
        left_title.setObjectName("subSectionTitle")
        left_layout.addWidget(left_title)
    
        self.clients_filter = QLineEdit()
        self.clients_filter.setObjectName("clientsFilter")
        self.clients_filter.setPlaceholderText("Filtra clienti...")
        self.clients_filter.textChanged.connect(self._apply_clients_filter)
        left_layout.addWidget(self.clients_filter)
    
        self.clients_tree = QTreeWidget()
        self.clients_tree.setObjectName("clientsTree")
        self.clients_tree.setHeaderHidden(True)
        self.clients_tree.setItemDelegate(_ClientsTreeDelegate(self.clients_tree))
        self.clients_tree.currentItemChanged.connect(self._on_archive_client_selected)
        self.clients_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.clients_tree.customContextMenuRequested.connect(self._on_clients_tree_menu)
        left_layout.addWidget(self.clients_tree, 1)
        splitter.addWidget(left)
    
        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(8)
    
        self.client_detail_title = QLabel("Seleziona un cliente")
        self.client_detail_title.setObjectName("subSectionTitle")
        right_layout.addWidget(self.client_detail_title)
    
        self.client_detail_meta = QLabel("")
        self.client_detail_meta.setObjectName("subText")
        right_layout.addWidget(self.client_detail_meta)
    
        self.client_tabs = QTabWidget()
        self.client_tabs.addTab(self._build_client_info_tab(), "Info Cliente")
        self.client_tabs.addTab(self._build_client_access_tab(), "Accessi")
        self.client_tabs.addTab(self._build_client_archive_tab(), "Archivio Cliente")
        self.client_tabs.addTab(self._build_client_contacts_tab(), "Rubrica")
        self.client_tabs.addTab(self._build_client_notes_tab(), "Note")
        right_layout.addWidget(self.client_tabs, 1)
        splitter.addWidget(right)
    
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([330, 820])
        layout.addWidget(splitter, 1)
        return page
    

    def _build_client_info_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
    
        data_title = QLabel("Dati Cliente")
        data_title.setObjectName("subSectionTitle")
        layout.addWidget(data_title)
    
        data_grid = QGridLayout()
        data_grid.setHorizontalSpacing(16)
        data_grid.setVerticalSpacing(10)
        self.info_name_value = self._add_info_row(data_grid, 0, "Nome")
        self.info_location_value = self._add_info_row(data_grid, 1, "Localita")
        self.info_link_value = self._add_info_row(data_grid, 2, "Link")
        self.info_link_value.setTextFormat(Qt.TextFormat.RichText)
        self.info_link_value.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        self.info_link_value.setOpenExternalLinks(True)
        layout.addLayout(data_grid)
    
        roles_title = QLabel("Risorse divise per Ruolo")
        roles_title.setObjectName("subSectionTitle")
        layout.addWidget(roles_title)
    
        self.info_roles_container = QWidget()
        self.info_roles_layout = QVBoxLayout(self.info_roles_container)
        self.info_roles_layout.setContentsMargins(0, 0, 0, 0)
        self.info_roles_layout.setSpacing(8)
        empty_roles = QLabel("Nessuna risorsa collegata.")
        empty_roles.setObjectName("subText")
        self.info_roles_layout.addWidget(empty_roles)
        self.info_roles_layout.addStretch()
    
        roles_scroll = QScrollArea()
        roles_scroll.setWidgetResizable(True)
        roles_scroll.setFrameShape(QFrame.Shape.NoFrame)
        roles_scroll.setWidget(self.info_roles_container)
        layout.addWidget(roles_scroll, 1)
    
        products_title = QLabel("Prodotti collegati")
        products_title.setObjectName("subSectionTitle")
        layout.addWidget(products_title)
    
        self.info_products_area = QLabel(
            "Area dedicata ai prodotti collegati. Le regole saranno implementate nel prossimo step."
        )
        self.info_products_area.setObjectName("hintCard")
        self.info_products_area.setWordWrap(True)
        layout.addWidget(self.info_products_area)
        return page
    

    def _add_info_row(self, layout: QGridLayout, row: int, label: str) -> QLabel:
        key_lbl = QLabel(label)
        key_lbl.setObjectName("infoKey")
        value_lbl = QLabel("-")
        value_lbl.setObjectName("infoValue")
        layout.addWidget(key_lbl, row, 0)
        layout.addWidget(value_lbl, row, 1)
        return value_lbl
    
    @staticmethod

    def _add_vpn_card_item(
        layout: QGridLayout, row: int, col: int, label: str, col_span: int = 1
    ) -> QLineEdit:
        container = QWidget()
        vbox = QVBoxLayout(container)
        vbox.setContentsMargins(0, 0, 0, 0)
        vbox.setSpacing(4)
    
        key_lbl = QLabel(label)
        key_lbl.setObjectName("infoKey")
        value_edit = QLineEdit("-")
        value_edit.setObjectName("copyField")
        value_edit.setReadOnly(True)
        value_edit.setCursorPosition(0)
        vbox.addWidget(key_lbl)
        vbox.addWidget(value_edit)
        layout.addWidget(container, row, col, 1, col_span)
        return value_edit
    

    def _build_client_access_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
    
        # Collapsible VPN section to free vertical space for credentials
        vpn_header = QHBoxLayout()
        vpn_header.setContentsMargins(0, 0, 0, 0)
        vpn_header.setSpacing(6)
    
        self.vpn_toggle_btn = QToolButton()
        self.vpn_toggle_btn.setText("VPN associata")
        self.vpn_toggle_btn.setCheckable(True)
        self.vpn_toggle_btn.setChecked(False)  # collapsed by default
        self.vpn_toggle_btn.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.vpn_toggle_btn.setArrowType(Qt.ArrowType.RightArrow)
        self.vpn_toggle_btn.setObjectName("vpnToggle")
        self.vpn_toggle_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.vpn_toggle_btn.setAutoRaise(True)
        vpn_header.addWidget(self.vpn_toggle_btn, alignment=Qt.AlignmentFlag.AlignLeft)
        vpn_header.addStretch()
        layout.addLayout(vpn_header)
    
        self.vpn_section = QWidget()
        vpn_section_layout = QVBoxLayout(self.vpn_section)
        vpn_section_layout.setContentsMargins(0, 0, 0, 0)
        vpn_section_layout.setSpacing(8)
    
        self.vpn_card = QFrame()
        self.vpn_card.setObjectName("vpnCard")
        vpn_card_layout = QGridLayout(self.vpn_card)
        vpn_card_layout.setContentsMargins(12, 10, 12, 10)
        vpn_card_layout.setHorizontalSpacing(14)
        vpn_card_layout.setVerticalSpacing(8)
    
        self.vpn_name_field = self._add_vpn_card_item(vpn_card_layout, 0, 0, "Nome Connessione")
        self.vpn_server_field = self._add_vpn_card_item(vpn_card_layout, 0, 1, "Server")
        self.vpn_type_field = self._add_vpn_card_item(vpn_card_layout, 0, 2, "Tipo VPN")
        self.vpn_access_info_field = self._add_vpn_card_item(vpn_card_layout, 1, 0, "Info Accesso")
        self.vpn_user_field = self._add_vpn_card_item(vpn_card_layout, 1, 1, "Utente")
        self.vpn_password_field = self._add_vpn_card_item(vpn_card_layout, 1, 2, "Password")
        self.vpn_path_field = self._add_vpn_card_item(vpn_card_layout, 2, 0, "Percorso VPN", col_span=3)
        vpn_section_layout.addWidget(self.vpn_card)
    
        self.vpn_connect_btn = QPushButton("Connetti")
        self.vpn_connect_btn.setObjectName("accessActionButton")
        self.vpn_connect_btn.setEnabled(False)
        self.vpn_connect_btn.clicked.connect(self._connect_vpn_action)
        self.vpn_connect_btn.setFixedWidth(140)
        vpn_section_layout.addWidget(self.vpn_connect_btn, alignment=Qt.AlignmentFlag.AlignLeft)
    
        self.vpn_section.setVisible(False)
        layout.addWidget(self.vpn_section)
    
        def _toggle_vpn_section(expanded: bool) -> None:
            self.vpn_section.setVisible(expanded)
            self.vpn_toggle_btn.setArrowType(
                Qt.ArrowType.DownArrow if expanded else Qt.ArrowType.RightArrow
            )
    
        self.vpn_toggle_btn.toggled.connect(_toggle_vpn_section)
    
        cred_title = QLabel("Credenziali Prodotto")
        cred_title.setObjectName("subSectionTitle")
        layout.addWidget(cred_title)
    
        self.access_product_label = QLabel("Seleziona un prodotto nella lista clienti.")
        self.access_product_label.setObjectName("subText")
        self.access_product_label.setWordWrap(True)
        layout.addWidget(self.access_product_label)
    
        action_bar = QHBoxLayout()
        action_bar.setSpacing(8)
        self.connect_rdp_ip_btn = QPushButton("Connetti RDP (IP)")
        self.connect_rdp_ip_btn.setObjectName("accessActionButton")
        self.connect_rdp_ip_btn.setEnabled(False)
        self.connect_rdp_ip_btn.clicked.connect(partial(self._run_selected_credential_action, "ip"))
        action_bar.addWidget(self.connect_rdp_ip_btn)
    
        self.connect_rdp_host_btn = QPushButton("Connetti RDP (HOST)")
        self.connect_rdp_host_btn.setObjectName("accessActionButton")
        self.connect_rdp_host_btn.setEnabled(False)
        self.connect_rdp_host_btn.clicked.connect(
            partial(self._run_selected_credential_action, "host")
        )
        action_bar.addWidget(self.connect_rdp_host_btn)
    
        self.open_url_btn = QPushButton("Apri Link")
        self.open_url_btn.setObjectName("accessActionButton")
        self.open_url_btn.setEnabled(False)
        self.open_url_btn.clicked.connect(partial(self._run_selected_credential_action, "url"))
        action_bar.addWidget(self.open_url_btn)
    
        self.connect_rdp_conf_btn = QPushButton("Connetti RDP conf")
        self.connect_rdp_conf_btn.setObjectName("accessActionButton")
        self.connect_rdp_conf_btn.setEnabled(False)
        self.connect_rdp_conf_btn.clicked.connect(
            partial(self._run_selected_credential_action, "rdp_path")
        )
        action_bar.addWidget(self.connect_rdp_conf_btn)
    
        action_bar.addStretch()
        layout.addLayout(action_bar)
    
        self.access_endpoint_columns: list[tuple[str, str]] = [
            ("ip", "IP"),
            ("host", "HOST"),
            ("url", "URL"),
            ("rdp_path", "RDP PreConfigurata"),
        ]
        self.access_credentials_columns: list[str] = []
    
        self.access_credentials_table = QTableWidget(0, 0)
        self.access_credentials_table.verticalHeader().setVisible(False)
        self.access_credentials_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.access_credentials_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.access_credentials_table.setAlternatingRowColors(True)
        self.access_credentials_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.access_credentials_table.customContextMenuRequested.connect(
            self._on_access_credentials_menu
        )
        self.access_credentials_table.itemSelectionChanged.connect(
            self._update_access_credentials_actions
        )
        self._configure_access_credentials_columns([])
        layout.addWidget(self.access_credentials_table, 1)
    
        actions = QHBoxLayout()
        actions.addStretch()
        self.edit_credential_btn = QPushButton("Modifica")
        self.edit_credential_btn.setObjectName("primaryActionButton")
        self.edit_credential_btn.setEnabled(False)
        self.edit_credential_btn.clicked.connect(self._edit_selected_credential)
        actions.addWidget(self.edit_credential_btn)
    
        self.delete_credential_btn = QPushButton("Elimina")
        self.delete_credential_btn.setObjectName("dangerActionButton")
        self.delete_credential_btn.setEnabled(False)
        self.delete_credential_btn.clicked.connect(self._delete_selected_credential)
        actions.addWidget(self.delete_credential_btn)
        layout.addLayout(actions)
        return page
    

    def _build_client_archive_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
    
        title = QLabel("Archivio Cliente")
        title.setObjectName("subSectionTitle")
        layout.addWidget(title)
    
        stats_grid = QGridLayout()
        stats_grid.setHorizontalSpacing(16)
        stats_grid.setVerticalSpacing(8)
    
        self.client_files_count = self._add_info_row(stats_grid, 0, "Totale File")
        self.client_links_count = self._add_info_row(stats_grid, 1, "Totale Link")
        layout.addLayout(stats_grid)
    
        docs_title = QLabel("Documenti cliente (Tag)")
        docs_title.setObjectName("subSectionTitle")
        layout.addWidget(docs_title)
    
        self.client_docs_files_table = QTableWidget(0, 2)
        self.client_docs_files_table.setHorizontalHeaderLabels(["Nome", "Percorso"])
        self.client_docs_files_table.verticalHeader().setVisible(False)
        self.client_docs_files_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.client_docs_files_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.client_docs_files_table.setAlternatingRowColors(True)
        self.client_docs_files_table.setColumnWidth(0, 260)
        self.client_docs_files_table.horizontalHeader().setStretchLastSection(True)
        self.client_docs_files_table.cellDoubleClicked.connect(
            self._open_client_doc_file
        )
        self.client_docs_files_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.client_docs_files_table.customContextMenuRequested.connect(
            self._on_client_docs_files_menu
        )
        layout.addWidget(self.client_docs_files_table, 1)
    
        self.client_docs_links_table = QTableWidget(0, 2)
        self.client_docs_links_table.setHorizontalHeaderLabels(["Nome", "URL"])
        self.client_docs_links_table.verticalHeader().setVisible(False)
        self.client_docs_links_table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectRows
        )
        self.client_docs_links_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.client_docs_links_table.setAlternatingRowColors(True)
        self.client_docs_links_table.setColumnWidth(0, 260)
        self.client_docs_links_table.horizontalHeader().setStretchLastSection(True)
        self.client_docs_links_table.cellDoubleClicked.connect(
            self._open_client_doc_link
        )
        self.client_docs_links_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.client_docs_links_table.customContextMenuRequested.connect(
            self._on_client_docs_links_menu
        )
        layout.addWidget(self.client_docs_links_table, 1)
        return page
    

    def _build_client_contacts_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)
    
        title = QLabel("Rubrica Cliente")
        title.setObjectName("subSectionTitle")
        layout.addWidget(title)
    
        actions = QHBoxLayout()
        actions.addStretch()
        self.contacts_add_btn = QPushButton("Nuovo contatto")
        self.contacts_edit_btn = QPushButton("Modifica")
        self.contacts_delete_btn = QPushButton("Elimina")
        self.contacts_add_btn.setObjectName("primaryActionButton")
        self.contacts_delete_btn.setObjectName("dangerActionButton")
        self.contacts_edit_btn.setEnabled(False)
        self.contacts_delete_btn.setEnabled(False)
        actions.addWidget(self.contacts_add_btn)
        actions.addWidget(self.contacts_edit_btn)
        actions.addWidget(self.contacts_delete_btn)
        layout.addLayout(actions)
    
        self.contacts_table = QTableWidget(0, 6)
        self.contacts_table.setHorizontalHeaderLabels(
            ["Nome", "Telefono", "Cellulare", "Mail", "Ruolo", "Note"]
        )
        self.contacts_table.verticalHeader().setVisible(False)
        self.contacts_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.contacts_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.contacts_table.setAlternatingRowColors(True)
        self.contacts_table.setColumnWidth(0, 200)
        self.contacts_table.setColumnWidth(1, 140)
        self.contacts_table.setColumnWidth(2, 140)
        self.contacts_table.setColumnWidth(3, 200)
        self.contacts_table.setColumnWidth(4, 160)
        self.contacts_table.horizontalHeader().setStretchLastSection(True)
        self.contacts_table.itemSelectionChanged.connect(
            self._update_contacts_actions
        )
        self.contacts_table.cellClicked.connect(self._handle_contact_cell_click)
        self.contacts_table.cellDoubleClicked.connect(self._handle_contact_cell_double_click)
        self.contacts_table.installEventFilter(self)
        layout.addWidget(self.contacts_table, 1)
    
        self.contacts_add_btn.clicked.connect(self._add_contact)
        self.contacts_edit_btn.clicked.connect(self._edit_contact)
        self.contacts_delete_btn.clicked.connect(self._delete_contact)
        return page

    def _build_client_notes_tab(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(8)

        title = QLabel("Note Cliente")
        title.setObjectName("subSectionTitle")
        layout.addWidget(title)

        main_h = QHBoxLayout()
        main_h.setSpacing(16)

        btn_column = QVBoxLayout()
        btn_column.setSpacing(6)
        btn_column.addWidget(QLabel("Lista note:"))
        self.notes_list_combo = QComboBox()
        self.notes_list_combo.setMinimumWidth(180)
        self.notes_list_combo.currentIndexChanged.connect(self._on_notes_list_selected)
        btn_column.addWidget(self.notes_list_combo)

        self.notes_save_btn = QPushButton("Salva nella lista")
        self.notes_save_btn.setObjectName("primaryActionButton")
        self.notes_save_btn.clicked.connect(self._save_note_to_list)
        btn_column.addWidget(self.notes_save_btn)

        self.notes_export_btn = QPushButton("Salva in file e archivio")
        self.notes_export_btn.setObjectName("archiveActionButton")
        self.notes_export_btn.clicked.connect(self._export_note_to_file_and_archive)
        btn_column.addWidget(self.notes_export_btn)

        self.notes_add_row_btn = QPushButton("Aggiungi riga")
        self.notes_add_row_btn.clicked.connect(self._notes_add_row)
        btn_column.addWidget(self.notes_add_row_btn)

        self.notes_clear_btn = QPushButton("Pulisci foglio")
        self.notes_clear_btn.clicked.connect(self._notes_clear_sheet)
        btn_column.addWidget(self.notes_clear_btn)

        self.notes_formulas_btn = QPushButton("Formule")
        self.notes_formulas_btn.clicked.connect(self._open_notes_formulas_dialog)
        btn_column.addWidget(self.notes_formulas_btn)

        self.notes_undo_btn = QPushButton("Annulla (Ctrl+Z)")
        self.notes_undo_btn.clicked.connect(self._notes_undo)
        self.notes_undo_btn.setEnabled(False)
        btn_column.addWidget(self.notes_undo_btn)
        self.notes_redo_btn = QPushButton("Ripristina (Ctrl+Y)")
        self.notes_redo_btn.clicked.connect(self._notes_redo)
        self.notes_redo_btn.setEnabled(False)
        btn_column.addWidget(self.notes_redo_btn)

        self.notes_new_btn = QPushButton("Nuova nota")
        self.notes_new_btn.clicked.connect(self._new_note)
        btn_column.addWidget(self.notes_new_btn)

        self.notes_delete_btn = QPushButton("Elimina")
        self.notes_delete_btn.setObjectName("dangerActionButton")
        self.notes_delete_btn.clicked.connect(self._delete_selected_note)
        btn_column.addWidget(self.notes_delete_btn)

        btn_column.addStretch()
        btn_widget = QWidget()
        btn_widget.setMaximumWidth(220)
        btn_widget.setLayout(btn_column)
        main_h.addWidget(btn_widget)

        content_column = QVBoxLayout()
        content_column.setSpacing(6)
        self.notes_title_edit = QLineEdit()
        self.notes_title_edit.setPlaceholderText("Titolo nota...")
        content_column.addWidget(self.notes_title_edit)

        self.notes_tabs = QTabWidget()
        self.notes_text_edit = QPlainTextEdit()
        self.notes_text_edit.setPlaceholderText("Scrivi qui il testo della nota...")
        self.notes_tabs.addTab(self.notes_text_edit, "Testo")

        self.notes_table = QTableWidget(0, 4)
        self._notes_undo_stack: list[str] = []
        self._notes_redo_stack: list[str] = []
        self._notes_restoring = False
        self._notes_max_undo = 50
        self.notes_table.setItemDelegate(_NotesTableCellDelegate(self.notes_table, self._notes_snapshot))
        self._notes_apply_excel_headers()
        self.notes_table.setAlternatingRowColors(True)
        self.notes_table.horizontalHeader().setStretchLastSection(True)
        self.notes_table.setEditTriggers(
            QTableWidget.EditTrigger.DoubleClicked
            | QTableWidget.EditTrigger.AnyKeyPressed
        )
        self.notes_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectItems)
        self.notes_table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.notes_table.setTabKeyNavigation(True)
        self.notes_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.notes_table.customContextMenuRequested.connect(self._on_notes_table_context_menu)
        self.notes_table.currentCellChanged.connect(self._on_notes_cell_changed)
        tabella_widget = QWidget()
        tabella_layout = QVBoxLayout(tabella_widget)
        tabella_layout.setContentsMargins(0, 0, 0, 0)
        cell_section = QFrame()
        cell_section.setObjectName("notesCellPreview")
        cell_section_layout = QVBoxLayout(cell_section)
        cell_section_layout.setContentsMargins(0, 0, 0, 8)
        cell_preview_label = QLabel("Contenuto cella")
        cell_preview_label.setObjectName("subSectionTitle")
        cell_section_layout.addWidget(cell_preview_label)
        self.notes_cell_preview = QPlainTextEdit()
        self.notes_cell_preview.setPlaceholderText("Seleziona una cella per visualizzare o modificare il contenuto...")
        self.notes_cell_preview.setMaximumHeight(52)
        self.notes_cell_preview.setMinimumHeight(40)
        self.notes_cell_preview.textChanged.connect(self._on_notes_cell_preview_changed)
        cell_section_layout.addWidget(self.notes_cell_preview)
        cell_preview_actions = QHBoxLayout()
        self.notes_cell_clear_btn = QPushButton("Elimina contenuto")
        self.notes_cell_clear_btn.setObjectName("dangerActionButton")
        self.notes_cell_clear_btn.clicked.connect(self._notes_clear_selected_cells)
        cell_preview_actions.addWidget(self.notes_cell_clear_btn)
        cell_preview_actions.addStretch()
        cell_section_layout.addLayout(cell_preview_actions)
        tabella_layout.addWidget(cell_section)
        tabella_layout.addWidget(self.notes_table, 1)
        self.notes_tabs.addTab(tabella_widget, "Tabella")
        self._notes_cell_preview_block = False

        content_column.addWidget(self.notes_tabs, 1)
        content_widget = QWidget()
        content_widget.setLayout(content_column)
        main_h.addWidget(content_widget, 1)
        layout.addLayout(main_h)

        self.notes_table.installEventFilter(self)
        self.notes_cell_preview.installEventFilter(self)
        self._notes_add_row()
        self._current_note_id: int | None = None
        return page

    def _notes_apply_excel_headers(self) -> None:
        cols = self.notes_table.columnCount()
        self.notes_table.setHorizontalHeaderLabels(
            [_notes_col_index_to_letter(c) for c in range(cols)]
        )
        rows = self.notes_table.rowCount()
        self.notes_table.verticalHeader().setVisible(True)
        self.notes_table.setVerticalHeaderLabels([str(r + 1) for r in range(rows)])

    def _notes_get_state(self) -> str:
        """Serializza lo stato corrente della tabella."""
        import json
        rows = []
        for r in range(self.notes_table.rowCount()):
            row = []
            for c in range(self.notes_table.columnCount()):
                item = self.notes_table.item(r, c)
                row.append(item.text() if item else "")
            rows.append(row)
        return json.dumps(rows, ensure_ascii=False)

    def _notes_set_state(self, state: str) -> None:
        """Ripristina la tabella dallo stato serializzato."""
        import json
        try:
            rows = json.loads(state)
        except json.JSONDecodeError:
            return
        self._notes_restoring = True
        self.notes_table.blockSignals(True)
        try:
            self.notes_table.setRowCount(len(rows))
            cols = max((len(r) for r in rows), default=4)
            self.notes_table.setColumnCount(cols)
            for r, row in enumerate(rows):
                for c in range(cols):
                    val = row[c] if c < len(row) else ""
                    item = self.notes_table.item(r, c)
                    if item:
                        item.setText(val)
                    else:
                        self.notes_table.setItem(r, c, QTableWidgetItem(val))
            self._notes_apply_excel_headers()
        finally:
            self.notes_table.blockSignals(False)
            self._notes_restoring = False

    def _notes_snapshot(self) -> None:
        """Salva lo stato corrente sullo stack undo (solo per tab Tabella)."""
        if not hasattr(self, "notes_tabs") or self.notes_tabs.currentIndex() != 1:
            return
        if self._notes_restoring:
            return
        state = self._notes_get_state()
        if hasattr(self, "_notes_undo_stack"):
            self._notes_undo_stack.append(state)
            if len(self._notes_undo_stack) > self._notes_max_undo:
                self._notes_undo_stack.pop(0)
            self._notes_redo_stack.clear()
            if hasattr(self, "notes_undo_btn"):
                self.notes_undo_btn.setEnabled(True)
            if hasattr(self, "notes_redo_btn"):
                self.notes_redo_btn.setEnabled(False)

    def _notes_undo(self) -> None:
        """Annulla l'ultima modifica."""
        if not self._notes_undo_stack or self.notes_tabs.currentIndex() != 1:
            return
        state = self._notes_undo_stack.pop()
        self._notes_redo_stack.append(self._notes_get_state())
        self._notes_set_state(state)
        self.notes_redo_btn.setEnabled(True)
        if not self._notes_undo_stack:
            self.notes_undo_btn.setEnabled(False)

    def _notes_redo(self) -> None:
        """Ripeti l'ultima modifica annullata."""
        if not self._notes_redo_stack or self.notes_tabs.currentIndex() != 1:
            return
        state = self._notes_redo_stack.pop()
        self._notes_undo_stack.append(self._notes_get_state())
        self._notes_set_state(state)
        self.notes_undo_btn.setEnabled(True)
        if not self._notes_redo_stack:
            self.notes_redo_btn.setEnabled(False)

    def _notes_add_row(self) -> None:
        r = self.notes_table.rowCount()
        self.notes_table.insertRow(r)
        for c in range(self.notes_table.columnCount()):
            self.notes_table.setItem(r, c, QTableWidgetItem(""))
        self._notes_apply_excel_headers()
        self.notes_table.setCurrentCell(r, 0)

    def _on_notes_table_context_menu(self, pos) -> None:
        menu = QMenu(self)
        has_selection = bool(self.notes_table.selectedRanges())
        elimina_righe = menu.addAction("Elimina righe selezionate", self._notes_delete_selected_rows)
        elimina_colonne = menu.addAction("Elimina colonne selezionate", self._notes_delete_selected_columns)
        elimina_righe.setEnabled(has_selection)
        elimina_colonne.setEnabled(has_selection)
        menu.addAction("Aggiungi colonna", self._notes_add_column)
        menu.addSeparator()
        copia_excel = menu.addAction("Copia (per Excel)", lambda: self._copy_notes_to_clipboard("excel"))
        copia_word = menu.addAction("Copia (per Word)", lambda: self._copy_notes_to_clipboard("word"))
        copia_excel.setEnabled(has_selection)
        copia_word.setEnabled(has_selection)
        menu.addSeparator()
        has_clipboard = bool(QGuiApplication.clipboard().text().strip())
        incolla_celle = menu.addAction("Incolla celle (come Excel)", self._paste_notes_table_from_clipboard)
        incolla_tabella = menu.addAction("Incolla come tabella (come Word)", self._paste_notes_as_table)
        incolla_celle.setEnabled(has_clipboard)
        incolla_tabella.setEnabled(has_clipboard)
        menu.exec(self.notes_table.viewport().mapToGlobal(pos))

    def _notes_delete_selected_rows(self) -> None:
        """Elimina tutte le righe contenenti celle selezionate (anche non consecutive)."""
        ranges = self.notes_table.selectedRanges()
        if not ranges:
            return
        rows_to_delete = set()
        for rng in ranges:
            for r in range(rng.topRow(), rng.bottomRow() + 1):
                rows_to_delete.add(r)
        if not rows_to_delete:
            return
        if len(rows_to_delete) >= self.notes_table.rowCount():
            QMessageBox.information(self, "Note", "Deve rimanere almeno una riga.")
            return
        self._notes_snapshot()
        for r in sorted(rows_to_delete, reverse=True):
            self.notes_table.removeRow(r)
        self._notes_apply_excel_headers()

    def _notes_delete_selected_columns(self) -> None:
        """Elimina tutte le colonne contenenti celle selezionate (anche non consecutive)."""
        ranges = self.notes_table.selectedRanges()
        if not ranges:
            return
        cols_to_delete = set()
        for rng in ranges:
            for c in range(rng.leftColumn(), rng.rightColumn() + 1):
                cols_to_delete.add(c)
        if not cols_to_delete:
            return
        if len(cols_to_delete) >= self.notes_table.columnCount():
            QMessageBox.information(self, "Note", "Deve rimanere almeno una colonna.")
            return
        self._notes_snapshot()
        for c in sorted(cols_to_delete, reverse=True):
            self.notes_table.removeColumn(c)
        self._notes_apply_excel_headers()

    def _notes_add_column(self) -> None:
        c = self.notes_table.columnCount()
        self.notes_table.setColumnCount(c + 1)
        for r in range(self.notes_table.rowCount()):
            self.notes_table.setItem(r, c, QTableWidgetItem(""))
        self._notes_apply_excel_headers()

    def _notes_extend_down(self) -> bool:
        """Aggiunge una riga in fondo e sposta la selezione lì. Ritorna True se ha aggiunto."""
        r = self.notes_table.rowCount()
        if r == 0:
            self._notes_add_row()
            return True
        cr = self.notes_table.currentRow()
        cc = self.notes_table.currentColumn()
        if cr == r - 1:
            self.notes_table.insertRow(r)
            for c in range(self.notes_table.columnCount()):
                self.notes_table.setItem(r, c, QTableWidgetItem(""))
            self._notes_apply_excel_headers()
            self.notes_table.setCurrentCell(r, min(cc, self.notes_table.columnCount() - 1))
            return True
        return False

    def _notes_extend_right(self) -> bool:
        """Aggiunge una colonna a destra e sposta la selezione lì. Ritorna True se ha aggiunto."""
        cols = self.notes_table.columnCount()
        cr = self.notes_table.currentRow()
        cc = self.notes_table.currentColumn()
        if cc == cols - 1:
            self.notes_table.setColumnCount(cols + 1)
            for r in range(self.notes_table.rowCount()):
                self.notes_table.setItem(r, cols, QTableWidgetItem(""))
            self._notes_apply_excel_headers()
            self.notes_table.setCurrentCell(cr, cols)
            return True
        return False

    def _notes_get_data_bounds(self) -> tuple[int, int, int, int]:
        """Restituisce (min_row, max_row, min_col, max_col) delle celle con contenuto."""
        min_r = min_c = 0
        max_r = self.notes_table.rowCount() - 1
        max_c = self.notes_table.columnCount() - 1
        if max_r < 0 or max_c < 0:
            return (0, 0, 0, 0)
        found_any = False
        for r in range(max_r + 1):
            for c in range(max_c + 1):
                item = self.notes_table.item(r, c)
                if item and item.text().strip():
                    if not found_any:
                        min_r = max_r = r
                        min_c = max_c = c
                        found_any = True
                    else:
                        min_r = min(min_r, r)
                        max_r = max(max_r, r)
                        min_c = min(min_c, c)
                        max_c = max(max_c, c)
        if not found_any:
            return (0, max_r, 0, max_c)
        return (min_r, max_r, min_c, max_c)

    def _notes_select_to_edge(self, direction: Qt.Key) -> bool:
        """Ctrl+Shift+frecce: seleziona fino all'ultima cella compilata in quella direzione."""
        cr = self.notes_table.currentRow()
        cc = self.notes_table.currentColumn()
        if cr < 0 or cc < 0:
            return False
        min_r, max_r, min_c, max_c = self._notes_get_data_bounds()
        if direction == Qt.Key.Key_Right:
            r1, c1, r2, c2 = cr, cc, cr, max_c
        elif direction == Qt.Key.Key_Down:
            r1, c1, r2, c2 = cr, cc, max_r, cc
        elif direction == Qt.Key.Key_Left:
            r1, c1, r2, c2 = cr, min_c, cr, cc
        elif direction == Qt.Key.Key_Up:
            r1, c1, r2, c2 = min_r, cc, cr, cc
        else:
            return False
        rng = QTableWidgetSelectionRange(r1, c1, r2, c2)
        self.notes_table.clearSelection()
        self.notes_table.setRangeSelected(rng, True)
        self.notes_table.setCurrentCell(cr, cc)
        return True

    def _notes_clear_sheet(self) -> None:
        """Pulisce il contenuto del tab corrente (Testo o Tabella)."""
        if self.notes_tabs.currentIndex() == 1:
            self._notes_snapshot()
        if self.notes_tabs.currentIndex() == 0:
            self.notes_text_edit.clear()
        else:
            for r in range(self.notes_table.rowCount()):
                for c in range(self.notes_table.columnCount()):
                    item = self.notes_table.item(r, c)
                    if item:
                        item.setText("")

    def _open_notes_formulas_dialog(self) -> None:
        """Apre la finestra Formule (solo nel tab Tabella)."""
        if self.notes_tabs.currentIndex() != 1:
            self.notes_tabs.setCurrentIndex(1)
        dlg = _NotesFormulasDialog(self.notes_table, self)
        dlg.exec()

    def _on_notes_cell_changed(self, row: int, col: int, _prev_row: int, _prev_col: int) -> None:
        """Aggiorna il riquadro anteprima quando cambia la cella selezionata."""
        self._notes_cell_preview_block = True
        try:
            if row >= 0 and col >= 0:
                item = self.notes_table.item(row, col)
                text = item.text() if item else ""
                self.notes_cell_preview.setPlainText(text)
            else:
                self.notes_cell_preview.setPlainText("")
        finally:
            self._notes_cell_preview_block = False

    def _on_notes_cell_preview_changed(self) -> None:
        """Propaga le modifiche dall'anteprima alla cella corrente."""
        if self._notes_cell_preview_block:
            return
        cr, cc = self.notes_table.currentRow(), self.notes_table.currentColumn()
        if cr < 0 or cc < 0:
            return
        text = self.notes_cell_preview.toPlainText()
        item = self.notes_table.item(cr, cc)
        if item:
            item.setText(text)
        else:
            self.notes_table.setItem(cr, cc, QTableWidgetItem(text))

    def _notes_clear_selected_cells(self) -> None:
        """Cancella il contenuto di tutte le celle selezionate."""
        self._notes_snapshot()
        for rng in self.notes_table.selectedRanges():
            for r in range(rng.topRow(), rng.bottomRow() + 1):
                for c in range(rng.leftColumn(), rng.rightColumn() + 1):
                    item = self.notes_table.item(r, c)
                    if item:
                        item.setText("")
        cr, cc = self.notes_table.currentRow(), self.notes_table.currentColumn()
        if cr >= 0 and cc >= 0:
            self._notes_cell_preview_block = True
            self.notes_cell_preview.setPlainText("")
            self._notes_cell_preview_block = False

    def _on_notes_list_selected(self, index: int) -> None:
        if index < 0:
            return
        note_id = self.notes_list_combo.currentData(Qt.ItemDataRole.UserRole)
        if note_id is None:
            self._clear_notes_editor()
            return
        self._load_note(int(note_id))

    def _clear_notes_editor(self) -> None:
        self._current_note_id = None
        self.notes_title_edit.clear()
        self.notes_text_edit.clear()
        self.notes_table.setRowCount(0)
        self.notes_table.setColumnCount(4)
        self._notes_apply_excel_headers()
        self._notes_add_row()
        self.notes_tabs.setCurrentIndex(0)
        if hasattr(self, "_notes_undo_stack"):
            self._notes_undo_stack.clear()
            self._notes_redo_stack.clear()
            self.notes_undo_btn.setEnabled(False)
            self.notes_redo_btn.setEnabled(False)

    def _render_client_notes(self, client: dict) -> None:
        if not hasattr(self, "notes_list_combo"):
            return
        self._current_note_id = None
        self.notes_list_combo.blockSignals(True)
        self.notes_list_combo.clear()
        self.notes_list_combo.addItem("-- Nuova nota --", None)
        client_id = client.get("id")
        if client_id is None:
            self.notes_list_combo.blockSignals(False)
            self._clear_notes_editor()
            return
        repo = self.repository.clients if hasattr(self.repository, "clients") else self.repository
        notes = repo.list_client_notes(int(client_id))
        for n in notes:
            title = (n.get("title") or "Senza titolo").strip()
            created = (n.get("created_at") or "")[:10]
            label = f"{title} ({created})"
            self.notes_list_combo.addItem(label, int(n["id"]))
        self.notes_list_combo.blockSignals(False)
        self._clear_notes_editor()

    def _load_note(self, note_id: int) -> None:
        repo = self.repository.clients if hasattr(self.repository, "clients") else self.repository
        note = repo.get_client_note(note_id)
        if not note:
            return
        self._current_note_id = note_id
        self.notes_title_edit.setText(note.get("title") or "")
        ct = note.get("content_type") or "text"
        content = note.get("content") or ""
        if ct == "text":
            self.notes_text_edit.setPlainText(content)
            self.notes_tabs.setCurrentIndex(0)
        else:
            try:
                import json
                rows = json.loads(content) if content else []
                self.notes_table.setRowCount(len(rows))
                cols = max((len(r) for r in rows), default=4)
                self.notes_table.setColumnCount(cols)
                for r, row in enumerate(rows):
                    for c, val in enumerate(row):
                        if c >= self.notes_table.columnCount():
                            self.notes_table.setColumnCount(c + 1)
                        item = QTableWidgetItem(str(val))
                        self.notes_table.setItem(r, c, item)
                self._notes_apply_excel_headers()
            except json.JSONDecodeError:
                self.notes_table.setRowCount(0)
                self._notes_apply_excel_headers()
            if self.notes_table.rowCount() == 0:
                self._notes_add_row()
            self.notes_tabs.setCurrentIndex(1)
            if self.notes_table.rowCount() > 0:
                self.notes_table.setCurrentCell(0, 0)
            if hasattr(self, "_notes_undo_stack"):
                self._notes_undo_stack.clear()
                self._notes_redo_stack.clear()
                self.notes_undo_btn.setEnabled(False)
                self.notes_redo_btn.setEnabled(False)

    def _persist_current_note(self, *, refresh_list: bool = True) -> bool:
        """Salva titolo e contenuto della nota corrente nel database."""
        client_id = getattr(self, "selected_client_id", None)
        if client_id is None:
            QMessageBox.information(self, "Note", "Seleziona un cliente prima di salvare.")
            return False
        title = self.notes_title_edit.text().strip() or "Senza titolo"
        is_table = self.notes_tabs.currentIndex() == 1
        if is_table:
            content = self._notes_table_to_json()
        else:
            content = self.notes_text_edit.toPlainText()
        repo = self.repository.clients if hasattr(self.repository, "clients") else self.repository
        try:
            repo.save_client_note(
                self._current_note_id, int(client_id), title,
                "table" if is_table else "text", content
            )
        except ValueError as exc:
            QMessageBox.warning(self, "Note", str(exc))
            return False
        if refresh_list:
            client = self.clients_by_id.get(int(client_id))
            if client:
                self._render_client_notes(client)
        return True

    def _save_note_to_list(self) -> None:
        if not self._persist_current_note(refresh_list=True):
            return
        QMessageBox.information(self, "Note", "Nota salvata nella lista.")

    def _notes_table_to_json(self) -> str:
        import json
        rows = []
        for r in range(self.notes_table.rowCount()):
            row = []
            for c in range(self.notes_table.columnCount()):
                item = self.notes_table.item(r, c)
                row.append(item.text() if item else "")
            rows.append(row)
        return json.dumps(rows, ensure_ascii=False)

    def _notes_table_from_plain_text(self, text: str) -> None:
        """Sostituisce la tabella con i dati dal testo (tab/CR separati)."""
        lines = [line for line in text.splitlines()]
        sep = "\t" if any("\t" in line for line in lines) else ","
        rows = [[c.strip() for c in line.split(sep)] for line in lines]
        if not rows:
            self.notes_table.setRowCount(0)
            self.notes_table.setColumnCount(4)
            self._notes_apply_excel_headers()
            return
        cols = max(len(r) for r in rows)
        self.notes_table.setRowCount(len(rows))
        self.notes_table.setColumnCount(cols)
        for r, row in enumerate(rows):
            for c in range(cols):
                val = row[c] if c < len(row) else ""
                self.notes_table.setItem(r, c, QTableWidgetItem(val))
        self._notes_apply_excel_headers()

    def _new_note(self) -> None:
        self._current_note_id = None
        self.notes_title_edit.clear()
        self.notes_text_edit.clear()
        self.notes_table.setRowCount(0)
        self.notes_table.setColumnCount(4)
        self._notes_apply_excel_headers()
        self._notes_add_row()
        self.notes_tabs.setCurrentIndex(0)
        if hasattr(self, "notes_list_combo"):
            self.notes_list_combo.setCurrentIndex(0)

    def _delete_selected_note(self) -> None:
        note_id = self.notes_list_combo.currentData(Qt.ItemDataRole.UserRole)
        if note_id is None:
            QMessageBox.information(self, "Note", "Seleziona una nota dalla lista da eliminare.")
            return
        if QMessageBox.question(
            self, "Conferma", "Eliminare questa nota dalla lista?"
        ) != QMessageBox.StandardButton.Yes:
            return
        repo = self.repository.clients if hasattr(self.repository, "clients") else self.repository
        repo.delete_client_note(int(note_id))
        client_id = getattr(self, "selected_client_id", None)
        if client_id:
            client = self.clients_by_id.get(int(client_id))
            if client:
                self._render_client_notes(client)
        self._new_note()

    def _export_note_to_file_and_archive(self) -> None:
        client_id = getattr(self, "selected_client_id", None)
        if client_id is None:
            QMessageBox.information(self, "Note", "Seleziona un cliente prima di esportare.")
            return
        client = self.clients_by_id.get(int(client_id))
        if not client:
            return
        if not self._persist_current_note(refresh_list=False):
            return
        client_name = (client.get("name") or "Cliente").strip()
        is_table = self.notes_tabs.currentIndex() == 1
        if is_table:
            ext_filter = "Fogli Excel (*.xlsx);;File di testo (*.txt);;Tutti i file (*.*)"
            default_ext = "xlsx"
        else:
            ext_filter = "File di testo (*.txt);;Tutti i file (*.*)"
            default_ext = "txt"

        try:
            path, selected_filter = QFileDialog.getSaveFileName(
                self, "Salva nota in file",
                f"Nota_{client_name}.{default_ext}",
                ext_filter,
            )
            if not path or not path.strip():
                return

            is_xlsx = path.lower().endswith(".xlsx") or "xlsx" in (selected_filter or "").lower()
            try:
                if is_xlsx:
                    from openpyxl import Workbook
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Note"
                    for r in range(self.notes_table.rowCount()):
                        for c in range(self.notes_table.columnCount()):
                            item = self.notes_table.item(r, c)
                            val = item.text() if item else ""
                            ws.cell(row=r + 1, column=c + 1, value=val)
                    wb.save(path)
                else:
                    if is_table:
                        lines = []
                        for r in range(self.notes_table.rowCount()):
                            row = []
                            for c in range(self.notes_table.columnCount()):
                                item = self.notes_table.item(r, c)
                                row.append(item.text() if item else "")
                            lines.append("\t".join(row))
                        text_content = "\n".join(lines)
                    else:
                        text_content = self.notes_text_edit.toPlainText()
                    with open(path, "w", encoding="utf-8") as f:
                        f.write(text_content)
            except OSError as exc:
                QMessageBox.warning(self, "Note", f"Errore salvataggio file:\n{exc}")
                return

            repo = self.repository.clients if hasattr(self.repository, "clients") else self.repository
            try:
                repo.get_or_create_tag_for_client(int(client_id), client_name)
            except ValueError:
                pass
            try:
                if hasattr(self.repository, "archive"):
                    self.repository.archive.add_file(None, path, client_name)
                else:
                    self.repository.add_archive_file(None, path, client_name)
            except ValueError as exc:
                QMessageBox.warning(self, "Note", f"File salvato ma errore archivio:\n{exc}")
                return
            QMessageBox.information(
                self,
                "Note",
                "Nota salvata nella lista, file esportato e aggiunto all'archivio con tag del cliente.",
            )
            if hasattr(self, "refresh_views"):
                self.refresh_views()
        finally:
            self._render_client_notes(client)

    def eventFilter(self, watched, event):
        if (
            hasattr(self, "contacts_table")
            and watched is self.contacts_table
            and event.type() == QEvent.Type.KeyPress
            and event.matches(QKeySequence.StandardKey.Paste)
        ):
            self._paste_contacts_from_clipboard()
            return True
        if (
            hasattr(self, "notes_table")
            and watched is self.notes_table
            and event.type() == QEvent.Type.KeyPress
        ):
            if event.matches(QKeySequence.StandardKey.Copy):
                self._copy_notes_to_clipboard("both")
                return True
            if event.matches(QKeySequence.StandardKey.Paste):
                self._paste_notes_table_from_clipboard()
                return True
            mods = event.modifiers()
            ctrl_shift = (
                (mods & Qt.KeyboardModifier.ControlModifier)
                and (mods & Qt.KeyboardModifier.ShiftModifier)
            )
            if ctrl_shift and event.key() in (
                Qt.Key.Key_Left, Qt.Key.Key_Right,
                Qt.Key.Key_Up, Qt.Key.Key_Down
            ):
                if self._notes_select_to_edge(event.key()):
                    return True
            if event.key() == Qt.Key.Key_Down and self._notes_extend_down():
                return True
            if event.key() == Qt.Key.Key_Right and self._notes_extend_right():
                return True
            if event.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
                sel = self.notes_table.selectedRanges()
                if sel:
                    self._notes_clear_selected_cells()
                    return True
            if event.matches(QKeySequence.StandardKey.Undo):
                self._notes_undo()
                return True
            if event.matches(QKeySequence.StandardKey.Redo):
                self._notes_redo()
                return True
        if (
            hasattr(self, "notes_cell_preview")
            and watched is self.notes_cell_preview
            and event.type() == QEvent.Type.FocusIn
        ):
            if self.notes_tabs.currentIndex() == 1:
                self._notes_snapshot()
        return super().eventFilter(watched, event)

    def _copy_notes_to_clipboard(self, mode: str = "both") -> None:
        """Copia le celle selezionate per Excel (TSV) e/o Word (HTML tabella)."""
        ranges = self.notes_table.selectedRanges()
        if not ranges:
            return
        min_r = min(rng.topRow() for rng in ranges)
        max_r = max(rng.bottomRow() for rng in ranges)
        min_c = min(rng.leftColumn() for rng in ranges)
        max_c = max(rng.rightColumn() for rng in ranges)

        def is_selected(r: int, c: int) -> bool:
            return any(
                rng.topRow() <= r <= rng.bottomRow()
                and rng.leftColumn() <= c <= rng.rightColumn()
                for rng in ranges
            )

        cells: list[list[str]] = []
        for r in range(min_r, max_r + 1):
            row_list: list[str] = []
            for c in range(min_c, max_c + 1):
                item = self.notes_table.item(r, c) if is_selected(r, c) else None
                row_list.append(item.text() if item else "")
            cells.append(row_list)

        tsv = "\n".join("\t".join(cell for cell in row) for row in cells)

        html_rows = []
        for row in cells:
            cells_html = "".join(f"<td>{escape(str(c))}</td>" for c in row)
            html_rows.append(f"<tr>{cells_html}</tr>")
        html = f"<html><body><table border='1'><tbody>{''.join(html_rows)}</tbody></table></body></html>"

        mime = QMimeData()
        if mode == "word":
            mime.setHtml(html)
        else:
            mime.setText(tsv)
            if mode == "both":
                mime.setHtml(html)
        QGuiApplication.clipboard().setMimeData(mime)

    def _paste_notes_as_table(self) -> None:
        """Incolla sostituendo l'intera tabella con la struttura incollata (come Word)."""
        self._notes_snapshot()
        text = QGuiApplication.clipboard().text()
        if not text or not text.strip():
            return
        self._notes_table_from_plain_text(text)
        if self.notes_table.rowCount() == 0:
            self._notes_add_row()

    def _paste_notes_table_from_clipboard(self) -> None:
        """Incolla nel punto corrente cella per cella (come Excel)."""
        if self.notes_tabs.currentIndex() != 1:
            return
        self._notes_snapshot()
        text = QGuiApplication.clipboard().text()
        if not text or not text.strip():
            return
        lines = [line for line in text.splitlines()]
        sep = "\t" if any("\t" in line for line in lines) else ","
        rows = [[c.strip() for c in line.split(sep)] for line in lines]
        if not rows:
            return
        anchor_r = max(0, self.notes_table.currentRow())
        anchor_c = max(0, self.notes_table.currentColumn())
        max_cols = max(len(row) for row in rows)
        need_rows = anchor_r + len(rows)
        need_cols = anchor_c + max_cols
        while self.notes_table.rowCount() < need_rows:
            r = self.notes_table.rowCount()
            self.notes_table.insertRow(r)
            for c in range(self.notes_table.columnCount()):
                self.notes_table.setItem(r, c, QTableWidgetItem(""))
        if need_cols > self.notes_table.columnCount():
            self.notes_table.setColumnCount(need_cols)
            for r in range(self.notes_table.rowCount()):
                for c in range(self.notes_table.columnCount()):
                    if self.notes_table.item(r, c) is None:
                        self.notes_table.setItem(r, c, QTableWidgetItem(""))
        for dr, row in enumerate(rows):
            r = anchor_r + dr
            for dc, val in enumerate(row):
                c = anchor_c + dc
                item = self.notes_table.item(r, c)
                if item:
                    item.setText(val)
                else:
                    self.notes_table.setItem(r, c, QTableWidgetItem(val))
        self._notes_apply_excel_headers()

    def _paste_contacts_from_clipboard(self) -> None:
        client_id = getattr(self, "selected_client_id", None)
        if client_id is None:
            QMessageBox.information(self, "Rubrica", "Seleziona un cliente prima di incollare.")
            return

        text = QGuiApplication.clipboard().text()
        if not text or not text.strip():
            return

        lines = [line for line in text.splitlines() if line.strip() != ""]
        if not lines:
            return

        created = 0
        errors: list[str] = []

        for idx, line in enumerate(lines, start=1):
            sep = "\t" if "\t" in line else ","
            parts = [p.strip() for p in line.split(sep)]
            # Expected columns: Nome, Telefono, Cellulare, Mail, Ruolo, Note
            while len(parts) < 6:
                parts.append("")
            name, phone, mobile, email, role, note = parts[:6]
            if not name:
                errors.append(f"Riga {idx}: Nome obbligatorio.")
                continue

            payload = {
                "name": name,
                "phone": phone,
                "mobile": mobile,
                "email": email,
                "role": role,
                "note": note,
            }
            try:
                if hasattr(self.repository, "clients"):
                    self.repository.clients.upsert_client_contact(None, int(client_id), **payload)
                else:
                    self.repository.upsert_client_contact(None, int(client_id), **payload)
                created += 1
            except ValueError as exc:
                errors.append(f"Riga {idx}: {exc}")

        client = self.clients_by_id.get(int(client_id))
        if client:
            self._render_client_contacts(client)

        if errors:
            QMessageBox.warning(
                self,
                "Rubrica",
                f"Incollati {created} contatti.\n\nErrori:\n" + "\n".join(errors),
            )
        else:
            QMessageBox.information(self, "Rubrica", f"Incollati {created} contatti.")
    
    # Actions

    def _on_archive_client_selected(
        self, current: QTreeWidgetItem | None, _previous: QTreeWidgetItem | None
    ) -> None:
        if current is None:
            self.selected_client_id = None
            self.selected_product_id = None
            self._clear_client_details("Nessun cliente selezionato.")
            self._apply_clients_tree_selection_style()
            return
    
        item_type = current.data(0, Qt.ItemDataRole.UserRole)
        if item_type == "product":
            client_id = current.data(0, Qt.ItemDataRole.UserRole + 1)
            product_id = current.data(0, Qt.ItemDataRole.UserRole + 2)
            if client_id is None:
                self._clear_client_details("Cliente non trovato.")
                return
            client = self.clients_by_id.get(int(client_id))
            if client is None:
                self._clear_client_details("Cliente non trovato.")
                return
            self.selected_client_id = int(client_id)
            self.selected_product_id = int(product_id) if product_id is not None else None
            self._render_client_details(client, self._product_by_id(self.selected_product_id))
            self._apply_clients_tree_selection_style()
            return
    
        client_id = current.data(0, Qt.ItemDataRole.UserRole + 1)
        if client_id is None:
            self._clear_client_details("Nessun cliente selezionato.")
            return
        client = self.clients_by_id.get(int(client_id))
        if client is None:
            self._clear_client_details("Cliente non trovato.")
            return
        self.selected_client_id = int(client_id)
        self.selected_product_id = None
        self._render_client_details(client, None)
        self._apply_clients_tree_selection_style()
    

    def _on_clients_tree_menu(self, position) -> None:
        item = self.clients_tree.itemAt(position)
        if item is None:
            return
        if item.data(0, Qt.ItemDataRole.UserRole) != "product":
            return
    
        client_id = item.data(0, Qt.ItemDataRole.UserRole + 1)
        product_id = item.data(0, Qt.ItemDataRole.UserRole + 2)
        if client_id is None or product_id is None:
            return
    
        client = self.clients_by_id.get(int(client_id))
        product = self._product_by_id(int(product_id))
        if client is None or product is None:
            return
    
        menu = QMenu(self)
        add_credential_action = menu.addAction("Nuova credenziale")
        chosen = menu.exec(self.clients_tree.viewport().mapToGlobal(position))
        if chosen != add_credential_action:
            return
    
        flags = (
            self.repository.credentials.get_product_type_flags_for_product(int(product_id))
            if hasattr(self.repository, "credentials")
            else self.repository.get_product_type_flags_for_product(int(product_id))
        )
        dialog = CredentialDialog(self.repository, client, product, flags, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh_views()

    def _on_client_docs_files_menu(self, pos) -> None:
        item = self.client_docs_files_table.itemAt(pos)
        if item is None:
            return
        row = item.row()
        path = self.client_docs_files_table.item(row, 1).text()
        file_rows = (
            self.repository.archive.list_files_all()
            if hasattr(self.repository, "archive")
            else self.repository.list_archive_files_all()
        )
        match = next((r for r in file_rows if str(r.get("path") or "") == path), None)
        if match is None:
            return
        menu = QMenu(self)
        menu.addAction(
            "Aggiungi ai preferiti",
            lambda: self._add_favorite("file", int(match["id"])),
        )
        menu.exec(self.client_docs_files_table.mapToGlobal(pos))

    def _on_client_docs_links_menu(self, pos) -> None:
        item = self.client_docs_links_table.itemAt(pos)
        if item is None:
            return
        row = item.row()
        url = self.client_docs_links_table.item(row, 1).text()
        link_rows = (
            self.repository.archive.list_links_all()
            if hasattr(self.repository, "archive")
            else self.repository.list_archive_links_all()
        )
        match = next((r for r in link_rows if str(r.get("url") or "") == url), None)
        if match is None:
            return
        menu = QMenu(self)
        menu.addAction(
            "Aggiungi ai preferiti",
            lambda: self._add_favorite("link", int(match["id"])),
        )
        menu.exec(self.client_docs_links_table.mapToGlobal(pos))

    def _on_access_credentials_menu(self, position) -> None:
        item = self.access_credentials_table.itemAt(position)
        if item is None:
            return
        self.access_credentials_table.setCurrentCell(item.row(), 0)
        self._update_access_credentials_actions()
        if self._selected_credential_id() is None:
            return
    
        menu = QMenu(self)
        edit_action = menu.addAction("Modifica credenziale")
        delete_action = menu.addAction("Elimina credenziale")
        chosen = menu.exec(self.access_credentials_table.viewport().mapToGlobal(position))
        if chosen == edit_action:
            self._edit_selected_credential()
            return
        if chosen == delete_action:
            self._delete_selected_credential()
    

    def _selected_credential_id(self) -> int | None:
        row = self.access_credentials_table.currentRow()
        if row < 0:
            return None
        item = self.access_credentials_table.item(row, 0)
        if item is None:
            return None
        value = item.data(Qt.ItemDataRole.UserRole)
        if value is None:
            return None
        try:
            return int(value)
        except (TypeError, ValueError):
            return None
    

    def _update_access_credentials_actions(self) -> None:
        if not hasattr(self, "edit_credential_btn") or not hasattr(self, "delete_credential_btn"):
            return
        has_selection = self._selected_credential_id() is not None
        self.edit_credential_btn.setEnabled(has_selection)
        self.delete_credential_btn.setEnabled(has_selection)
        self._refresh_access_connection_buttons()
    

    def _refresh_access_connection_buttons(self) -> None:
        buttons = [
            getattr(self, "connect_rdp_ip_btn", None),
            getattr(self, "connect_rdp_host_btn", None),
            getattr(self, "open_url_btn", None),
            getattr(self, "connect_rdp_conf_btn", None),
        ]
        for btn in buttons:
            if btn is not None:
                btn.setEnabled(False)
    
        credential = self._selected_credential_row()
        if credential is None:
            return
    
        flags = getattr(self, "_access_current_flags", {}) or {}
        ip_value = str(credential.get("ip", "") or "").strip()
        host_value = str(credential.get("host", "") or "").strip()
        url_value = str(credential.get("url", "") or "").strip()
        rdp_path = str(credential.get("rdp_path", "") or "").strip()
    
        if self._is_truthy(flags.get("flag_ip")) and ip_value and self.connect_rdp_ip_btn is not None:
            self.connect_rdp_ip_btn.setEnabled(True)
        if self._is_truthy(flags.get("flag_host")) and host_value and self.connect_rdp_host_btn is not None:
            self.connect_rdp_host_btn.setEnabled(True)
        if self._is_truthy(flags.get("flag_url")) and url_value and self.open_url_btn is not None:
            self.open_url_btn.setEnabled(True)
        if (
            self._is_truthy(flags.get("flag_preconfigured"))
            and rdp_path
            and self.connect_rdp_conf_btn is not None
        ):
            self.connect_rdp_conf_btn.setEnabled(True)
    

    def _current_access_target(self) -> tuple[dict, dict] | None:
        if self.selected_client_id is None or self.selected_product_id is None:
            return None
        client = self.clients_by_id.get(int(self.selected_client_id))
        product = self._product_by_id(int(self.selected_product_id))
        if client is None or product is None:
            return None
        return client, product
    

    def _selected_credential_row(self) -> dict | None:
        cred_id = self._selected_credential_id()
        if cred_id is None:
            return None
        rows = getattr(self, "_access_current_rows", []) or []
        for row in rows:
            if int(row.get("id", -1)) == int(cred_id):
                return row
        return None
    

    def _run_selected_credential_action(self, action_key: str) -> None:
        credential = self._selected_credential_row()
        if credential is None:
            QMessageBox.information(
                self,
                "Credenziali Prodotto",
                "Seleziona una credenziale per avviare la connessione.",
            )
            return
        username = str(credential.get("username", "") or "").strip()
        password = str(credential.get("password", "") or "").strip()
        port_value = str(credential.get("port", "") or "").strip()
    
        if action_key == "ip":
            target = str(credential.get("ip", "") or "").strip()
            self._connect_rdp_endpoint(target, port_value, username, password, "IP")
            return
        if action_key == "host":
            target = str(credential.get("host", "") or "").strip()
            self._connect_rdp_endpoint(target, port_value, username, password, "HOST")
            return
        if action_key == "url":
            target = str(credential.get("url", "") or "").strip()
            self._open_credential_url(target)
            return
        if action_key == "rdp_path":
            target = str(credential.get("rdp_path", "") or "").strip()
            self._open_preconfigured_rdp(target)
            return
    

    def _edit_selected_credential(self) -> None:
        credential_id = self._selected_credential_id()
        if credential_id is None:
            QMessageBox.information(
                self,
                "Credenziali Prodotto",
                "Seleziona una credenziale da modificare.",
            )
            return
    
        target = self._current_access_target()
        if target is None:
            QMessageBox.warning(
                self,
                "Credenziali Prodotto",
                "Seleziona un prodotto nella lista clienti prima di modificare la credenziale.",
            )
            return
        client, product = target
    
        try:
            if hasattr(self.repository, "credentials"):
                credential = self.repository.credentials.get_credential_detail(credential_id)
                flags = self.repository.credentials.get_product_type_flags_for_product(int(product["id"]))
            else:
                credential = self.repository.get_product_credential_detail(credential_id)
                flags = self.repository.get_product_type_flags_for_product(int(product["id"]))
        except ValueError as exc:
            QMessageBox.warning(self, "Credenziali Prodotto", str(exc))
            return
    
        dialog = CredentialDialog(
            self.repository,
            client,
            product,
            flags,
            credential=credential,
            parent=self,
        )
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh_views()
    

    def _delete_selected_credential(self) -> None:
        credential_id = self._selected_credential_id()
        if credential_id is None:
            QMessageBox.information(
                self,
                "Credenziali Prodotto",
                "Seleziona una credenziale da eliminare.",
            )
            return
    
        row = self.access_credentials_table.currentRow()
        credential_name = ""
        if row >= 0:
            item = self.access_credentials_table.item(row, 0)
            credential_name = item.text().strip() if item is not None else ""
        label = credential_name or f"ID {credential_id}"
    
        confirm = QMessageBox.question(
            self,
            "Elimina credenziale",
            f"Vuoi eliminare la credenziale '{label}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return
    
        try:
            if hasattr(self.repository, "credentials"):
                self.repository.credentials.delete_credential(credential_id)
            else:
                self.repository.delete_product_credential(credential_id)
        except ValueError as exc:
            QMessageBox.warning(self, "Elimina credenziale", str(exc))
            return
        self.refresh_views()
    
    # Data rendering

    def _render_clients_table(self) -> None:
        if not hasattr(self, "clients_table"):
            return
        self.clients_table.setRowCount(0)
        for row_data in self.clients_cache:
            row = self.clients_table.rowCount()
            self.clients_table.insertRow(row)
            self._set_table_item(self.clients_table, row, 0, row_data.get("code", ""))
            self._set_table_item(self.clients_table, row, 1, row_data.get("name", ""))
            self._set_table_item(
                self.clients_table, row, 2, row_data.get("location", "") or "Italia"
            )
        if hasattr(self, "clients_count_lbl"):
            self.clients_count_lbl.setText(f"{len(self.clients_cache)} clienti")
    

    def _render_archive_client_list(
        self, selected_client_id: int | None, selected_product_id: int | None
    ) -> None:
        self.clients_tree.blockSignals(True)
        self.clients_tree.clear()
        selected_item: QTreeWidgetItem | None = None
    
        filter_text = ""
        if hasattr(self, "clients_filter"):
            filter_text = self.clients_filter.text().strip().lower()
    
        for client in self.clients_cache:
            client_id = int(client["id"])
            client_name = str(client.get("name", "")).strip()
            if filter_text and filter_text not in client_name.lower():
                continue
    
            client_item = QTreeWidgetItem([client_name])
            client_item.setData(0, Qt.ItemDataRole.UserRole, "client")
            client_item.setData(0, Qt.ItemDataRole.UserRole + 1, client_id)
            self.clients_tree.addTopLevelItem(client_item)
    
            linked_products = [
                product
                for product in self.products_cache
                if self._csv_contains(product.get("clients", ""), client_name)
            ]
            linked_products = sorted(
                linked_products, key=lambda row: str(row.get("name", "")).lower()
            )
            for product in linked_products:
                product_id = int(product["id"])
                product_name = str(product.get("name", "")).strip()
                product_item = QTreeWidgetItem([product_name])
                product_item.setData(0, Qt.ItemDataRole.UserRole, "product")
                product_item.setData(0, Qt.ItemDataRole.UserRole + 1, client_id)
                product_item.setData(0, Qt.ItemDataRole.UserRole + 2, product_id)
                client_item.addChild(product_item)
                if (
                    selected_client_id is not None
                    and selected_product_id is not None
                    and client_id == selected_client_id
                    and product_id == selected_product_id
                ):
                    selected_item = product_item
    
            if selected_item is None and selected_client_id is not None and client_id == selected_client_id:
                selected_item = client_item
    
            client_item.setExpanded(False)
    
        self.clients_tree.blockSignals(False)
    
        if self.clients_tree.topLevelItemCount() == 0:
            self.selected_client_id = None
            self.selected_product_id = None
            self._clear_client_details("Nessun cliente disponibile.")
            return
    
        if selected_item is None:
            selected_item = self.clients_tree.topLevelItem(0)
        self.clients_tree.setCurrentItem(selected_item)
        self._apply_clients_tree_selection_style()

    def _apply_clients_tree_selection_style(self) -> None:
        # Delegate-based painting: just request a repaint.
        if hasattr(self, "clients_tree"):
            self.clients_tree.viewport().update()
    

    def _apply_clients_filter(self) -> None:
        selected_client_id, selected_product_id = self._selected_archive_selection()
        self._render_archive_client_list(selected_client_id, selected_product_id)
    

    def _render_client_details(self, client: dict, selected_product: dict | None = None) -> None:
        client_name = client.get("name", "")
        resources_names = self._csv_values(client.get("resources", ""))
        linked_products = [
            product
            for product in self.products_cache
            if self._csv_contains(product.get("clients", ""), client_name)
        ]
    
        selected_name = str(selected_product.get("name", "")).strip() if selected_product else ""
        self.client_detail_title.setText(f"Cliente: {client_name}")
        self.client_detail_meta.setText(
            f"{len(resources_names)} risorse, {len(linked_products)} prodotti"
            + (f" | Prodotto selezionato: {selected_name}" if selected_name else "")
        )
    
        self.info_name_value.setText(str(client_name or "-"))
        self.info_location_value.setText(str(client.get("location", "") or "Italia"))
        self._set_clickable_client_link(str(client.get("link", "") or ""))
        self._render_info_role_sections(resources_names)
        self._render_info_products_area(client, selected_product)
    
        self._render_vpn_access(client)
        self._render_access_credentials(client, selected_product)
        self._render_client_tag_documents(client)
        self._render_client_contacts(client)
        self._render_client_notes(client)

    def _render_access_credentials(self, client: dict, selected_product: dict | None) -> None:
        self.access_credentials_table.setRowCount(0)
        if selected_product is None:
            self._configure_access_credentials_columns([])
            self.access_product_label.setText(
                "Seleziona un prodotto nella lista clienti per vedere le credenziali."
            )
            self._update_access_credentials_actions()
            return
    
        client_id = int(client["id"])
        product_id = int(selected_product["id"])
        product_name = str(selected_product.get("name", "")).strip()
        self.access_product_label.setText(
            f"Prodotto selezionato: {product_name}"
        )
    
        rows = (
            self.repository.credentials.list_product_credentials(client_id, product_id)
            if hasattr(self.repository, "credentials")
            else self.repository.list_product_credentials(client_id, product_id)
        )
        try:
            flags = (
                self.repository.credentials.get_product_type_flags_for_product(product_id)
                if hasattr(self.repository, "credentials")
                else self.repository.get_product_type_flags_for_product(product_id)
            )
        except ValueError:
            flags = {}
    
        self._access_current_rows = rows
        self._access_current_flags = flags
        self._configure_access_credentials_columns(rows)
        if not rows:
            self.access_product_label.setText(
                f"Prodotto selezionato: {product_name} (nessuna credenziale configurata)"
            )
            self._update_access_credentials_actions()
            return
        for row_data in rows:
            row = self.access_credentials_table.rowCount()
            self.access_credentials_table.insertRow(row)
    
            for col, key in enumerate(self.access_credentials_columns):
                self._set_table_item(self.access_credentials_table, row, col, row_data.get(key, ""))
    
            id_item = self.access_credentials_table.item(row, 0)
            if id_item is not None:
                id_item.setData(Qt.ItemDataRole.UserRole, row_data.get("id"))
        self._update_access_credentials_actions()
    

    def _configure_access_credentials_columns(self, rows: list[dict] | None = None) -> None:
        rows = rows or []
        endpoint_columns: list[tuple[str, str]] = []
        for key, label in self.access_endpoint_columns:
            if any(str(row.get(key, "") or "").strip() for row in rows):
                endpoint_columns.append((key, label))
    
        columns: list[tuple[str, str]] = [
            ("credential_name", "Nome"),
            ("environments_versions", "Ambienti/Versioni"),
            *endpoint_columns,
            ("username", "Username"),
            ("password", "Password"),
        ]
        self.access_credentials_columns = [key for key, _ in columns]
        self.access_credentials_table.setColumnCount(len(columns))
        self.access_credentials_table.setHorizontalHeaderLabels([label for _, label in columns])
    
        width_map = {
            "credential_name": 180,
            "environments_versions": 230,
            "ip": 140,
            "host": 170,
            "url": 260,
            "rdp_path": 270,
            "username": 180,
            "password": 160,
        }
        for idx, key in enumerate(self.access_credentials_columns):
            width = width_map.get(key)
            if width is not None:
                self.access_credentials_table.setColumnWidth(idx, width)
        self.access_credentials_table.horizontalHeader().setStretchLastSection(True)
    
    @staticmethod

    def _normalize_rdp_target(target: str, port: str) -> str:
        clean_target = str(target or "").strip()
        clean_port = str(port or "").strip()
        if not clean_port:
            return clean_target
        try:
            port_num = int(clean_port)
        except ValueError:
            return clean_target
        if port_num < 1 or port_num > 65535:
            return clean_target
        if ":" in clean_target:
            return clean_target
        return f"{clean_target}:{port_num}"
    
    @staticmethod

    def _rdp_target_variants(clean_target: str, rdp_target: str) -> list[str]:
        variants: list[str] = []
        seen: set[str] = set()
        for value in (clean_target, rdp_target):
            current = str(value or "").strip()
            if not current:
                continue
            for candidate in (current, current.split(":")[0]):
                key = candidate.lower()
                if candidate and key not in seen:
                    seen.add(key)
                    variants.append(candidate)
        return variants
    
    @staticmethod

    def _build_rdp_launch_file(target: str, username: str) -> str:
        lines = [
            f"full address:s:{target}",
            f"username:s:{username}",
            "prompt for credentials:i:0",
            "promptcredentialonce:i:1",
            "enablecredsspsupport:i:1",
            "authentication level:i:2",
            "negotiate security layer:i:1",
            "redirectclipboard:i:1",
            "screen mode id:i:2",
        ]
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            suffix=".rdp",
            prefix="hdm_rdp_",
            delete=False,
        ) as tmp_file:
            tmp_file.write("\n".join(lines))
            return tmp_file.name
    

    def _connect_rdp_endpoint(
        self, target: str, port: str, username: str, password: str, source_label: str
    ) -> None:
        clean_target = str(target or "").strip()
        clean_user = str(username or "").strip()
        clean_password = str(password or "").strip()
        if not clean_target:
            QMessageBox.warning(self, "Connetti RDP", f"{source_label} non valorizzato.")
            return
        if not clean_user or not clean_password:
            QMessageBox.warning(
                self,
                "Connetti RDP",
                "Username e Password sono obbligatori per avviare la connessione RDP.",
            )
            return

        system = getattr(self, "system", None)
        worker = _RdpConnectWorker(
            target=clean_target,
            port=port,
            username=clean_user,
            password=clean_password,
            system=system,
            parent=self,
        )
        worker.error_occurred.connect(
            lambda msg: QMessageBox.warning(self, "Connetti RDP", msg)
        )
        worker.start()
    

    def _open_credential_url(self, raw_url: str) -> None:
        clean_url = str(raw_url or "").strip()
        if not clean_url:
            QMessageBox.warning(self, "Apri Link", "URL non valorizzato.")
            return
        system = getattr(self, "system", None)
        if system is not None:
            if not system.open_url(clean_url):
                QMessageBox.warning(self, "Apri Link", f"Impossibile aprire il link: {clean_url}")
            return
        if "://" not in clean_url:
            clean_url = f"https://{clean_url}"
        url = QUrl(clean_url)
        if not url.isValid() or not QDesktopServices.openUrl(url):
            QMessageBox.warning(self, "Apri Link", f"Impossibile aprire il link: {clean_url}")
    

    def _open_preconfigured_rdp(self, rdp_path: str) -> None:
        clean_path = str(rdp_path or "").strip()
        if not clean_path:
            QMessageBox.warning(self, "Connetti RDP conf", "Percorso RDP non valorizzato.")
            return
        system = getattr(self, "system", None)
        if system is not None:
            try:
                system.start_file(clean_path)
            except (ValueError, OSError) as exc:
                QMessageBox.warning(
                    self,
                    "Connetti RDP conf",
                    f"Impossibile aprire il file RDP configurato.\n{exc}",
                )
            return
        try:
            if hasattr(os, "startfile"):
                os.startfile(clean_path)  # type: ignore[attr-defined]
            else:
                url = QUrl.fromLocalFile(clean_path)
                if not QDesktopServices.openUrl(url):
                    raise OSError("Apertura file non riuscita.")
        except OSError as exc:
            QMessageBox.warning(
                self,
                "Connetti RDP conf",
                f"Impossibile aprire il file RDP configurato.\n{exc}",
            )
    

    def _render_info_role_sections(self, resource_names: list[str]) -> None:
        self._clear_layout(self.info_roles_layout)
    
        if not resource_names:
            empty = QLabel("Nessuna risorsa collegata.")
            empty.setObjectName("subText")
            self.info_roles_layout.addWidget(empty)
            self.info_roles_layout.addStretch()
            return
    
        lookup: dict[str, dict] = {}
        for resource in self.resources_cache:
            full_name = f"{resource.get('name', '')} {resource.get('surname', '')}".strip().lower()
            if full_name:
                lookup[full_name] = resource
    
        grouped: dict[str, list[dict]] = defaultdict(list)
        for resource_name in resource_names:
            row = lookup.get(resource_name.lower())
            if row is None:
                grouped["Senza ruolo"].append(
                    {
                        "name": resource_name,
                        "surname": "",
                        "phone": "",
                        "email": "",
                        "note": "",
                        "role_name": "Senza ruolo",
                    }
                )
            else:
                role_name = str(row.get("role_name") or "Senza ruolo").strip()
                grouped[role_name].append(row)
    
        for role_name in sorted(
            grouped.keys(),
            key=lambda value: (self.role_order_map.get(value.lower(), 999), value.lower()),
        ):
            resources = grouped[role_name]
            is_multi = self.role_multi_map.get(role_name.lower(), False)
    
            card = QWidget()
            card.setObjectName("roleCard")
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(10, 10, 10, 10)
            card_layout.setSpacing(6)
    
            title = QLabel(role_name)
            title.setObjectName("roleCardTitle")
            card_layout.addWidget(title)
    
            # Mostriamo solo le risorse collegate; non esponiamo "piu clienti si/no"
            # perché è una proprietà tecnica del ruolo e risulta confondente nell'Info Cliente.
    
            if is_multi:
                table = QTableWidget(len(resources), 4)
                table.setHorizontalHeaderLabels(["Risorsa", "Telefono", "Email", "Note"])
                table.verticalHeader().setVisible(False)
                table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
                table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
                table.setAlternatingRowColors(True)
                table.setColumnWidth(0, 250)
                table.setColumnWidth(1, 160)
                table.setColumnWidth(2, 230)
                table.horizontalHeader().setStretchLastSection(True)

                # Rendiamo cliccabile la colonna Email.
                def _on_cell_clicked(r: int, c: int, tbl=table) -> None:
                    if c != 2:
                        return
                    item = tbl.item(r, c)
                    if item is None:
                        return
                    email = item.text().strip()
                    if not email:
                        return
                    QDesktopServices.openUrl(QUrl(f"mailto:{email}"))

                table.cellClicked.connect(_on_cell_clicked)
    
                for row_idx, resource in enumerate(resources):
                    full = f"{resource.get('name', '')} {resource.get('surname', '')}".strip()
                    self._set_table_item(table, row_idx, 0, full)
                    self._set_table_item(table, row_idx, 1, resource.get("phone", ""))
                    self._set_table_item(table, row_idx, 2, resource.get("email", ""))
                    self._set_table_item(table, row_idx, 3, resource.get("note", ""))
    
                card_layout.addWidget(table)
            else:
                parts: list[str] = []
                for resource in resources:
                    full = escape(
                        f"{resource.get('name', '')} {resource.get('surname', '')}".strip()
                    )
                    phone = str(resource.get("phone", "") or "").strip()
                    email = str(resource.get("email", "") or "").strip()
                    note = str(resource.get("note", "") or "").strip()

                    details: list[str] = []
                    if phone:
                        details.append(escape(phone))
                    if email:
                        safe_email = escape(email, quote=True)
                        safe_label = escape(email)
                        details.append(
                            f'<a href="mailto:{safe_email}">{safe_label}</a>'
                        )

                    if details:
                        line = f"{full} ({' - '.join(details)})"
                    else:
                        line = full if full else "-"

                    if note:
                        line += f"<br/>Nota: {escape(note)}"

                    parts.append(line)

                single = QLabel("<br/>".join(parts) if parts else "-")
                single.setObjectName("infoValue")
                single.setWordWrap(True)
                single.setTextFormat(Qt.TextFormat.RichText)
                single.setOpenExternalLinks(True)
                single.setTextInteractionFlags(
                    Qt.TextInteractionFlag.TextBrowserInteraction
                )
                card_layout.addWidget(single)
    
            self.info_roles_layout.addWidget(card)
    
        self.info_roles_layout.addStretch()
    

    def _render_info_products_area(self, client: dict, selected_product: dict | None = None) -> None:
        client_id = client.get("id")
        if client_id is None:
            self.info_products_area.setText("Cliente non valido.")
            return
    
        products = (
            self.repository.clients.list_client_product_environment_releases(int(client_id))
            if hasattr(self.repository, "clients")
            else self.repository.list_client_product_environment_releases(int(client_id))
        )
        if not products:
            self.info_products_area.setText(
                "Nessun prodotto collegato al cliente."
            )
            return
    
        selected_product_id = int(selected_product["id"]) if selected_product and selected_product.get("id") is not None else None
        lines: list[str] = []
        for row in products:
            product_name = str(row.get("product_name", "")).strip()
            summary = str(row.get("summary", "")).strip() or "Nessun ambiente associato"
            product_id = row.get("product_id")
            marker = "-> " if selected_product_id is not None and int(product_id) == selected_product_id else ""
            lines.append(f"{marker}{product_name} => {summary}")
        self.info_products_area.setText("\n".join(lines))
    

    def _render_vpn_access(self, client: dict) -> None:
        vpn_name = (client.get("vpn_name") or "").strip()
        if not vpn_name:
            self._set_vpn_card_values()
            self.vpn_connect_btn.setEnabled(False)
            self._current_vpn_row = None
            return
    
        vpn_row = next(
            (vpn for vpn in self.vpns_cache if vpn.get("connection_name") == vpn_name),
            None,
        )
        if vpn_row is None:
            self._set_vpn_card_values(name=vpn_name)
            self.vpn_connect_btn.setEnabled(False)
            self._current_vpn_row = None
            return
    
        self._set_vpn_card_values(
            name=vpn_name,
            server=vpn_row.get("server_address", ""),
            vpn_type=vpn_row.get("vpn_type", ""),
            access_info=vpn_row.get("access_info_type", ""),
            username=vpn_row.get("username", ""),
            password=vpn_row.get("password", ""),
            vpn_path=vpn_row.get("vpn_path", ""),
        )
        self._current_vpn_row = vpn_row
        vpn_type = str(vpn_row.get("vpn_type", "") or "").strip()
        self.vpn_connect_btn.setText(
            "Connetti" if vpn_type == "VPN Windows" else "Avvia"
        )
        self.vpn_connect_btn.setEnabled(True)
    

    def _set_vpn_card_values(
        self,
        name: str = "-",
        server: str = "-",
        vpn_type: str = "-",
        access_info: str = "-",
        username: str = "-",
        password: str = "-",
        vpn_path: str = "-",
    ) -> None:
        self.vpn_name_field.setText(str(name or "-"))
        self.vpn_server_field.setText(str(server or "-"))
        self.vpn_type_field.setText(str(vpn_type or "-"))
        self.vpn_access_info_field.setText(str(access_info or "-"))
        self.vpn_user_field.setText(str(username or "-"))
        self.vpn_password_field.setText(str(password or "-"))
        self.vpn_path_field.setText(str(vpn_path or "-"))
    

    def _connect_vpn_action(self) -> None:
        vpn_row = getattr(self, "_current_vpn_row", None)
        if not vpn_row:
            QMessageBox.information(self, "VPN", "Nessuna VPN associata al cliente.")
            return
        vpn_type = str(vpn_row.get("vpn_type", "") or "").strip()
        connection_name = str(vpn_row.get("connection_name", "") or "").strip()
        vpn_path = str(vpn_row.get("vpn_path", "") or "").strip()
    
        if vpn_type == "VPN Windows":
            if not connection_name:
                QMessageBox.warning(self, "VPN", "Nome connessione VPN non valido.")
                return
            system = getattr(self, "system", None)
            if system is not None:
                try:
                    system.start_vpn_windows(connection_name)
                except (ValueError, OSError) as exc:
                    QMessageBox.warning(self, "VPN", f"Impossibile avviare VPN Windows.\n{exc}")
                return
            try:
                subprocess.Popen(
                    ["rasdial", connection_name],
                    creationflags=_SUBPROCESS_FLAGS,
                )
            except OSError as exc:
                QMessageBox.warning(self, "VPN", f"Impossibile avviare VPN Windows.\n{exc}")
            return
    
        if not vpn_path:
            QMessageBox.warning(self, "VPN", "Percorso VPN non impostato.")
            return
        system = getattr(self, "system", None)
        if system is not None:
            try:
                system.start_file(vpn_path)
            except (ValueError, OSError) as exc:
                QMessageBox.warning(self, "VPN", f"Impossibile avviare VPN.\n{exc}")
            return
        try:
            if hasattr(os, "startfile"):
                os.startfile(vpn_path)  # type: ignore[attr-defined]
            else:
                subprocess.Popen([vpn_path], creationflags=_SUBPROCESS_FLAGS)
        except OSError as exc:
            QMessageBox.warning(self, "VPN", f"Impossibile avviare VPN.\n{exc}")
    

    def _render_client_tag_documents(self, client: dict) -> None:
        if not hasattr(self, "client_docs_files_table"):
            return
        client_id = client.get("id")
        if client_id is None:
            return
        tags = (
            self.repository.clients.list_tags_for_client(int(client_id))
            if hasattr(self.repository, "clients")
            else self.repository.list_tags_for_client(int(client_id))
        )
        tag_names = {tag["name"] for tag in tags}
    
        files = [
            row
            for row in (
                self.repository.archive.list_files_all()
                if hasattr(self.repository, "archive")
                else self.repository.list_archive_files_all()
            )
            if row.get("tag_name") in tag_names
        ]
        links = [
            row
            for row in (
                self.repository.archive.list_links_all()
                if hasattr(self.repository, "archive")
                else self.repository.list_archive_links_all()
            )
            if row.get("tag_name") in tag_names
        ]
    
        self.client_files_count.setText(str(len(files)))
        self.client_links_count.setText(str(len(links)))
    
        self.client_docs_files_table.setRowCount(0)
        for row_data in files:
            row = self.client_docs_files_table.rowCount()
            self.client_docs_files_table.insertRow(row)
            self._set_table_item(self.client_docs_files_table, row, 0, row_data.get("name", ""))
            self._set_table_item(self.client_docs_files_table, row, 1, row_data.get("path", ""))
    
        self.client_docs_links_table.setRowCount(0)
        for row_data in links:
            row = self.client_docs_links_table.rowCount()
            self.client_docs_links_table.insertRow(row)
            self._set_table_item(self.client_docs_links_table, row, 0, row_data.get("name", ""))
            self._set_table_item(self.client_docs_links_table, row, 1, row_data.get("url", ""))
    

    def _render_client_contacts(self, client: dict) -> None:
        if not hasattr(self, "contacts_table"):
            return
        client_id = client.get("id")
        if client_id is None:
            return
        rows = (
            self.repository.clients.list_client_contacts(int(client_id))
            if hasattr(self.repository, "clients")
            else self.repository.list_client_contacts(int(client_id))
        )
        self.contacts_table.setRowCount(0)
        for row_data in rows:
            row = self.contacts_table.rowCount()
            self.contacts_table.insertRow(row)
            self._set_table_item(self.contacts_table, row, 0, row_data.get("name", ""))
            self._set_table_item(self.contacts_table, row, 1, row_data.get("phone", ""))
            self._set_table_item(self.contacts_table, row, 2, row_data.get("mobile", ""))
            self._set_table_item(self.contacts_table, row, 3, row_data.get("email", ""))
            self._set_table_item(self.contacts_table, row, 4, row_data.get("role", ""))
            self._set_table_item(self.contacts_table, row, 5, row_data.get("note", ""))
            id_item = self.contacts_table.item(row, 0)
            if id_item is not None:
                id_item.setData(Qt.ItemDataRole.UserRole, row_data.get("id"))
    
        self._update_contacts_actions()
    

    def _update_contacts_actions(self) -> None:
        has_selection = bool(self.contacts_table.selectedIndexes())
        self.contacts_edit_btn.setEnabled(has_selection)
        self.contacts_delete_btn.setEnabled(has_selection)
    

    def _selected_contact(self) -> dict | None:
        selected = self.contacts_table.selectedIndexes()
        if not selected:
            return None
        row = selected[0].row()
        id_item = self.contacts_table.item(row, 0)
        contact_id = id_item.data(Qt.ItemDataRole.UserRole) if id_item else None
        if contact_id is None:
            return None
        return {
            "id": contact_id,
            "name": self.contacts_table.item(row, 0).text(),
            "phone": self.contacts_table.item(row, 1).text(),
            "mobile": self.contacts_table.item(row, 2).text(),
            "email": self.contacts_table.item(row, 3).text(),
            "role": self.contacts_table.item(row, 4).text(),
            "note": self.contacts_table.item(row, 5).text(),
        }
    

    def _add_contact(self) -> None:
        client_id = getattr(self, "selected_client_id", None)
        if client_id is None:
            QMessageBox.information(self, "Rubrica", "Seleziona un cliente.")
            return
        roles = (
            [row["label"] for row in self.repository.clients.list_roles_lookup()]
            if hasattr(self.repository, "clients")
            else [row["label"] for row in self.repository.list_roles_lookup()]
        )
        dialog = ContactDialog(roles, parent=self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        payload = dialog.values()
        try:
            if hasattr(self.repository, "clients"):
                self.repository.clients.upsert_client_contact(None, int(client_id), **payload)
            else:
                self.repository.upsert_client_contact(None, int(client_id), **payload)
            client = self.clients_by_id.get(int(client_id))
            if client:
                self._render_client_contacts(client)
        except ValueError as exc:
            QMessageBox.warning(self, "Rubrica", str(exc))
    

    def _edit_contact(self) -> None:
        client_id = getattr(self, "selected_client_id", None)
        if client_id is None:
            QMessageBox.information(self, "Rubrica", "Seleziona un cliente.")
            return
        contact = self._selected_contact()
        if contact is None:
            QMessageBox.information(self, "Rubrica", "Seleziona un contatto.")
            return
        roles = (
            [row["label"] for row in self.repository.clients.list_roles_lookup()]
            if hasattr(self.repository, "clients")
            else [row["label"] for row in self.repository.list_roles_lookup()]
        )
        dialog = ContactDialog(roles, contact, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        payload = dialog.values()
        try:
            if hasattr(self.repository, "clients"):
                self.repository.clients.upsert_client_contact(
                    int(contact["id"]),
                    int(client_id),
                    **payload,
                )
            else:
                self.repository.upsert_client_contact(
                    int(contact["id"]),
                    int(client_id),
                    **payload,
                )
            client = self.clients_by_id.get(int(client_id))
            if client:
                self._render_client_contacts(client)
        except ValueError as exc:
            QMessageBox.warning(self, "Rubrica", str(exc))
    

    def _delete_contact(self) -> None:
        contact = self._selected_contact()
        if contact is None:
            QMessageBox.information(self, "Rubrica", "Seleziona un contatto.")
            return
        if (
            QMessageBox.question(
                self,
                "Elimina contatto",
                "Eliminare il contatto selezionato?",
            )
            != QMessageBox.StandardButton.Yes
        ):
            return
        if hasattr(self.repository, "clients"):
            self.repository.clients.delete_client_contact(int(contact["id"]))
        else:
            self.repository.delete_client_contact(int(contact["id"]))
        client_id = getattr(self, "selected_client_id", None)
        if client_id is not None:
            client = self.clients_by_id.get(int(client_id))
            if client:
                self._render_client_contacts(client)
    

    def _handle_contact_cell_click(self, row: int, column: int) -> None:
        if column in {1, 2}:
            item = self.contacts_table.item(row, column)
            if item is None:
                return
            value = item.text().strip()
            if not value:
                return
            clipboard = QGuiApplication.clipboard()
            clipboard.setText(value)
            QMessageBox.information(self, "Rubrica", "Numero copiato negli appunti.")
    

    def _handle_contact_cell_double_click(self, row: int, column: int) -> None:
        if column == 3:
            item = self.contacts_table.item(row, column)
            if item is None:
                return
            email = item.text().strip()
            if not email:
                return
            QDesktopServices.openUrl(QUrl(f"mailto:{email}"))
    

    def _open_client_doc_file(self, row: int, column: int) -> None:
        if not hasattr(self, "client_docs_files_table"):
            return
        path_item = self.client_docs_files_table.item(row, 1)
        if path_item is None:
            return
        path = path_item.text().strip()
        if not path:
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))
    

    def _open_client_doc_link(self, row: int, column: int) -> None:
        if not hasattr(self, "client_docs_links_table"):
            return
        url_item = self.client_docs_links_table.item(row, 1)
        if url_item is None:
            return
        url = url_item.text().strip()
        if not url:
            return
        QDesktopServices.openUrl(QUrl(url))
    

    def _clear_client_details(self, message: str) -> None:
        self.client_detail_title.setText(message)
        self.client_detail_meta.setText("")
        self.info_name_value.setText("-")
        self.info_location_value.setText("-")
        self._set_clickable_client_link("")
        self._clear_layout(self.info_roles_layout)
        empty = QLabel("Nessuna risorsa collegata.")
        empty.setObjectName("subText")
        self.info_roles_layout.addWidget(empty)
        self.info_roles_layout.addStretch()
        self.info_products_area.setText(
            "Area dedicata ai prodotti collegati. Le regole saranno implementate nel prossimo step."
        )
        self._set_vpn_card_values()
        self.vpn_connect_btn.setEnabled(False)
        self.access_product_label.setText("Seleziona un prodotto nella lista clienti.")
        self._configure_access_credentials_columns([])
        self.access_credentials_table.setRowCount(0)
        self._update_access_credentials_actions()
        self._refresh_access_connection_buttons()
        if hasattr(self, "client_docs_files_table"):
            self.client_docs_files_table.setRowCount(0)
        if hasattr(self, "client_docs_links_table"):
            self.client_docs_links_table.setRowCount(0)
        if hasattr(self, "contacts_table"):
            self.contacts_table.setRowCount(0)

    def _set_clickable_client_link(self, raw_link: str) -> None:
        link = (raw_link or "").strip()
        if not link:
            self.info_link_value.setText("-")
            return

        href = link
        lowered = link.lower()
        if not (
            lowered.startswith("http://")
            or lowered.startswith("https://")
            or lowered.startswith("mailto:")
        ):
            href = f"https://{link}"

        safe_label = escape(link)
        safe_href = escape(href, quote=True)
        self.info_link_value.setText(f'<a href="{safe_href}">{safe_label}</a>')
    
    # Helpers
    @staticmethod

    def _resource_to_string(resource: dict) -> str:
        full = f"{resource.get('name', '')} {resource.get('surname', '')}".strip()
        details: list[str] = []
        phone = str(resource.get("phone", "") or "").strip()
        email = str(resource.get("email", "") or "").strip()
        if phone:
            details.append(phone)
        if email:
            details.append(email)
        if details:
            return f"{full} ({' - '.join(details)})"
        return full
    

    def _selected_archive_selection(self) -> tuple[int | None, int | None]:
        item = self.clients_tree.currentItem()
        if item is None:
            return None, None
    
        item_type = item.data(0, Qt.ItemDataRole.UserRole)
        if item_type == "product":
            client_id = item.data(0, Qt.ItemDataRole.UserRole + 1)
            product_id = item.data(0, Qt.ItemDataRole.UserRole + 2)
            if client_id is None:
                return None, None
            return int(client_id), int(product_id) if product_id is not None else None
    
        client_id = item.data(0, Qt.ItemDataRole.UserRole + 1)
        if client_id is None:
            return None, None
        return int(client_id), None
