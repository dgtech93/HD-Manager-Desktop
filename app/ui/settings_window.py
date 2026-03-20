from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

from PyQt6.QtCore import QEvent, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QGuiApplication, QIcon, QKeySequence
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QDialogButtonBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QStyledItemDelegate,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
    QComboBox,
    QColorDialog,
    QFileDialog,
)

from app.services.packaging_service import PackagingService


@dataclass
class TableColumn:
    key: str
    label: str
    width: int = 140
    required: bool = False
    editor: str = "text"  # text, combo, multi, bool
    options_loader: Callable[[], list[str]] | None = None
    visible: bool = True


class ComboBoxDelegate(QStyledItemDelegate):
    def __init__(self, page: "EditableTablePage", parent=None) -> None:
        super().__init__(parent)
        self.page = page

    def createEditor(self, parent, option, index):
        col_cfg = self.page.columns[index.column()]
        if col_cfg.editor != "combo":
            return super().createEditor(parent, option, index)

        combo = QComboBox(parent)
        combo.addItem("")
        combo.addItems(self.page.column_options.get(index.column(), []))
        combo.setEditable(False)
        return combo

    def setEditorData(self, editor, index):
        if isinstance(editor, QComboBox):
            value = index.model().data(index, Qt.ItemDataRole.EditRole) or ""
            editor.setCurrentText(str(value))
            return
        super().setEditorData(editor, index)

    def setModelData(self, editor, model, index):
        if isinstance(editor, QComboBox):
            model.setData(index, editor.currentText(), Qt.ItemDataRole.EditRole)
            return
        super().setModelData(editor, model, index)


class MultiSelectDialog(QDialog):
    def __init__(self, title: str, options: list[str], selected: list[str], parent=None) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(380, 420)

        layout = QVBoxLayout(self)
        info = QLabel("Seleziona uno o piu valori:")
        layout.addWidget(info)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        layout.addWidget(self.list_widget, 1)

        selected_set = {value.strip() for value in selected if value.strip()}
        for option in options:
            item = QListWidgetItem(option)
            self.list_widget.addItem(item)
            if option in selected_set:
                item.setSelected(True)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def selected_values(self) -> list[str]:
        return [item.text().strip() for item in self.list_widget.selectedItems()]


class EditableTablePage(QWidget):
    def __init__(
        self,
        title: str,
        columns: list[TableColumn],
        load_rows: Callable[[], list[dict]],
        save_row: Callable[[int | None, dict], int],
        delete_row: Callable[[int], None],
        get_help_text: Callable[[], str] | None = None,
        on_change: Callable[[], None] | None = None,
    ) -> None:
        super().__init__()
        self.title = title
        self.columns = columns
        self.load_rows = load_rows
        self.save_row = save_row
        self.delete_row = delete_row
        self.get_help_text = get_help_text
        self.on_change = on_change
        self.edit_mode = False
        self.quick_insert_mode = False
        self.column_options: dict[int, list[str]] = {}

        self._build_ui()
        self.refresh_page()

    def _build_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        title_lbl = QLabel(self.title)
        title_lbl.setObjectName("pageTitle")
        title_lbl.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        layout.addWidget(title_lbl)

        self.help_lbl = QLabel("")
        self.help_lbl.setObjectName("helpLabel")
        self.help_lbl.setWordWrap(True)
        layout.addWidget(self.help_lbl)

        actions = QHBoxLayout()
        actions.setSpacing(8)
        self.new_btn = QPushButton("Nuova riga")
        self.edit_btn = QPushButton("Modifica")
        self.edit_btn.setCheckable(True)
        self.save_btn = QPushButton("Salva")
        self.delete_btn = QPushButton("Elimina")
        self.refresh_btn = QPushButton("Aggiorna")
        self.new_btn.setObjectName("primaryActionButton")
        self.save_btn.setObjectName("primaryActionButton")
        self.delete_btn.setObjectName("dangerActionButton")
        self.refresh_btn.setObjectName("secondaryActionButton")

        actions.addWidget(self.new_btn)
        actions.addWidget(self.edit_btn)
        actions.addWidget(self.save_btn)
        actions.addWidget(self.delete_btn)
        actions.addWidget(self.refresh_btn)
        actions.addStretch()
        layout.addLayout(actions)

        self.table = QTableWidget()
        self.table.setColumnCount(len(self.columns))
        self.table.setHorizontalHeaderLabels([c.label for c in self.columns])
        self.table.verticalHeader().setVisible(False)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.horizontalHeader().setStretchLastSection(True)
        for idx, col in enumerate(self.columns):
            self.table.setColumnWidth(idx, col.width)
            if not col.visible:
                self.table.setColumnHidden(idx, True)
        self.table.setItemDelegate(ComboBoxDelegate(self, self.table))
        self.table.installEventFilter(self)
        layout.addWidget(self.table, 1)

        self.status_lbl = QLabel("")
        self.status_lbl.setObjectName("statusLabel")
        layout.addWidget(self.status_lbl)

        self.new_btn.clicked.connect(self.add_row)
        self.edit_btn.toggled.connect(self.toggle_edit_mode)
        self.save_btn.clicked.connect(self.save_changes)
        self.delete_btn.clicked.connect(self.delete_selected)
        self.refresh_btn.clicked.connect(self.refresh_page)
        self.table.cellDoubleClicked.connect(self._handle_multi_selector)

    def _first_input_column(self) -> int:
        for idx, col in enumerate(self.columns):
            if col.visible and col.key not in {"id", "code"}:
                return idx
        return 0

    @staticmethod
    def _is_truthy(value: str | int | bool | None) -> bool:
        if isinstance(value, bool):
            return value
        if isinstance(value, int):
            return value == 1
        text = ("" if value is None else str(value)).strip().lower()
        return text in {"1", "true", "si", "yes", "y"}

    def _refresh_lookup_options(self) -> None:
        self.column_options = {}
        for idx, col in enumerate(self.columns):
            if col.options_loader is not None:
                self.column_options[idx] = col.options_loader()

    def refresh_page(self) -> None:
        self._refresh_lookup_options()
        self.table.setRowCount(0)
        rows = self.load_rows()
        for row in rows:
            self._append_row(row)
        self.quick_insert_mode = False
        self._update_help()
        self.status_lbl.setText(f"Righe: {len(rows)}")

    def _append_row(self, row_data: dict) -> None:
        row = self.table.rowCount()
        self.table.insertRow(row)
        for col_idx, col in enumerate(self.columns):
            value = row_data.get(col.key, "")
            if col.editor == "bool":
                item = QTableWidgetItem("")
                item.setCheckState(
                    Qt.CheckState.Checked
                    if self._is_truthy(value)
                    else Qt.CheckState.Unchecked
                )
            else:
                item = QTableWidgetItem("" if value is None else str(value))
            self.table.setItem(row, col_idx, item)
        self._apply_row_flags(row)

    def _apply_row_flags(self, row: int) -> None:
        for col_idx, col in enumerate(self.columns):
            item = self.table.item(row, col_idx)
            if item is None:
                continue

            flags = item.flags()
            flags |= Qt.ItemFlag.ItemIsSelectable

            if col.key in {"id", "code"}:
                flags &= ~Qt.ItemFlag.ItemIsEditable
                flags &= ~Qt.ItemFlag.ItemIsUserCheckable
                item.setFlags(flags)
                continue

            if col.editor == "bool":
                flags &= ~Qt.ItemFlag.ItemIsEditable
                if self.edit_mode:
                    flags |= Qt.ItemFlag.ItemIsUserCheckable
                else:
                    flags &= ~Qt.ItemFlag.ItemIsUserCheckable
                item.setFlags(flags)
                continue

            if col.editor == "multi":
                flags &= ~Qt.ItemFlag.ItemIsEditable
                flags &= ~Qt.ItemFlag.ItemIsUserCheckable
                item.setFlags(flags)
                continue

            if col.editor == "color":
                flags &= ~Qt.ItemFlag.ItemIsEditable
                flags &= ~Qt.ItemFlag.ItemIsUserCheckable
                item.setFlags(flags)
                continue

            flags &= ~Qt.ItemFlag.ItemIsUserCheckable
            if self.edit_mode:
                flags |= Qt.ItemFlag.ItemIsEditable
            else:
                flags &= ~Qt.ItemFlag.ItemIsEditable
            item.setFlags(flags)

    def _update_editability(self) -> None:
        triggers = (
            QAbstractItemView.EditTrigger.AllEditTriggers
            if self.edit_mode
            else QAbstractItemView.EditTrigger.NoEditTriggers
        )
        self.table.setEditTriggers(triggers)
        for row in range(self.table.rowCount()):
            self._apply_row_flags(row)

    def toggle_edit_mode(self, enabled: bool) -> None:
        self.edit_mode = enabled
        if not enabled:
            self.quick_insert_mode = False
        self._update_editability()
        state = "attiva" if enabled else "disattiva"
        self.status_lbl.setText(f"Modalita modifica {state}.")

    def _handle_multi_selector(self, row: int, column: int) -> None:
        if not self.edit_mode:
            return
        col_cfg = self.columns[column]
        if col_cfg.editor == "color":
            item = self.table.item(row, column)
            current = item.text().strip() if item else ""
            color = QColorDialog.getColor(parent=self)
            if color.isValid():
                value = color.name().lower()
                if item is None:
                    item = QTableWidgetItem(value)
                    self.table.setItem(row, column, item)
                else:
                    item.setText(value)
            return
        if col_cfg.editor != "multi":
            return

        options = self.column_options.get(column, [])
        item = self.table.item(row, column)
        current_text = item.text().strip() if item else ""
        current_values = [part.strip() for part in current_text.split(",") if part.strip()]

        dialog = MultiSelectDialog(col_cfg.label, options, current_values, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            selected_text = ", ".join(dialog.selected_values())
            if item is None:
                item = QTableWidgetItem(selected_text)
                self.table.setItem(row, column, item)
            else:
                item.setText(selected_text)

    def add_row(self, from_arrow: bool = False) -> None:
        if not self.edit_mode:
            self.edit_btn.setChecked(True)

        if not from_arrow:
            self.quick_insert_mode = True

        self._append_row({column.key: "" for column in self.columns})
        new_row = self.table.rowCount() - 1
        target_col = self.table.currentColumn()
        if (
            target_col < 0
            or target_col >= len(self.columns)
            or self.columns[target_col].key in {"id", "code"}
            or not self.columns[target_col].visible
        ):
            target_col = self._first_input_column()
        self.table.setCurrentCell(new_row, target_col)
        self.table.scrollToBottom()
        if from_arrow:
            self.status_lbl.setText("Riga successiva creata (inserimento rapido).")
        else:
            self.status_lbl.setText("Nuova riga aggiunta. Compila i campi in tabella e premi Salva.")

    def eventFilter(self, watched, event):
        if (
            watched is self.table
            and event.type() == QEvent.Type.KeyPress
            and self.quick_insert_mode
            and self.edit_mode
            and event.key() == Qt.Key.Key_Down
        ):
            current_row = self.table.currentRow()
            if current_row >= 0 and current_row == self.table.rowCount() - 1:
                current_col = self.table.currentColumn()
                self.add_row(from_arrow=True)
                if current_col >= 0 and current_col < len(self.columns):
                    if self.columns[current_col].key not in {"id", "code"} and self.columns[current_col].visible:
                        self.table.setCurrentCell(self.table.rowCount() - 1, current_col)
                return True
        if watched is self.table and event.type() == QEvent.Type.KeyPress:
            if event.matches(QKeySequence.StandardKey.Paste):
                self._paste_from_clipboard()
                return True
        return super().eventFilter(watched, event)

    def _paste_from_clipboard(self) -> None:
        text = QGuiApplication.clipboard().text()
        if not text:
            return
        if not self.edit_mode:
            self.edit_btn.setChecked(True)

        lines = [line for line in text.splitlines() if line.strip() != ""]
        if not lines:
            return

        editable_cols = [
            idx
            for idx, col in enumerate(self.columns)
            if col.visible and col.key not in {"id", "code"}
        ]
        if not editable_cols:
            return

        start_row = self.table.currentRow()
        if start_row < 0:
            start_row = self.table.rowCount()

        start_col = self.table.currentColumn()
        if start_col < 0 or start_col not in editable_cols:
            start_col = editable_cols[0]

        start_col_index = editable_cols.index(start_col)

        for row_offset, line in enumerate(lines):
            values = line.split("\t")
            row_index = start_row + row_offset
            if row_index >= self.table.rowCount():
                self._append_row({column.key: "" for column in self.columns})

            for col_offset, value in enumerate(values):
                if start_col_index + col_offset >= len(editable_cols):
                    break
                col_idx = editable_cols[start_col_index + col_offset]
                col_cfg = self.columns[col_idx]
                item = self.table.item(row_index, col_idx)
                if col_cfg.editor == "bool":
                    if item is None:
                        item = QTableWidgetItem("")
                        self.table.setItem(row_index, col_idx, item)
                    item.setCheckState(
                        Qt.CheckState.Checked
                        if str(value).strip() in {"1", "si", "true", "yes", "y"}
                        else Qt.CheckState.Unchecked
                    )
                else:
                    if item is None:
                        item = QTableWidgetItem(str(value).strip())
                        self.table.setItem(row_index, col_idx, item)
                    else:
                        item.setText(str(value).strip())

        self.status_lbl.setText("Dati incollati dalla clipboard. Premi Salva.")

    def _validate_lookup_value(self, col_idx: int, value: str) -> None:
        col_cfg = self.columns[col_idx]
        if col_cfg.editor not in {"combo", "multi"}:
            return
        allowed = self.column_options.get(col_idx, [])
        if not allowed:
            return
        if col_cfg.editor == "combo":
            if value and value not in allowed:
                raise ValueError(
                    f"Campo '{col_cfg.label}': valore '{value}' non valido."
                )
            return
        for name in [part.strip() for part in value.split(",") if part.strip()]:
            if name not in allowed:
                raise ValueError(
                    f"Campo '{col_cfg.label}': valore '{name}' non valido."
                )

    def _row_to_payload(self, row_index: int) -> tuple[int | None, dict] | None:
        payload: dict[str, str] = {}
        row_id: int | None = None
        has_data = False

        for col_idx, col in enumerate(self.columns):
            item = self.table.item(row_index, col_idx)
            if col.editor == "bool":
                if item is None:
                    text = "0"
                else:
                    text = "1" if item.checkState() == Qt.CheckState.Checked else "0"
            else:
                text = (item.text() if item else "").strip()
            payload[col.key] = text
            self._validate_lookup_value(col_idx, text)
            if col.key == "id":
                if text:
                    row_id = int(text)
            elif text:
                has_data = True

        if row_id is None and not has_data:
            return None

        for col in self.columns:
            if col.required and col.key != "id" and not payload.get(col.key, "").strip():
                raise ValueError(f"Riga {row_index + 1}: campo '{col.label}' obbligatorio.")

        return row_id, payload

    def save_changes(self) -> None:
        if not self.edit_mode:
            QMessageBox.information(
                self,
                "Modifica non attiva",
                "Attiva Modifica per salvare le variazioni in tabella.",
            )
            return

        saved = 0
        for row in range(self.table.rowCount()):
            try:
                converted = self._row_to_payload(row)
                if converted is None:
                    continue
                row_id, payload = converted
                self.save_row(row_id, payload)
                saved += 1
            except ValueError as exc:
                QMessageBox.warning(self, "Errore salvataggio", str(exc))
                return

        self.refresh_page()
        self.status_lbl.setText(f"Salvate {saved} righe.")
        if self.on_change is not None and saved:
            self.on_change()

    def delete_selected(self) -> None:
        selected_rows = sorted(
            {idx.row() for idx in self.table.selectedIndexes()}, reverse=True
        )
        if not selected_rows:
            QMessageBox.information(self, "Elimina", "Seleziona almeno una riga.")
            return

        if (
            QMessageBox.question(
                self,
                "Conferma eliminazione",
                f"Eliminare {len(selected_rows)} righe selezionate?",
            )
            != QMessageBox.StandardButton.Yes
        ):
            return

        deleted = 0
        id_col_index = next(
            (index for index, col in enumerate(self.columns) if col.key == "id"), 0
        )
        for row in selected_rows:
            id_item = self.table.item(row, id_col_index)
            id_text = id_item.text().strip() if id_item else ""
            if id_text:
                try:
                    self.delete_row(int(id_text))
                    deleted += 1
                except ValueError as exc:
                    QMessageBox.warning(self, "Errore eliminazione", str(exc))
                    return
            else:
                self.table.removeRow(row)

        self.refresh_page()
        self.status_lbl.setText(f"Eliminate {deleted} righe.")
        if self.on_change is not None and deleted:
            self.on_change()

    def _update_help(self) -> None:
        if self.get_help_text is None:
            self.help_lbl.setText("")
            return
        self.help_lbl.setText(self.get_help_text())


class SettingsWindow(QWidget):
    data_changed = pyqtSignal()

    def __init__(self, repository, parent=None) -> None:
        super().__init__()
        self.repository = repository
        # Imposta l'icona anche per la finestra impostazioni (non solo QApplication).
        try:
            from pathlib import Path

            icon_path = Path(__file__).resolve().parent.parent / "assets" / "image.ico"
            if icon_path.exists():
                self.setWindowIcon(QIcon(str(icon_path)))
        except Exception:
            pass
        self.setWindowFlag(Qt.WindowType.Window, True)
        self.setWindowModality(Qt.WindowModality.NonModal)
        self._windows_vpn_names: list[str] = []
        self._vpn_worker: VpnListWorker | None = None
        self._vpn_page: EditableTablePage | None = None
        self._vpn_page_index: int | None = None
        self.setWindowTitle("Impostazioni - HD Manager")
        self.resize(1450, 900)

        self._build_ui()
        self._apply_style()

    def _build_ui(self) -> None:
        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self.menu = QListWidget()
        self.menu.setObjectName("sideMenu")
        self.menu.setFixedWidth(250)

        self.stack = QStackedWidget()
        self.stack.setObjectName("contentStack")

        root.addWidget(self.menu)
        root.addWidget(self.stack, 1)

        pages: list[tuple[str, QWidget]] = [
            ("Setup Tipi Prodotti", self._build_product_types_page()),
            ("Tag", self._build_tags_page()),
            ("Prodotti", self._build_products_page()),
            ("Ambienti", self._build_environments_page()),
            ("Release", self._build_releases_page()),
            ("Clienti", self._build_clients_page()),
            ("Risorse", self._build_resources_page()),
            ("Ruoli", self._build_roles_page()),
            ("VPN", self._build_vpns_page()),
            ("Pacchetti", self._build_packages_page()),
        ]

        for index, (title, page) in enumerate(pages):
            self.menu.addItem(QListWidgetItem(title))
            self.stack.addWidget(page)
            if title == "VPN" and isinstance(page, EditableTablePage):
                self._vpn_page = page
                self._vpn_page_index = index

        self.menu.currentRowChanged.connect(self._show_page)
        self.menu.setCurrentRow(0)
        self._refresh_windows_vpn_async()

    def _show_page(self, index: int) -> None:
        self.stack.setCurrentIndex(index)
        current = self.stack.currentWidget()
        if isinstance(current, EditableTablePage):
            current.refresh_page()
        if self._vpn_page_index is not None and index == self._vpn_page_index:
            self._refresh_windows_vpn_async()

    # Pages
    def _build_competences_page(self) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        return EditableTablePage(
            title="Competenze",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 420, required=True),
                TableColumn(
                    "resources",
                    "Risorse",
                    520,
                    editor="multi",
                    options_loader=lambda: [
                        row["label"] for row in repo.list_resources_lookup()
                    ],
                ),
            ],
            load_rows=repo.list_competences_with_resources,
            save_row=lambda row_id, data: self._save_competence_and_resources(repo, row_id, data),
            delete_row=repo.delete_competence,
            get_help_text=lambda: (
                "Inserire Competenze delle risorse. "
                "Puoi incollare righe da Excel/CSV con Ctrl+V."
            ),
            on_change=self._notify_data_changed,
        )

    def _save_competence_and_resources(
        self,
        repo,
        row_id: int | None,
        data: dict[str, str],
    ) -> int:
        competence_id = repo.upsert_competence(row_id, data["name"])
        raw = str(data.get("resources") or "").strip()
        labels = [part.strip() for part in raw.split(",") if part.strip()]
        repo.assign_resources_to_competence(data["name"], labels)
        return int(competence_id)

    def _build_product_types_page(self) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        return EditableTablePage(
            title="Setup Tipi Prodotti",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 260, required=True),
                TableColumn("flag_ip", "Flag IP", 100, editor="bool"),
                TableColumn("flag_host", "Flag HOST", 110, editor="bool"),
                TableColumn(
                    "flag_preconfigured",
                    "Flag PreConfigurata",
                    170,
                    editor="bool",
                ),
                TableColumn("flag_url", "Flag URL", 110, editor="bool"),
                TableColumn("flag_port", "Flag Porta", 110, editor="bool"),
            ],
            load_rows=lambda: [
                {
                    **row,
                    "flag_ip": str(row.get("flag_ip", 0)),
                    "flag_host": str(row.get("flag_host", 0)),
                    "flag_preconfigured": str(row.get("flag_preconfigured", 0)),
                    "flag_url": str(row.get("flag_url", 0)),
                    "flag_port": str(row.get("flag_port", 0)),
                }
                for row in repo.list_product_types()
            ],
            save_row=lambda row_id, data: repo.upsert_product_type(row_id, **data),
            delete_row=repo.delete_product_type,
            get_help_text=lambda: (
                "Inserisci il tipo di prodotto con i flag a seconda della tipologia e l'utilizzo. "
                "Ogni tipo determina le gestione degli accessi."
            ),
            on_change=self._notify_data_changed,
        )

    def _build_tags_page(self) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        return EditableTablePage(
            title="Tag",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 220, required=True),
                TableColumn("color", "Colore", 140, required=True, editor="color"),
                TableColumn(
                    "client_name",
                    "Cliente",
                    220,
                    editor="combo",
                    options_loader=lambda: [
                        row["label"] for row in repo.list_clients_lookup()
                    ],
                ),
            ],
            load_rows=lambda: [
                {
                    "id": row["id"],
                    "code": row["code"],
                    "name": row["name"],
                    "color": row.get("color") or "#0f766e",
                    "client_name": row.get("client_name") or "",
                }
                for row in repo.list_tags()
            ],
            save_row=lambda row_id, data: repo.upsert_tag(row_id, **data),
            delete_row=repo.delete_tag,
            get_help_text=lambda: (
                "Crea i tag da utilizzare e associare per le varie entita."
            ),
            on_change=self._notify_data_changed,
        )

    def _build_products_page(self) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        return EditableTablePage(
            title="Prodotti",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 220, required=True),
                TableColumn(
                    "product_type",
                    "Tipo Prodotto",
                    220,
                    required=True,
                    editor="combo",
                    options_loader=lambda: [row["label"] for row in repo.list_product_types_lookup()],
                ),
                TableColumn(
                    "clients",
                    "Clienti",
                    280,
                    editor="multi",
                    options_loader=lambda: [row["label"] for row in repo.list_clients_lookup()],
                ),
                TableColumn(
                    "environments",
                    "Ambienti",
                    280,
                    editor="multi",
                    options_loader=lambda: [
                        row["label"] for row in repo.list_environments_lookup()
                    ],
                ),
            ],
            load_rows=lambda: [
                {
                    "id": row["id"],
                    "code": row["code"],
                    "name": row["name"],
                    "product_type": row["product_type"],
                    "clients": row.get("clients") or "",
                    "environments": row.get("environments") or "",
                }
                for row in repo.list_products()
            ],
            save_row=lambda row_id, data: repo.upsert_product(row_id, **data),
            delete_row=repo.delete_product,
            get_help_text=lambda: (
                "Gestione dei tuoi prodotti."
            ),
            on_change=self._notify_data_changed,
        )

    def _build_environments_page(self) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        return EditableTablePage(
            title="Ambienti",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 320, required=True),
            ],
            load_rows=lambda: [
                {"id": row["id"], "code": row["code"], "name": row["name"]}
                for row in repo.list_environments()
            ],
            save_row=lambda row_id, data: repo.upsert_environment(row_id, **data),
            delete_row=repo.delete_environment,
            get_help_text=lambda: (
                "Crea gli ambienti da associare ai prodotti. "
                "Puoi incollare righe da Excel/CSV con Ctrl+V."
            ),
            on_change=self._notify_data_changed,
        )

    def _build_releases_page(self) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        return EditableTablePage(
            title="Release",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 320, required=True),
            ],
            load_rows=lambda: [
                {"id": row["id"], "code": row["code"], "name": row["name"]}
                for row in repo.list_releases()
            ],
            save_row=lambda row_id, data: repo.upsert_release(row_id, **data),
            delete_row=repo.delete_release,
            get_help_text=lambda: (
                "Crea la lista delle versioni da associare agli ambienti. "
                "Puoi incollare righe da Excel/CSV con Ctrl+V."
            ),
            on_change=self._notify_data_changed,
        )

    def _build_clients_page(self) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        return EditableTablePage(
            title="Clienti",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 220, required=True),
                TableColumn("location", "Localita", 180, required=True),
                TableColumn("link", "Link", 300),
                TableColumn(
                    "resources",
                    "Risorse",
                    340,
                    editor="multi",
                    options_loader=lambda: [
                        row["label"] for row in repo.list_resources_lookup()
                    ],
                ),
            ],
            load_rows=lambda: [
                {
                    "id": row["id"],
                    "code": row["code"],
                    "name": row["name"],
                    "location": row["location"],
                    "link": row.get("link") or "",
                    "resources": row.get("resources") or "",
                }
                for row in repo.list_clients()
            ],
            save_row=lambda row_id, data: repo.upsert_client(row_id, **data),
            delete_row=repo.delete_client,
            get_help_text=lambda: (
                "Creazione e Gestione dei Clienti. "
                "Puoi incollare righe da Excel/CSV con Ctrl+V."
            ),
            on_change=self._notify_data_changed,
        )

    def _build_resources_page(self) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        return EditableTablePage(
            title="Risorse",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 150, required=True),
                TableColumn("surname", "Cognome", 170, required=True),
                TableColumn(
                    "role_name",
                    "Ruolo",
                    200,
                    required=True,
                    editor="combo",
                    options_loader=lambda: [row["label"] for row in repo.list_roles_lookup()],
                ),
                TableColumn("phone", "Telefono", 150),
                TableColumn("email", "Email", 240),
                TableColumn("note", "Note", 260),
            ],
            load_rows=lambda: [
                {
                    "id": row["id"],
                    "code": row["code"],
                    "name": row["name"],
                    "surname": row["surname"],
                    "role_name": row.get("role_name") or "",
                    "phone": row.get("phone") or "",
                    "email": row.get("email") or "",
                    "note": row.get("note") or "",
                }
                for row in repo.list_resources()
            ],
            save_row=lambda row_id, data: repo.upsert_resource(row_id, **data),
            delete_row=repo.delete_resource,
            get_help_text=lambda: "Creazione e Gestione delle risorse della tua azienda.",
            on_change=self._notify_data_changed,
        )

    def _build_roles_page(self) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        return EditableTablePage(
            title="Ruoli",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("name", "Nome", 220, required=True),
                TableColumn(
                    "display_order",
                    "Ordine visualizzazione",
                    190,
                    required=False,
                ),
                TableColumn(
                    "multi_clients",
                    "Più di uno sullo stesso cliente",
                    140,
                    editor="bool",
                ),
            ],
            load_rows=lambda: [
                {
                    "id": row["id"],
                    "code": row["code"],
                    "name": row["name"],
                    "display_order": str(row.get("display_order", "") or ""),
                    "multi_clients": str(row.get("multi_clients", 0)),
                }
                for row in repo.list_roles()
            ],
            save_row=lambda row_id, data: repo.upsert_role(row_id, **data),
            delete_row=repo.delete_role,
            get_help_text=lambda: (
                "Creazioni dei Ruoli da associare alle tuo risorse es PM, Consulenti... "
                "Puoi incollare righe da Excel/CSV con Ctrl+V."
            ),
            on_change=self._notify_data_changed,
        )

    def _build_vpns_page(self) -> EditableTablePage:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        return EditableTablePage(
            title="VPN",
            columns=[
                TableColumn("id", "ID", 70, visible=False),
                TableColumn("code", "Codice", 150),
                TableColumn("connection_name", "Nome Connessione", 220, required=True),
                TableColumn("server_address", "Indirizzo Server", 220, required=True),
                TableColumn(
                    "vpn_type",
                    "Tipo VPN",
                    150,
                    required=True,
                    editor="combo",
                    options_loader=lambda: ["Vpn Proprietario", "VPN Windows"],
                ),
                TableColumn(
                    "vpn_windows_name",
                    "VPN Windows",
                    200,
                    editor="combo",
                    options_loader=lambda: list(self._windows_vpn_names),
                ),
                TableColumn("access_info_type", "Tipo Info Accesso", 170),
                TableColumn("username", "Nome Utente", 160, required=True),
                TableColumn("password", "Password", 150, required=True),
                TableColumn("vpn_path", "Percorso VPN", 220),
                TableColumn(
                    "clients",
                    "Clienti",
                    260,
                    editor="multi",
                    options_loader=lambda: [
                        row["label"] for row in repo.list_clients_lookup()
                    ],
                ),
            ],
            load_rows=lambda: [
                {
                    "id": row["id"],
                    "code": row["code"],
                    "connection_name": row["connection_name"],
                    "server_address": row["server_address"],
                    "vpn_type": row["vpn_type"],
                    "vpn_windows_name": row["connection_name"]
                    if row.get("vpn_type") == "VPN Windows"
                    else "",
                    "access_info_type": row.get("access_info_type") or "",
                    "username": row["username"],
                    "password": row["password"],
                    "vpn_path": row.get("vpn_path") or "",
                    "clients": row.get("clients") or "",
                }
                for row in repo.list_vpns()
            ],
            save_row=lambda row_id, data: self._save_vpn_with_clients(repo, row_id, data),
            delete_row=repo.delete_vpn,
            get_help_text=lambda: "Gestione delle VPN da associare al tuo cliente.",
            on_change=self._notify_data_changed,
        )

    def _build_packages_page(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(14)

        base_repo = (
            self.repository.repository
            if hasattr(self.repository, "repository")
            else self.repository
        )
        packaging = PackagingService(base_repo)

        layout.addWidget(
            QLabel(
                "Qui puoi esportare/importare configurazioni predefinite.\n"
                "I pacchetti sono in formato JSON e sono pensati per importare su un altro PC.\n\n"
                "Nota: le relazioni dei Clienti (VPN + Risorse) vengono risolte automaticamente quando importi "
                "Core + Risorse + VPN."
            )
        )

        def _export(do_export):
            path, _ = QFileDialog.getSaveFileName(
                self,
                "Salva pacchetto",
                "hdmanager_package.json",
                "JSON (*.json)",
            )
            if not path:
                return
            try:
                do_export(path)
                QMessageBox.information(self, "Pacchetto esportato", f"Salvato in:\n{path}")
            except Exception as exc:  # pragma: no cover - UI only
                QMessageBox.critical(self, "Errore export", str(exc))

        def _import():
            path, _ = QFileDialog.getOpenFileName(
                self,
                "Importa pacchetto",
                "",
                "JSON (*.json)",
            )
            if not path:
                return
            try:
                packaging.import_package(path)
                QMessageBox.information(self, "Pacchetto importato", "Import completato.")
            except Exception as exc:  # pragma: no cover - UI only
                QMessageBox.critical(self, "Errore import", str(exc))

        export_row = QHBoxLayout()
        export_row.setSpacing(10)
        btn_export_core = QPushButton("Esporta Core")
        btn_export_resources = QPushButton("Esporta Risorse")
        btn_export_vpns = QPushButton("Esporta VPN")
        export_row.addWidget(btn_export_core)
        export_row.addWidget(btn_export_resources)
        export_row.addWidget(btn_export_vpns)
        layout.addLayout(export_row)

        btn_export_core.clicked.connect(lambda: _export(packaging.export_core))
        btn_export_resources.clicked.connect(lambda: _export(packaging.export_resources))
        btn_export_vpns.clicked.connect(lambda: _export(packaging.export_vpns))

        import_row = QHBoxLayout()
        import_row.setSpacing(10)
        btn_import_core = QPushButton("Importa Core")
        btn_import_resources = QPushButton("Importa Risorse")
        btn_import_vpns = QPushButton("Importa VPN")
        import_row.addWidget(btn_import_core)
        import_row.addWidget(btn_import_resources)
        import_row.addWidget(btn_import_vpns)
        layout.addLayout(import_row)

        btn_import_core.clicked.connect(_import)
        btn_import_resources.clicked.connect(_import)
        btn_import_vpns.clicked.connect(_import)

        layout.addStretch(1)
        return widget

    def _refresh_windows_vpn_async(self) -> None:
        if self._vpn_worker is not None and self._vpn_worker.isRunning():
            return
        self._vpn_worker = VpnListWorker(self.repository)
        self._vpn_worker.finished.connect(self._on_windows_vpn_loaded)
        self._vpn_worker.start()

    def _on_windows_vpn_loaded(self, names: list[str]) -> None:
        self._windows_vpn_names = names
        if self._vpn_page is not None:
            self._vpn_page.refresh_page()

    def _apply_style(self) -> None:
        self.setStyleSheet(
            """
            QWidget {
                font-family: "Segoe UI";
                font-size: 12px;
                color: #1e293b;
                background: #f3f7fa;
            }
            #sideMenu {
                background: #0f172a;
                color: #e2e8f0;
                border: none;
                padding: 8px;
            }
            #sideMenu::item {
                padding: 10px 12px;
                margin: 4px 6px;
                border-radius: 8px;
            }
            #sideMenu::item:selected {
                background: #2563eb;
                color: #ffffff;
            }
            #contentStack {
                background: #ffffff;
                border-left: 1px solid #dbe5ee;
            }
            #pageTitle {
                color: #0f172a;
            }
            #helpLabel {
                color: #0c4a6e;
                background: #e0f2fe;
                border: 1px solid #bae6fd;
                border-radius: 8px;
                padding: 8px;
            }
            QTableWidget {
                background: #ffffff;
                border: 1px solid #cbd5e1;
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
            QPushButton {
                background: #ffffff;
                border: 1px solid #94a3b8;
                border-radius: 8px;
                padding: 6px 12px;
                font-weight: 600;
                color: #0f172a;
            }
            QPushButton:hover {
                background: #e0f2fe;
                border-color: #0284c7;
            }
            QPushButton:checked {
                background: #2563eb;
                color: #ffffff;
                border: 1px solid #1d4ed8;
            }
            #primaryActionButton {
                background: #2563eb;
                color: #ffffff;
                border: 1px solid #1d4ed8;
                font-weight: 700;
            }
            #primaryActionButton:hover {
                background: #1d4ed8;
                border-color: #1e40af;
            }
            #dangerActionButton {
                background: #dc2626;
                color: #ffffff;
                border: 1px solid #b91c1c;
                font-weight: 700;
            }
            #dangerActionButton:hover {
                background: #b91c1c;
                border-color: #991b1b;
            }
            #secondaryActionButton {
                background: #ffffff;
                color: #0f172a;
                border: 1px solid #94a3b8;
                font-weight: 600;
            }
            #secondaryActionButton:hover {
                background: #e0f2fe;
                border-color: #0284c7;
            }
            #statusLabel {
                color: #334155;
            }
            """
        )

    def _notify_data_changed(self) -> None:
        self.data_changed.emit()

    def _save_vpn_with_clients(self, repo, row_id: int | None, data: dict) -> int:
        """Salva VPN assicurando che il campo Clienti sia sempre passato correttamente."""
        clients_raw = data.get("clients")
        if clients_raw is None:
            clients_val = ""
        elif isinstance(clients_raw, list):
            clients_val = ", ".join(str(x).strip() for x in clients_raw if str(x).strip())
        else:
            clients_val = str(clients_raw).strip() if clients_raw else ""
        payload = dict(data)
        payload["clients"] = clients_val
        return repo.upsert_vpn(row_id, **payload)


class VpnListWorker(QThread):
    finished = pyqtSignal(list)

    def __init__(self, repository, parent=None) -> None:
        super().__init__(parent)
        self.repository = repository

    def run(self) -> None:
        repo = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        names = repo.list_windows_vpn_connections()
        self.finished.emit(names)
