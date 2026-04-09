"""SQL: conservazione query SQL e editor (linguetta principale)."""

from __future__ import annotations

import re
import uuid
from datetime import date
from pathlib import Path
from typing import Any

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QStackedWidget,
    QToolBox,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from app.sql_query_tables import (
    apply_alias_replacements,
    apply_gestione_replacements,
    extract_sql_alias_suggestions,
    extract_sql_qualified_column_refs,
    extract_sql_table_names,
)
from app.ui.ui_constants import SIDE_MENU_TAB_CONTENT_MARGINS, SIDE_MENU_WIDTH_PX


class SqlArchiveWorkspaceWidget(QWidget):
    """Lista query salvate + menu laterale Nuova / Gestione."""

    def __init__(self, main_window: QWidget, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._mw = main_window
        self._editing_id: str | None = None
        self._gestione_inputs: dict[str, QLineEdit] = {}
        self._gestione_column_labels: list[QLabel] = []
        self._build_ui()
        self._reload_query_list()

    @staticmethod
    def _gestione_mono_font() -> QFont:
        mono = QFont("Consolas", 10)
        if not mono.exactMatch():
            mono = QFont("Courier New", 10)
        return mono

    @staticmethod
    def _gestione_qualified_field_caption(ref: dict) -> str:
        """Una sola riga: NomeAlias.NomeCampo (campo come nel testo, es. [Document No_])."""
        qual = (ref.get("qualifier") or "").strip()
        col_part = (ref.get("column_raw") or ref.get("column") or "").strip()
        return f"{qual}.{col_part}" if qual or col_part else ""

    def _polish_gestione_toolbox_buttons(self) -> None:
        """Assicura che lo stylesheet colori i titoli sezione (su Windows i QToolButton ignorano spesso il bg)."""
        for btn in self._gestione_toolbox.findChildren(QToolButton):
            btn.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)

    def _apply_gestione_rilevato_label_metrics(self, lab: QLabel) -> None:
        """
        Con word-wrap attivo la colonna «Rilevato» resta troppo stretta e il testo viene tagliato.
        Una riga + larghezza minima dal testo ripristina la lettura completa.
        """
        lab.setWordWrap(False)
        lab.setSizePolicy(QSizePolicy.Policy.MinimumExpanding, QSizePolicy.Policy.Preferred)
        fm = lab.fontMetrics()
        tw = fm.horizontalAdvance(lab.text()) + 20
        lab.setMinimumWidth(min(720, max(240, tw)))

    @staticmethod
    def _gestione_pair_headers() -> tuple[QLabel, QLabel]:
        hdr_a = QLabel("Rilevato")
        hdr_b = QLabel("Sostituisci con")
        for h in (hdr_a, hdr_b):
            h.setStyleSheet(
                "font-weight: 700; font-size: 10px; color: #64748b; letter-spacing: 0.04em;"
            )
        return hdr_a, hdr_b

    def _settings(self):
        repo = getattr(self._mw, "repository", None)
        if repo is not None and hasattr(repo, "settings"):
            return repo.settings
        return None

    def _controller(self):
        return getattr(self._mw, "repository", None)

    def _data_repository(self):
        """Repository dati (RepositoryService), per es. upsert tag."""
        ctl = self._controller()
        if ctl is None:
            return None
        return getattr(ctl, "repository", None)

    def _archive_service(self):
        ctl = self._controller()
        if ctl is None:
            return None
        return getattr(ctl, "archive", None)

    @staticmethod
    def _sanitize_filename_base(name: str) -> str:
        s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", (name or "").strip() or "Senza_nome")
        s = re.sub(r"\s+", "_", s).strip("._")
        return (s or "Senza_nome")[:180]

    def _ensure_query_tag(self) -> None:
        repo = self._data_repository()
        if repo is not None:
            repo.upsert_tag(None, "Query", "#2563eb", "")

    def _write_query_txt_and_archive(self, display_name: str, sql_text: str) -> Path:
        """Scrive `{nome}_Query_{data}.txt` nella Directory Salvataggi e registra in Archivio."""
        s = self._settings()
        if s is None:
            raise RuntimeError("Impostazioni non disponibili.")
        out_dir = s.get_calculator_export_base_path()
        safe = self._sanitize_filename_base(display_name)
        day = date.today().strftime("%Y-%m-%d")
        base = f"{safe}_Query_{day}.txt"
        path = out_dir / base
        n = 1
        while path.exists():
            path = out_dir / f"{safe}_Query_{day}_{n}.txt"
            n += 1
        path.write_text(sql_text, encoding="utf-8")
        arch = self._archive_service()
        if arch is None:
            raise RuntimeError("Servizio archivio non disponibile.")
        arch.add_file(None, str(path.resolve()), "Query")
        return path

    def _build_ui(self) -> None:
        self.setObjectName("sqlWorkspacePage")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        root = QHBoxLayout(self)
        root.setContentsMargins(*SIDE_MENU_TAB_CONTENT_MARGINS)
        root.setSpacing(0)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setObjectName("sqlArchiveSplitter")

        self._sql_mode_menu = QListWidget()
        self._sql_mode_menu.setObjectName("sqlModeSideMenu")
        self._sql_mode_menu.setFixedWidth(SIDE_MENU_WIDTH_PX)
        self._sql_mode_menu.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._sql_mode_menu.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._sql_mode_menu.addItem(QListWidgetItem("Nuova"))
        self._sql_mode_menu.addItem(QListWidgetItem("Gestione"))

        mode_col = QFrame()
        mode_col.setObjectName("sqlModeColumn")
        mode_col.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        mode_l = QVBoxLayout(mode_col)
        mode_l.setContentsMargins(0, 0, 8, 0)
        mode_l.setSpacing(0)
        mode_l.addWidget(self._sql_mode_menu, 1)

        left = QFrame()
        left.setObjectName("sqlArchiveListPane")
        left.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        left_l = QVBoxLayout(left)
        left_l.setContentsMargins(12, 12, 12, 12)
        left_l.setSpacing(8)
        list_title = QLabel("Query salvate")
        list_title.setObjectName("subSectionTitle")
        self._query_list = QListWidget()
        self._query_list.setObjectName("sqlArchiveQueryList")
        self._query_list.setMinimumWidth(SIDE_MENU_WIDTH_PX)
        self._query_list.setAlternatingRowColors(True)
        self._query_list.currentItemChanged.connect(self._on_list_selection_changed)
        left_l.addWidget(list_title)
        left_l.addWidget(self._query_list, 1)
        list_btns = QHBoxLayout()
        self._btn_delete = QPushButton("Elimina dalla libreria")
        self._btn_delete.setObjectName("dangerActionButton")
        self._btn_delete.setToolTip("Rimuove la query selezionata dall’elenco salvato.")
        self._btn_delete.clicked.connect(self._on_delete_query)
        list_btns.addWidget(self._btn_delete)
        list_btns.addStretch(1)
        left_l.addLayout(list_btns)

        right = QFrame()
        right.setObjectName("sqlArchiveEditorPane")
        right_l = QVBoxLayout(right)
        right_l.setContentsMargins(0, 0, 0, 0)
        right_l.setSpacing(0)

        self._sql_editor_stack = QStackedWidget()
        self._sql_editor_stack.setObjectName("sqlEditorStack")

        nuova = QWidget()
        nuova_l = QVBoxLayout(nuova)
        nuova_l.setContentsMargins(0, 0, 0, 0)
        nuova_l.setSpacing(10)

        name_row = QHBoxLayout()
        nl = QLabel("Nome:")
        nl.setObjectName("subText")
        name_row.addWidget(nl)
        self._name_edit = QLineEdit()
        self._name_edit.setObjectName("archiveFilter")
        self._name_edit.setPlaceholderText("Nome della query (es. Elenco clienti attivi)")
        name_row.addWidget(self._name_edit, 1)
        nuova_l.addLayout(name_row)

        sql_label = QLabel("Testo SQL")
        sql_label.setObjectName("sectionTitle")
        nuova_l.addWidget(sql_label)

        self._sql_edit = QPlainTextEdit()
        self._sql_edit.setObjectName("sqlArchiveSqlEdit")
        self._sql_edit.setPlaceholderText(
            "Scrivi qui la query SQL (solo lettura consigliata; l’esecuzione sarà in una fase successiva)."
        )
        mono = QFont("Consolas", 10)
        if not mono.exactMatch():
            mono = QFont("Courier New", 10)
        self._sql_edit.setFont(mono)
        self._sql_edit.setTabStopDistance(self._sql_edit.fontMetrics().horizontalAdvance(" ") * 4)
        nuova_l.addWidget(self._sql_edit, 1)

        btn_row = QHBoxLayout()
        self._btn_save = QPushButton("Salva nella libreria")
        self._btn_save.setObjectName("primaryActionButton")
        self._btn_save.setToolTip(
            "Salva nel database, crea un file .txt nella Directory Salvataggi (Setup Strumenti) "
            "con nome Nome_Query_Data, e registra il file in Archivio con tag «Query»."
        )
        self._btn_clear = QPushButton("Nuova query vuota")
        self._btn_clear.setObjectName("secondaryActionButton")
        self._btn_save.clicked.connect(self._on_save_query)
        self._btn_clear.clicked.connect(self._on_new_blank)
        btn_row.addWidget(self._btn_save)
        btn_row.addWidget(self._btn_clear)
        btn_row.addStretch(1)
        nuova_l.addLayout(btn_row)

        self._gestione_page = QWidget()
        gest_l = QVBoxLayout(self._gestione_page)
        gest_l.setContentsMargins(12, 12, 12, 12)
        gest_l.setSpacing(10)
        g_title = QLabel("Adattamento SQL")
        g_title.setObjectName("sectionTitle")
        g_intro = QLabel(
            "Analizza il testo SQL (scheda Nuova): tabelle, alias di tabella, campi qualificati "
            "(solo il nome dopo il punto; l’alias si modifica sopra). "
            "La colonna «Sostituisci con» è precompilata con i valori rilevati, così puoi modificarli senza riscriverli. "
            "Indicazione se il campo compare anche in WHERE/HAVING. "
            "Ordine applicato: nome campo → tabelle → alias. "
            "Le modifiche si applicano al testo nell’editor; usa «Salva nella libreria» in Nuova per memorizzare."
        )
        g_intro.setObjectName("subText")
        g_intro.setWordWrap(True)
        self._gestione_empty = QLabel("Seleziona una query nella lista oppure scrivi SQL in Nuova.")
        self._gestione_empty.setObjectName("subText")
        self._gestione_empty.setWordWrap(True)

        btn_row_g = QHBoxLayout()
        self._btn_gestione_analyze = QPushButton("Rianalizza dal testo")
        self._btn_gestione_analyze.setObjectName("secondaryActionButton")
        self._btn_gestione_analyze.setToolTip(
            "Rileva di nuovo le tabelle dal contenuto attuale dell’editor SQL (scheda Nuova)."
        )
        self._btn_gestione_apply = QPushButton("Applica sostituzioni")
        self._btn_gestione_apply.setObjectName("primaryActionButton")
        self._btn_gestione_apply.setToolTip(
            "Applica nel testo SQL: solo nomi campo, poi tabelle, poi alias (vedi ordine nelle istruzioni)."
        )
        self._btn_gestione_analyze.clicked.connect(self._refresh_gestione_mapping)
        self._btn_gestione_apply.clicked.connect(self._on_apply_table_replacements)
        btn_row_g.addWidget(self._btn_gestione_analyze)
        btn_row_g.addWidget(self._btn_gestione_apply)
        btn_row_g.addStretch(1)

        self._gestione_scroll = QScrollArea()
        self._gestione_scroll.setWidgetResizable(True)
        self._gestione_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self._gestione_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._gestione_toolbox = QToolBox()
        self._gestione_toolbox.setObjectName("sqlGestioneToolBox")
        self._gestione_toolbox.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self._gestione_scroll.setWidget(self._gestione_toolbox)

        gest_l.addWidget(g_title)
        gest_l.addWidget(g_intro)
        gest_l.addWidget(self._gestione_empty)
        gest_l.addLayout(btn_row_g)
        gest_l.addWidget(self._gestione_scroll, 1)
        self._gestione_scroll.setVisible(False)
        self._btn_gestione_apply.setEnabled(False)

        self._sql_editor_stack.addWidget(nuova)
        self._sql_editor_stack.addWidget(self._gestione_page)

        self._sql_mode_menu.currentRowChanged.connect(self._on_sql_mode_row_changed)

        right_l.addWidget(self._sql_editor_stack, 1)

        self._sql_mode_menu.setCurrentRow(0)

        splitter.addWidget(mode_col)
        splitter.addWidget(left)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 0)
        splitter.setStretchFactor(2, 1)
        splitter.setSizes([SIDE_MENU_WIDTH_PX, 260, 720])

        root.addWidget(splitter, 1)

        self.setStyleSheet(
            """
            #sqlArchiveSqlEdit {
                background: #f8fafc;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                padding: 10px;
                color: #0f172a;
            }
            #sqlArchiveSqlEdit:focus {
                border-color: #2563eb;
                background: #ffffff;
            }
            #sqlGestioneToolBox {
                background: #f4f7f9;
                border: 1px solid #dbe5ee;
                border-radius: 12px;
                padding: 10px;
            }
            /*
             * Intestazioni sezione: colorate come il resto dell’app (cyan/teal + blu primario).
             * Doppio selettore: QToolButton copre i pulsanti interni su tutti i temi.
             */
            #sqlGestioneToolBox QToolBoxButton,
            #sqlGestioneToolBox QToolButton {
                min-width: 280px;
                min-height: 44px;
                padding: 12px 16px 12px 12px;
                margin-bottom: 8px;
                background: #e0f2fe;
                border: 1px solid #7dd3fc;
                border-left: 5px solid #0ea5e9;
                border-radius: 10px;
                color: #0369a1;
                font-size: 13px;
                font-weight: 700;
                text-align: left;
            }
            #sqlGestioneToolBox QToolBoxButton:hover,
            #sqlGestioneToolBox QToolButton:hover {
                background: #bae6fd;
                border-color: #38bdf8;
                border-left-color: #0284c7;
                color: #0c4a6e;
            }
            #sqlGestioneToolBox QToolBoxButton:checked,
            #sqlGestioneToolBox QToolButton:checked {
                background: #2563eb;
                border: 1px solid #1d4ed8;
                border-left: 5px solid #1e40af;
                color: #ffffff;
            }
            #sqlGestioneToolBox QToolBoxButton:checked:hover,
            #sqlGestioneToolBox QToolButton:checked:hover {
                background: #1d4ed8;
                border-color: #1e40af;
                color: #ffffff;
            }
            #sqlGestioneSectionPage {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 10px;
                margin-top: 4px;
            }
            #sqlArchiveReplaceEdit {
                background: #ffffff;
                border: 1px solid #cbd5e1;
                border-radius: 8px;
                padding: 6px 10px;
                color: #0f172a;
            }
            #sqlArchiveReplaceEdit:focus {
                border-color: #2563eb;
                background: #ffffff;
            }
            """
        )

    def _queries_by_id(self) -> dict[str, dict[str, Any]]:
        s = self._settings()
        if s is None:
            return {}
        return {str(q["id"]): q for q in s.get_sql_archive_queries()}

    def _reload_query_list(self) -> None:
        self._query_list.blockSignals(True)
        self._query_list.clear()
        s = self._settings()
        rows: list[dict[str, Any]] = s.get_sql_archive_queries() if s else []
        for q in rows:
            item = QListWidgetItem(q["name"])
            item.setData(Qt.ItemDataRole.UserRole, q["id"])
            item.setToolTip(q["name"])
            self._query_list.addItem(item)
        self._query_list.blockSignals(False)

    def _on_list_selection_changed(self, current: QListWidgetItem | None, _previous) -> None:
        if current is None:
            return
        qid = current.data(Qt.ItemDataRole.UserRole)
        if not qid:
            return
        q = self._queries_by_id().get(str(qid))
        if not q:
            return
        self._editing_id = str(q["id"])
        self._name_edit.setText(q["name"])
        self._sql_edit.setPlainText(q["sql_text"])
        self._sql_mode_menu.blockSignals(True)
        self._sql_mode_menu.setCurrentRow(0)
        self._sql_editor_stack.setCurrentIndex(0)
        self._sql_mode_menu.blockSignals(False)

    def _on_sql_mode_row_changed(self, row: int) -> None:
        if row < 0 or row >= self._sql_editor_stack.count():
            return
        self._sql_editor_stack.setCurrentIndex(row)
        if row == 1:
            self._refresh_gestione_mapping()

    def _clear_gestione_form(self) -> None:
        self._gestione_inputs.clear()
        self._gestione_column_labels.clear()
        while self._gestione_toolbox.count() > 0:
            self._gestione_toolbox.removeItem(0)

    def _refresh_gestione_mapping(self) -> None:
        sql = self._sql_edit.toPlainText().strip()
        prev_tab = self._gestione_toolbox.currentIndex() if self._gestione_toolbox.count() else 0
        self._clear_gestione_form()
        if not sql:
            self._gestione_empty.setText(
                "Nessun testo SQL nell’editor. Apri una query dalla lista o incolla SQL in Nuova."
            )
            self._gestione_empty.setVisible(True)
            self._gestione_scroll.setVisible(False)
            self._btn_gestione_apply.setEnabled(False)
            return
        tables = extract_sql_table_names(sql)
        cols = extract_sql_qualified_column_refs(sql)
        aliases = extract_sql_alias_suggestions(sql)
        if not tables and not cols and not aliases:
            self._gestione_empty.setText(
                "Nessun elemento rilevato (tabelle FROM/JOIN, campi qualificati tipo Alias.campo o "
                "[Tab].[Campo], alias). Verifica la sintassi o usa «Rianalizza» dopo aver modificato il testo."
            )
            self._gestione_empty.setVisible(True)
            self._gestione_scroll.setVisible(False)
            self._btn_gestione_apply.setEnabled(False)
            return
        self._gestione_empty.setVisible(False)
        self._gestione_scroll.setVisible(True)
        self._btn_gestione_apply.setEnabled(True)

        mono = self._gestione_mono_font()

        if tables:
            page_t = QWidget()
            page_t.setObjectName("sqlGestioneSectionPage")
            vl_t = QVBoxLayout(page_t)
            vl_t.setContentsMargins(12, 12, 12, 12)
            vl_t.setSpacing(10)
            form_t = QFormLayout()
            form_t.setSpacing(10)
            form_t.setHorizontalSpacing(16)
            form_t.setLabelAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            form_t.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
            ha, hb = self._gestione_pair_headers()
            form_t.addRow(ha, hb)
            for t in tables:
                lab = QLabel(t)
                lab.setFont(mono)
                self._apply_gestione_rilevato_label_metrics(lab)
                edit = QLineEdit()
                edit.setText(t)
                edit.setPlaceholderText("Modifica o lascia vuoto per non sostituire questa tabella")
                edit.setFont(mono)
                edit.setObjectName("sqlArchiveReplaceEdit")
                edit.setMinimumWidth(200)
                edit.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
                self._gestione_inputs[f"t:{t}"] = edit
                form_t.addRow(lab, edit)
            vl_t.addLayout(form_t)
            self._gestione_toolbox.addItem(page_t, "\u25b8  Tabelle")

        if aliases:
            page_a = QWidget()
            page_a.setObjectName("sqlGestioneSectionPage")
            vl_a = QVBoxLayout(page_a)
            vl_a.setContentsMargins(4, 8, 4, 8)
            vl_a.setSpacing(10)
            hint_a = QLabel(
                "Sostituisce l’alias ovunque compaia (anche in «Alias.campo»). "
                "I campi sono precompilati con l’alias attuale: puoi correggere solo una parte o aggiungere testo. "
                "All’uscita dal campo (Tab / Invio) la sezione Campi si aggiorna in anteprima."
            )
            hint_a.setObjectName("subText")
            hint_a.setWordWrap(True)
            vl_a.addWidget(hint_a)
            form_a = QFormLayout()
            form_a.setSpacing(10)
            form_a.setHorizontalSpacing(16)
            form_a.setLabelAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            form_a.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
            ha, hb = self._gestione_pair_headers()
            form_a.addRow(ha, hb)
            for a in aliases:
                lab = QLabel(a)
                lab.setFont(mono)
                self._apply_gestione_rilevato_label_metrics(lab)
                edit = QLineEdit()
                edit.setText(a)
                edit.setPlaceholderText("Modifica o lascia vuoto per non sostituire questo alias")
                edit.setFont(mono)
                edit.setObjectName("sqlArchiveReplaceEdit")
                self._gestione_inputs[f"a:{a}"] = edit
                edit.editingFinished.connect(self._on_gestione_alias_committed)
                edit.returnPressed.connect(self._on_gestione_alias_committed)
                edit.setMinimumWidth(200)
                edit.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
                form_a.addRow(lab, edit)
            vl_a.addLayout(form_a)
            self._gestione_toolbox.addItem(page_a, "\u25b8  Alias di tabella")

        if cols:
            page_c = QWidget()
            page_c.setObjectName("sqlGestioneSectionPage")
            vl_c = QVBoxLayout(page_c)
            vl_c.setContentsMargins(12, 12, 12, 12)
            vl_c.setSpacing(10)
            hint_c = QLabel(
                "Modifica solo la parte dopo il punto (es. [Document No_]). "
                "Ogni campo è precompilato con il nome colonna attuale: puoi editarlo in loco senza riscriverlo da zero. "
                "L’alias a sinistra va cambiato nella sezione «Alias di tabella». "
                "Ordine applicato: nome campo → tabelle → alias."
            )
            hint_c.setObjectName("subText")
            hint_c.setWordWrap(True)
            vl_c.addWidget(hint_c)
            form_c = QFormLayout()
            form_c.setSpacing(10)
            form_c.setHorizontalSpacing(16)
            form_c.setLabelAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
            form_c.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.ExpandingFieldsGrow)
            ha, hb = self._gestione_pair_headers()
            form_c.addRow(ha, hb)
            for ref in cols:
                col_raw = ref.get("column_raw") or ref["sample"].split(".", 1)[-1].strip()
                cap = self._gestione_qualified_field_caption(ref)
                lab = QLabel(cap)
                lab.setFont(mono)
                if ref.get("in_where_having"):
                    lab.setToolTip("Presente anche in WHERE / HAVING.")
                self._apply_gestione_rilevato_label_metrics(lab)
                self._gestione_column_labels.append(lab)
                edit = QLineEdit()
                edit.setText(col_raw)
                edit.setPlaceholderText("Modifica il nome campo o lascia vuoto per non sostituirlo")
                edit.setFont(mono)
                edit.setObjectName("sqlArchiveReplaceEdit")
                edit.setProperty("gestioneColumnRaw", col_raw)
                edit.setMinimumWidth(200)
                edit.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
                self._gestione_inputs[f"c:{ref['key']}"] = edit
                form_c.addRow(lab, edit)
            vl_c.addLayout(form_c)
            self._gestione_toolbox.addItem(page_c, "\u25b8  Campi Qualificati")

        n_tabs = self._gestione_toolbox.count()
        if n_tabs > 0:
            self._gestione_toolbox.setCurrentIndex(min(max(0, prev_tab), n_tabs - 1))
            self._polish_gestione_toolbox_buttons()

    def _on_gestione_alias_committed(self) -> None:
        """Aggiorna le etichette della sezione Campi con l’anteprima degli alias (senza scrivere nell’editor)."""
        if not self._gestione_column_labels:
            return
        sql = self._sql_edit.toPlainText()
        if not sql.strip():
            return
        pending: dict[str, str] = {}
        for k, edit in self._gestione_inputs.items():
            if not k.startswith("a:"):
                continue
            old_a = k[2:]
            nv = edit.text().strip()
            if nv and nv != old_a:
                pending[old_a] = nv
        sim = apply_alias_replacements(sql, pending) if pending else sql
        cols = extract_sql_qualified_column_refs(sim)
        mono = self._gestione_mono_font()
        for i, lbl in enumerate(self._gestione_column_labels):
            if i >= len(cols):
                break
            ref = cols[i]
            cap = self._gestione_qualified_field_caption(ref)
            lbl.setText(cap)
            lbl.setToolTip("Presente anche in WHERE / HAVING." if ref.get("in_where_having") else "")
            lbl.setFont(mono)
            self._apply_gestione_rilevato_label_metrics(lbl)

    def _on_apply_table_replacements(self) -> None:
        sql = self._sql_edit.toPlainText()
        if not sql.strip():
            QMessageBox.warning(self, "Gestione", "Nessun testo SQL da modificare.")
            return
        table_map: dict[str, str] = {}
        alias_map: dict[str, str] = {}
        column_map: dict[str, str] = {}
        for key, edit in self._gestione_inputs.items():
            new = edit.text().strip()
            if not new:
                continue
            if key.startswith("t:"):
                old = key[2:]
                if new != old:
                    table_map[old] = new
            elif key.startswith("a:"):
                old = key[2:]
                if new != old:
                    alias_map[old] = new
            elif key.startswith("c:"):
                bl = edit.property("gestioneColumnRaw")
                if bl is not None and str(bl).strip() == new:
                    continue
                column_map[key[2:]] = new
        if not table_map and not alias_map and not column_map:
            QMessageBox.information(
                self,
                "Gestione",
                "Compila almeno un campo «Sostituisci con» (diverso dall’originale dove serve).",
            )
            return
        new_sql = apply_gestione_replacements(
            sql,
            column_map=column_map or None,
            table_map=table_map or None,
            alias_map=alias_map or None,
        )
        self._sql_edit.setPlainText(new_sql)
        self._refresh_gestione_mapping()
        QMessageBox.information(
            self,
            "Gestione",
            "Sostituzioni applicate al testo SQL. Vai in Nuova e premi «Salva nella libreria» per aggiornare "
            "libreria, file .txt e voce in Archivio.",
        )

    def _on_new_blank(self) -> None:
        self._editing_id = None
        self._name_edit.clear()
        self._sql_edit.clear()
        self._query_list.clearSelection()
        self._sql_mode_menu.blockSignals(True)
        self._sql_mode_menu.setCurrentRow(0)
        self._sql_editor_stack.setCurrentIndex(0)
        self._sql_mode_menu.blockSignals(False)
        self._clear_gestione_form()
        self._gestione_empty.setText(
            "Seleziona una query nella lista oppure scrivi SQL in Nuova, poi apri Gestione."
        )
        self._gestione_empty.setVisible(True)
        self._gestione_scroll.setVisible(False)
        self._btn_gestione_apply.setEnabled(False)

    def _on_delete_query(self) -> None:
        s = self._settings()
        if s is None:
            QMessageBox.warning(self, "SQL", "Servizio impostazioni non disponibile.")
            return
        qid = self._editing_id
        if not qid:
            cur = self._query_list.currentItem()
            if cur is not None:
                qid = cur.data(Qt.ItemDataRole.UserRole)
        if not qid:
            QMessageBox.information(self, "SQL", "Seleziona una query da eliminare.")
            return
        qid = str(qid)
        name = self._queries_by_id().get(qid, {}).get("name", qid)
        reply = QMessageBox.question(
            self,
            "Elimina query",
            f"Rimuovere «{name}» dalla libreria?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        rows = [r for r in s.get_sql_archive_queries() if str(r.get("id")) != qid]
        s.set_sql_archive_queries(rows)
        self._editing_id = None
        self._name_edit.clear()
        self._sql_edit.clear()
        self._reload_query_list()
        self._query_list.clearSelection()
        self._clear_gestione_form()
        self._gestione_empty.setText(
            "Seleziona una query nella lista oppure scrivi SQL in Nuova, poi apri Gestione."
        )
        self._gestione_empty.setVisible(True)
        self._gestione_scroll.setVisible(False)
        self._btn_gestione_apply.setEnabled(False)
        if hasattr(self._mw, "refresh_views"):
            self._mw.refresh_views()
        QMessageBox.information(self, "SQL", "Query rimossa dalla libreria.")

    def _on_save_query(self) -> None:
        s = self._settings()
        if s is None:
            QMessageBox.warning(self, "SQL", "Servizio impostazioni non disponibile.")
            return
        sql_text = self._sql_edit.toPlainText().strip()
        if not sql_text:
            QMessageBox.warning(self, "SQL", "Inserisci il testo della query prima di salvare.")
            return
        name = self._name_edit.text().strip() or "Senza nome"
        rows = list(s.get_sql_archive_queries())
        eid = self._editing_id or str(uuid.uuid4())
        found = False
        for i, r in enumerate(rows):
            if str(r.get("id")) == eid:
                rows[i] = {"id": eid, "name": name, "sql_text": sql_text}
                found = True
                break
        if not found:
            rows.append({"id": eid, "name": name, "sql_text": sql_text})
        s.set_sql_archive_queries(rows)
        self._editing_id = eid
        self._reload_query_list()
        for i in range(self._query_list.count()):
            it = self._query_list.item(i)
            if it and it.data(Qt.ItemDataRole.UserRole) == eid:
                self._query_list.setCurrentRow(i)
                break

        export_note = ""
        try:
            self._ensure_query_tag()
            out_path = self._write_query_txt_and_archive(name, sql_text)
            export_note = (
                f"\n\nFile creato nella Directory Salvataggi (Setup Strumenti):\n{out_path}\n"
                f"Registrato in Archivio con tag «Query»."
            )
        except Exception as exc:
            QMessageBox.warning(
                self,
                "SQL",
                f"Query salvata nella libreria, ma esportazione file/archivio non riuscita:\n{exc}",
            )
            if hasattr(self._mw, "refresh_views"):
                self._mw.refresh_views()
            return

        if hasattr(self._mw, "refresh_views"):
            self._mw.refresh_views()
        QMessageBox.information(
            self,
            "SQL",
            "Query salvata nella libreria." + export_note,
        )

    def refresh(self) -> None:
        """Ricarica la lista (es. dopo sincronizzazione impostazioni)."""
        self._reload_query_list()
        if self._sql_editor_stack.currentIndex() == 1:
            self._refresh_gestione_mapping()
