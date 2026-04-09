"""Widget note in header + finestra Note (pulsante in alto a destra)."""

from __future__ import annotations

from typing import Any

from PyQt6.QtCore import QDate, QDateTime, QTime, Qt, QTimer
from PyQt6.QtGui import QColor, QMouseEvent
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QFrame,
    QGraphicsDropShadowEffect,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QDateEdit,
    QTimeEdit,
    QHeaderView,
)


def _title_preview(text: str) -> tuple[str, str]:
    """Prima riga = titolo; resto = anteprima (finestra Note / altro)."""
    raw = (text or "").strip()
    if not raw:
        return ("(vuoto)", "")
    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]
    if not lines:
        return ("(vuoto)", "")
    t0 = lines[0]
    title = t0 if len(t0) <= 44 else t0[:43] + "…"
    if len(lines) >= 2:
        preview = " ".join(lines[1:]).strip()
    else:
        preview = ""
    if not preview and len(t0) > 44:
        preview = t0[43:].strip()
    if len(preview) > 160:
        preview = preview[:159] + "…"
    return title, preview


def _header_note_title(text: str) -> str:
    """Solo titolo (prima riga) per il widget compatto in header."""
    t = _title_preview(text)[0]
    return t if len(t) <= 36 else t[:35] + "…"


# Stile dialog nota (allineato a colori e radius dell’app HD Manager).
_STICKY_NOTE_DIALOG_QSS = """
#stickyNoteDialog {
    background-color: #f1f5f9;
}
#stickyNoteDialogCard {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 14px;
}
#stickyNoteDialogTitle {
    font-size: 18px;
    font-weight: 700;
    color: #0f172a;
}
#stickyNoteDialogHint {
    color: #64748b;
    font-size: 12px;
}
#stickyNoteDialogSection {
    font-size: 12px;
    font-weight: 700;
    color: #334155;
}
#stickyNoteDialogTextEdit {
    background: #f8fafc;
    border: 1px solid #cbd5e1;
    border-radius: 10px;
    padding: 12px;
    color: #0f172a;
    font-size: 13px;
    selection-background-color: #2563eb;
    selection-color: #ffffff;
}
#stickyNoteDialogTextEdit:focus {
    border: 2px solid #2563eb;
    background: #ffffff;
}
#stickyNoteDialogDateEdit, #stickyNoteDialogTimeEdit {
    background: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 10px;
    padding: 6px 10px;
    min-height: 30px;
    color: #0f172a;
}
#stickyNoteDialogDateEdit:focus, #stickyNoteDialogTimeEdit:focus {
    border: 2px solid #2563eb;
}
#stickyNoteDialogBtnCancel {
    background: #ffffff;
    color: #0f172a;
    border: 1px solid #94a3b8;
    border-radius: 10px;
    padding: 8px 18px;
    font-weight: 600;
    min-width: 100px;
}
#stickyNoteDialogBtnCancel:hover {
    background: #f8fafc;
    border-color: #64748b;
}
#stickyNoteDialogBtnCancel:pressed {
    background: #e2e8f0;
}
#stickyNoteDialogBtnSave {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #3b82f6, stop:1 #2563eb);
    color: #ffffff;
    border: 1px solid #1d4ed8;
    border-radius: 10px;
    padding: 8px 18px;
    font-weight: 700;
    min-width: 128px;
}
#stickyNoteDialogBtnSave:hover {
    background: #1d4ed8;
    border-color: #1e40af;
}
#stickyNoteDialogBtnSave:pressed {
    background: #1e40af;
}
"""


class StickyNoteDialog(QDialog):
    """Testo nota + data/ora di scadenza (poi la nota sparisce dal widget)."""

    def __init__(
        self,
        parent: QWidget | None,
        *,
        initial_text: str = "",
        initial_expires: QDateTime | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("stickyNoteDialog")
        self.setModal(True)
        self.setMinimumWidth(500)
        self.setMinimumHeight(380)
        self.setWindowTitle("Nota")

        try:
            win = parent.window() if parent is not None else None
            if win is not None:
                ico = win.windowIcon()
                if not ico.isNull():
                    self.setWindowIcon(ico)
        except Exception:
            pass

        is_edit = bool((initial_text or "").strip())
        heading = "Modifica nota" if is_edit else "Nuova nota"

        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 16)
        outer.setSpacing(0)

        card = QFrame()
        card.setObjectName("stickyNoteDialogCard")
        shadow = QGraphicsDropShadowEffect(card)
        shadow.setBlurRadius(20)
        shadow.setOffset(0, 4)
        shadow.setColor(QColor(15, 23, 42, 38))
        card.setGraphicsEffect(shadow)

        inner = QVBoxLayout(card)
        inner.setContentsMargins(24, 22, 24, 22)
        inner.setSpacing(18)

        title_lbl = QLabel(heading)
        title_lbl.setObjectName("stickyNoteDialogTitle")
        inner.addWidget(title_lbl)

        intro = QLabel(
            "La prima riga del testo è il titolo nel widget in alto. "
            "Passata la scadenza, la nota non compare più nel widget."
        )
        intro.setObjectName("stickyNoteDialogHint")
        intro.setWordWrap(True)
        inner.addWidget(intro)

        lbl_text = QLabel("Contenuto")
        lbl_text.setObjectName("stickyNoteDialogSection")
        inner.addWidget(lbl_text)

        self._text = QTextEdit()
        self._text.setObjectName("stickyNoteDialogTextEdit")
        self._text.setPlainText(initial_text)
        self._text.setMinimumHeight(150)
        self._text.setPlaceholderText("Scrivi la nota…")
        inner.addWidget(self._text)

        lbl_exp = QLabel("Scadenza")
        lbl_exp.setObjectName("stickyNoteDialogSection")
        inner.addWidget(lbl_exp)

        hint_exp = QLabel(
            "Dopo la data e l’ora indicate la nota non sarà più mostrata nel riquadro note in testata."
        )
        hint_exp.setObjectName("stickyNoteDialogHint")
        hint_exp.setWordWrap(True)
        inner.addWidget(hint_exp)

        ex = initial_expires
        if ex is None or not ex.isValid():
            ex = QDateTime.currentDateTime().addDays(1)

        row_dt = QHBoxLayout()
        row_dt.setSpacing(20)

        col_date = QVBoxLayout()
        col_date.setSpacing(6)
        lbl_date = QLabel("Data")
        lbl_date.setObjectName("stickyNoteDialogSection")
        self._date_edit = QDateEdit()
        self._date_edit.setObjectName("stickyNoteDialogDateEdit")
        self._date_edit.setDisplayFormat("dd/MM/yyyy")
        self._date_edit.setCalendarPopup(True)
        self._date_edit.setDate(ex.date())
        self._date_edit.setToolTip(
            "Giorno di scadenza della nota. Usa il pulsante ▾ per aprire il calendario."
        )
        hint_date = QLabel(
            "Scegli il giorno in cui la nota deve smettere di comparire (insieme all’ora accanto)."
        )
        hint_date.setObjectName("stickyNoteDialogHint")
        hint_date.setWordWrap(True)
        col_date.addWidget(lbl_date)
        col_date.addWidget(self._date_edit)
        col_date.addWidget(hint_date)

        col_time = QVBoxLayout()
        col_time.setSpacing(6)
        lbl_time = QLabel("Ora")
        lbl_time.setObjectName("stickyNoteDialogSection")
        self._time_edit = QTimeEdit()
        self._time_edit.setObjectName("stickyNoteDialogTimeEdit")
        self._time_edit.setDisplayFormat("HH:mm")
        self._time_edit.setTime(ex.time())
        self._time_edit.setToolTip(
            "Ora esatta (ore e minuti) in cui la nota non compare più nel widget in alto."
        )
        hint_time = QLabel(
            "Imposta l’ora della scadenza; a partire da quel momento la nota non è più mostrata."
        )
        hint_time.setObjectName("stickyNoteDialogHint")
        hint_time.setWordWrap(True)
        col_time.addWidget(lbl_time)
        col_time.addWidget(self._time_edit)
        col_time.addWidget(hint_time)

        row_dt.addLayout(col_date, 1)
        row_dt.addLayout(col_time, 1)
        inner.addLayout(row_dt)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        btn_row.addStretch(1)
        btn_cancel = QPushButton("Annulla")
        btn_cancel.setObjectName("stickyNoteDialogBtnCancel")
        btn_cancel.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton("Salva nota")
        btn_ok.setObjectName("stickyNoteDialogBtnSave")
        btn_ok.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_ok.setDefault(True)
        btn_ok.setAutoDefault(True)
        btn_ok.clicked.connect(self.accept)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        inner.addLayout(btn_row)

        outer.addWidget(card)
        self.setStyleSheet(_STICKY_NOTE_DIALOG_QSS)

    def values(self) -> tuple[str, QDateTime]:
        d = self._date_edit.date()
        t = self._time_edit.time()
        return (self._text.toPlainText().strip(), QDateTime(d, t))


class StickyNotesHeaderWidget(QWidget):
    """Scheda bianca con «+ Nuova nota» e lista verticale a schede (mockup)."""

    def __init__(self, main_window: QWidget, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._mw = main_window
        self._root = QHBoxLayout(self)
        self._root.setContentsMargins(4, 2, 8, 4)
        self._root.setSpacing(0)
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        self.setMinimumWidth(200)
        self.setMaximumWidth(320)
        self.setFixedHeight(142)
        self.setObjectName("stickyNotesHeader")
        self._timer = QTimer(self)
        self._timer.setInterval(30_000)
        self._timer.timeout.connect(self.refresh)
        self._timer.start()
        self._deadline = QTimer(self)
        self._deadline.setSingleShot(True)
        self._deadline.timeout.connect(self.refresh)
        self.refresh()

    def _settings(self):
        repo = getattr(self._mw, "repository", None)
        if repo is not None and hasattr(repo, "settings"):
            return repo.settings
        return None

    def _arm_next_expiry(self, notes: list[dict[str, Any]], now: QDateTime) -> None:
        self._deadline.stop()
        deltas: list[int] = []
        svc = self._settings()
        if svc is None:
            return
        for it in notes:
            exp = svc.sticky_note_expires_at(it)
            if exp is None or not exp.isValid() or now >= exp:
                continue
            s = now.secsTo(exp)
            if s > 0:
                deltas.append(s)
        if not deltas:
            return
        wait_ms = min(deltas) * 1000
        wait_ms = max(500, min(wait_ms, 3_600_000))
        self._deadline.start(int(wait_ms))

    @staticmethod
    def _expiry_date_label(exp: QDateTime | None, now: QDateTime) -> str:
        """Solo data/ora scadenza (nessun countdown) per anteprima compatta."""
        if exp is None or not exp.isValid():
            return ""
        if now >= exp:
            return "Scaduta"
        return exp.toString("dd/MM/yyyy HH:mm")

    def refresh(self) -> None:
        while self._root.count():
            item = self._root.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        card = QFrame()
        card.setObjectName("stickyNotesCard")
        card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Maximum)
        card.setMaximumHeight(132)
        shadow = QGraphicsDropShadowEffect(card)
        shadow.setBlurRadius(12)
        shadow.setOffset(0, 2)
        shadow.setColor(QColor(15, 23, 42, 45))
        card.setGraphicsEffect(shadow)

        outer = QVBoxLayout(card)
        outer.setContentsMargins(8, 5, 8, 6)
        outer.setSpacing(0)

        svc = self._settings()
        if svc is None:
            lbl = QLabel("—")
            lbl.setObjectName("stickyNoteCardMuted")
            outer.addWidget(lbl)
            self._root.addWidget(card, 1)
            return

        head = QHBoxLayout()
        head.setSpacing(8)
        head.addStretch(1)
        btn_new = QPushButton("+ Nuova nota")
        btn_new.setObjectName("stickyNotesNewButton")
        btn_new.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_new.clicked.connect(lambda: self._open_dialog(None))
        head.addWidget(btn_new, 0, Qt.AlignmentFlag.AlignRight)
        outer.addLayout(head)

        sep = QFrame()
        sep.setObjectName("stickyNotesCardSeparator")
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFixedHeight(1)
        outer.addWidget(sep)
        outer.addSpacing(4)

        notes = svc.get_sticky_notes()
        now = QDateTime.currentDateTime()
        self._arm_next_expiry(notes, now)

        scroll = QScrollArea()
        scroll.setObjectName("stickyNotesScroll")
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        # Fino a ~2 note visibili (36 + gap + 36 ≈ 76); oltre compare lo scroll
        scroll.setMinimumHeight(36)
        scroll.setMaximumHeight(82)
        scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        inner = QWidget()
        inner.setObjectName("stickyNotesScrollInner")
        col = QVBoxLayout(inner)
        col.setContentsMargins(0, 0, 2, 0)
        col.setSpacing(4)

        if not notes:
            empty = QLabel("Nessuna nota. Usa «+ Nuova nota» per aggiungerne una.")
            empty.setObjectName("stickyNoteCardMuted")
            empty.setWordWrap(True)
            col.addWidget(empty)
        else:
            for it in notes[:24]:
                col.addWidget(self._note_item_card(it, now, svc))
        scroll.setWidget(inner)
        outer.addWidget(scroll)
        self._root.addWidget(card, 1)

    def _note_item_card(self, it: dict[str, Any], now: QDateTime, svc) -> QFrame:
        exp = svc.sticky_note_expires_at(it) if svc else None
        text = str(it.get("text", "") or "")
        title = _header_note_title(text)

        fr = QFrame()
        fr.setObjectName("stickyNoteItemCard")
        fr.setCursor(Qt.CursorShape.PointingHandCursor)
        fr.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        fr.setMaximumHeight(36)
        lay = QVBoxLayout(fr)
        lay.setContentsMargins(5, 2, 5, 2)
        lay.setSpacing(0)

        title_lbl = QLabel(title)
        title_lbl.setObjectName("stickyNoteItemTitle")
        title_lbl.setWordWrap(False)
        title_lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        date_lbl = QLabel(self._expiry_date_label(exp, now))
        date_lbl.setObjectName("stickyNoteItemDate")
        date_lbl.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        date_lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        lay.addWidget(title_lbl)
        lay.addWidget(date_lbl)

        nid = str(it.get("id", ""))

        def on_click(e: QMouseEvent) -> None:
            if e.button() == Qt.MouseButton.LeftButton:
                self._open_dialog(nid)

        fr.mousePressEvent = on_click  # type: ignore[method-assign]
        return fr

    def _open_dialog(self, note_id: str | None = None) -> None:
        svc = self._settings()
        if svc is None:
            return
        initial = ""
        initial_exp: QDateTime | None = None
        if note_id:
            for it in svc.get_sticky_notes():
                if str(it.get("id")) == note_id:
                    initial = str(it.get("text", ""))
                    initial_exp = svc.sticky_note_expires_at(it)
                    break
        dlg = StickyNoteDialog(self, initial_text=initial, initial_expires=initial_exp)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        text, exp = dlg.values()
        if not text:
            QMessageBox.warning(self, "Nota", "Inserisci un testo.")
            return
        if exp <= QDateTime.currentDateTime():
            QMessageBox.warning(self, "Nota", "La data di scadenza deve essere nel futuro.")
            return
        svc.upsert_sticky_note(note_id, text, exp)
        self.refresh()
        if hasattr(self._mw, "refresh_views"):
            self._mw.refresh_views()

class NotesWorkspaceWidget(QWidget):
    """Finestra Note: gestione note (stesso storage dell’header)."""

    def __init__(self, main_window: QWidget, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._mw = main_window
        self._ctl = main_window.repository
        self.setObjectName("notesWindowPage")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(16, 16, 16, 16)
        lay.setSpacing(16)

        card = QFrame()
        card.setObjectName("clientDashboardCard")
        card.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        card_l = QVBoxLayout(card)
        card_l.setContentsMargins(0, 0, 0, 0)
        card_l.setSpacing(0)

        header = QFrame()
        header.setObjectName("clientInfoCardHeader")
        header.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        hl = QHBoxLayout(header)
        hl.setContentsMargins(14, 10, 14, 10)
        hl.setSpacing(10)
        ic = QLabel("📝")
        ic.setObjectName("clientInfoCardHeaderIcon")
        tl = QLabel("Note")
        tl.setObjectName("clientInfoCardHeaderTitle")
        hl.addWidget(ic)
        hl.addWidget(tl, 1)
        card_l.addWidget(header)

        body = QWidget()
        body_l = QVBoxLayout(body)
        body_l.setContentsMargins(14, 14, 14, 14)
        body_l.setSpacing(12)

        hint = QLabel(
            "Le note con scadenza sono visibili nel widget in alto (se abilitato in Impostazioni → Setup Strumenti). "
            "Passata la scadenza, la nota non compare più nel widget."
        )
        hint.setObjectName("accessProductHint")
        hint.setWordWrap(True)
        body_l.addWidget(hint)
        row = QHBoxLayout()
        btn_new = QPushButton("Nuova nota")
        btn_new.setObjectName("primaryActionButton")
        btn_edit = QPushButton("Modifica")
        btn_edit.setObjectName("secondaryActionButton")
        btn_del = QPushButton("Elimina")
        btn_del.setObjectName("secondaryActionButton")
        btn_new.clicked.connect(self._new)
        btn_edit.clicked.connect(self._edit)
        btn_del.clicked.connect(self._delete)
        row.addWidget(btn_new)
        row.addWidget(btn_edit)
        row.addWidget(btn_del)
        row.addStretch(1)
        body_l.addLayout(row)

        table_wrap = QFrame()
        table_wrap.setObjectName("clientDashboardTableWrap")
        table_wrap.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        tw = QVBoxLayout(table_wrap)
        tw.setContentsMargins(0, 0, 0, 0)
        self._table = QTableWidget(0, 3)
        self._table.setObjectName("notesWorkspaceTable")
        self._table.setHorizontalHeaderLabels(["Testo", "Scadenza", "id"])
        self._table.setColumnHidden(2, True)
        hh = self._table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        hh.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._table.setAlternatingRowColors(True)
        self._table.doubleClicked.connect(lambda _i: self._edit())
        tw.addWidget(self._table, 1)
        body_l.addWidget(table_wrap, 1)

        card_l.addWidget(body, 1)
        lay.addWidget(card, 1)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

    def _settings(self):
        if hasattr(self._ctl, "settings"):
            return self._ctl.settings
        return None

    def refresh(self) -> None:
        svc = self._settings()
        self._table.setRowCount(0)
        if svc is None:
            return
        notes = svc.get_sticky_notes()
        for it in notes:
            r = self._table.rowCount()
            self._table.insertRow(r)
            self._table.setItem(r, 0, QTableWidgetItem(str(it.get("text", ""))))
            exp = svc.sticky_note_expires_at(it)
            exp_s = exp.toString("dd/MM/yyyy HH:mm") if exp and exp.isValid() else "—"
            self._table.setItem(r, 1, QTableWidgetItem(exp_s))
            self._table.setItem(r, 2, QTableWidgetItem(str(it.get("id", ""))))

    def _selected_id(self) -> str | None:
        r = self._table.currentRow()
        if r < 0:
            return None
        it = self._table.item(r, 2)
        return (it.text() if it else "").strip() or None

    def _new(self) -> None:
        dlg = StickyNoteDialog(self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        text, exp = dlg.values()
        if not text:
            QMessageBox.warning(self, "Note", "Inserisci un testo.")
            return
        if exp <= QDateTime.currentDateTime():
            QMessageBox.warning(self, "Note", "La data di scadenza deve essere nel futuro.")
            return
        svc = self._settings()
        if svc is None:
            return
        svc.upsert_sticky_note(None, text, exp)
        self.refresh()
        if hasattr(self._mw, "refresh_views"):
            self._mw.refresh_views()

    def _edit(self) -> None:
        nid = self._selected_id()
        if not nid:
            QMessageBox.information(self, "Note", "Seleziona una nota.")
            return
        svc = self._settings()
        if svc is None:
            return
        initial = ""
        initial_exp: QDateTime | None = None
        for it in svc.get_sticky_notes():
            if str(it.get("id")) == nid:
                initial = str(it.get("text", ""))
                initial_exp = svc.sticky_note_expires_at(it)
                break
        dlg = StickyNoteDialog(self, initial_text=initial, initial_expires=initial_exp)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        text, exp = dlg.values()
        if not text:
            QMessageBox.warning(self, "Note", "Inserisci un testo.")
            return
        if exp <= QDateTime.currentDateTime():
            QMessageBox.warning(self, "Note", "La data di scadenza deve essere nel futuro.")
            return
        svc.upsert_sticky_note(nid, text, exp)
        self.refresh()
        if hasattr(self._mw, "refresh_views"):
            self._mw.refresh_views()

    def _delete(self) -> None:
        nid = self._selected_id()
        if not nid:
            QMessageBox.information(self, "Note", "Seleziona una nota.")
            return
        if QMessageBox.question(self, "Note", "Eliminare questa nota?") != QMessageBox.StandardButton.Yes:
            return
        svc = self._settings()
        if svc is None:
            return
        svc.delete_sticky_note(nid)
        self.refresh()
        if hasattr(self._mw, "refresh_views"):
            self._mw.refresh_views()
