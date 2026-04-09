"""Dialog per scegliere tabelle e formato di export dal database."""

from __future__ import annotations

from PyQt6.QtCore import Qt

from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from app.db_export import (
    EXPORT_FORMAT_CSV,
    EXPORT_FORMAT_JSON,
    EXPORT_FORMAT_XLSX,
    EXPORT_FORMAT_XML,
    table_label_it,
)

BTN_W, BTN_H = 260, 32


def _std_btn(btn: QPushButton) -> None:
    btn.setFixedSize(BTN_W, BTN_H)


class PackageExportDialog(QDialog):
    def __init__(self, table_names: list[str], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setWindowTitle("Esportazione dati")
        self.resize(520, 480)
        self._checks: dict[str, QCheckBox] = {}

        root = QVBoxLayout(self)
        root.setSpacing(12)

        intro = QLabel(
            "Scegli quali tabelle esportare e il formato del file. "
            "<b>Excel</b>: una tabella per foglio. "
            "<b>XML</b>: un solo file con sezioni per tabella. "
            "<b>JSON</b>: un file con tutte le tabelle in un oggetto. "
            "<b>CSV</b>: un file se esporti una sola tabella, altrimenti un file per tabella nella cartella scelta."
        )
        intro.setWordWrap(True)
        intro.setTextFormat(Qt.TextFormat.RichText)
        intro.setObjectName("helpLabel")
        root.addWidget(intro)

        fmt_row = QHBoxLayout()
        fmt_row.addWidget(QLabel("Formato file:"))
        self._fmt = QComboBox()
        self._fmt.addItem("Foglio Excel (.xlsx)", EXPORT_FORMAT_XLSX)
        self._fmt.addItem("Documento XML (.xml)", EXPORT_FORMAT_XML)
        self._fmt.addItem("File JSON (.json)", EXPORT_FORMAT_JSON)
        self._fmt.addItem("Testo CSV (.csv)", EXPORT_FORMAT_CSV)
        fmt_row.addWidget(self._fmt, 1)
        root.addLayout(fmt_row)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        inner = QWidget()
        inner_l = QVBoxLayout(inner)
        inner_l.setSpacing(6)

        sel_row = QHBoxLayout()
        b_all = QPushButton("Seleziona tutto")
        b_none = QPushButton("Deseleziona tutto")
        _std_btn(b_all)
        _std_btn(b_none)
        b_all.clicked.connect(self._select_all)
        b_none.clicked.connect(self._select_none)
        sel_row.addWidget(b_all)
        sel_row.addWidget(b_none)
        sel_row.addStretch(1)
        inner_l.addLayout(sel_row)

        for name in sorted(table_names, key=lambda n: (table_label_it(n).lower(), n)):
            cb = QCheckBox(table_label_it(name))
            cb.setToolTip(
                f"Nome tecnico nel database: {name}\n"
                "Nei file CSV il nome file usa questo nome (es. clients.csv)."
            )
            cb.setChecked(True)
            self._checks[name] = cb
            inner_l.addWidget(cb)
        inner_l.addStretch(1)
        scroll.setWidget(inner)
        root.addWidget(scroll, 1)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Conferma")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("Annulla")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

    def _select_all(self) -> None:
        for cb in self._checks.values():
            cb.setChecked(True)

    def _select_none(self) -> None:
        for cb in self._checks.values():
            cb.setChecked(False)

    def selected_tables(self) -> list[str]:
        return [name for name, cb in self._checks.items() if cb.isChecked()]

    def selected_format(self) -> str:
        return self._fmt.currentData()
