from __future__ import annotations

from pathlib import Path

from PyQt6.QtCore import QDate, Qt
from PyQt6.QtGui import QFont, QGuiApplication, QIcon
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
    QMessageBox,
)


def _apply_dialog_style(dialog: QDialog) -> None:
    # Icona di programma (anche per titlebar/taskbar dei dialog).
    try:
        icon_path = Path(__file__).resolve().parent.parent / "assets" / "image.ico"
        if icon_path.exists():
            dialog.setWindowIcon(QIcon(str(icon_path)))
    except Exception:
        pass

    dialog.setStyleSheet(
        """
        QDialog {
            background: #f6f9fc;
            color: #0f172a;
            font-family: "Segoe UI";
            font-size: 12px;
        }
        QLabel {
            color: #1f2937;
        }
        QLineEdit,
        QComboBox,
        QDateEdit,
        QSpinBox,
        QListWidget {
            background: #ffffff;
            border: 1px solid #dbe5ee;
            border-radius: 8px;
            padding: 6px 8px;
        }
        QGroupBox {
            border: 1px solid #dbe5ee;
            border-radius: 12px;
            margin-top: 10px;
            background: #f6f9fc;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 12px;
            padding: 0 6px;
            color: #0f172a;
            font-weight: 600;
        }
        #hintCard {
            background: #eef6ff;
            border: 1px solid #dbe5ee;
            border-radius: 12px;
        }
        #primaryButton {
            background: #0f766e;
            color: #ffffff;
            border: 1px solid #0f766e;
            border-radius: 8px;
            padding: 7px 16px;
            font-weight: 600;
        }
        #primaryButton:hover {
            background: #0b5f58;
        }
        #secondaryButton {
            background: #e2e8f0;
            border: 1px solid #cbd5e1;
            border-radius: 8px;
            padding: 7px 16px;
        }
        """
    )


def _build_header(title: str, subtitle: str) -> QWidget:
    card = QWidget()
    card.setObjectName("hintCard")
    box = QVBoxLayout(card)
    box.setContentsMargins(12, 10, 12, 10)
    box.setSpacing(4)

    title_lbl = QLabel(title)
    title_lbl.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
    box.addWidget(title_lbl)

    sub_lbl = QLabel(subtitle)
    sub_lbl.setWordWrap(True)
    sub_lbl.setStyleSheet("color: #334155;")
    box.addWidget(sub_lbl)
    return card


class CredentialDialog(QDialog):
    def __init__(
        self,
        repository,
        client: dict,
        product: dict,
        flags: dict,
        credential: dict | None = None,
        credential_suggestion: dict | None = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.repository = repository
        self.client = client
        self.product = product
        self.flags = flags
        self.credential = credential or {}
        self.credential_suggestion = credential_suggestion or {}
        self.is_edit_mode = self.credential.get("id") is not None
        self.version_combos: dict[str, QComboBox] = {}

        settings = self.repository.settings if hasattr(self.repository, "settings") else self.repository
        all_environments = [row["label"] for row in settings.list_environments_lookup()]
        all_env_keys = {name.strip().lower() for name in all_environments}
        product_envs = self._csv_values(self.product.get("environments", ""))
        self.environments = [name for name in product_envs if name.lower() in all_env_keys]
        self.versions = [row["label"] for row in settings.list_releases_lookup()]

        self.setWindowTitle(
            "Modifica credenziale prodotto" if self.is_edit_mode else "Nuova credenziale prodotto"
        )
        self.resize(760, 760)
        self._build_ui()

    def _build_ui(self) -> None:
        self.setObjectName("credentialDialog")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(12)

        _apply_dialog_style(self)
        header = _build_header(
            "Modifica credenziale prodotto" if self.is_edit_mode else "Nuova credenziale prodotto",
            f"Cliente: {self.client.get('name', '')}   |   Prodotto: {self.product.get('name', '')}",
        )
        layout.addWidget(header)

        name_value = str(self.product.get("name", "")).strip()
        if self.is_edit_mode:
            name_value = str(self.credential.get("credential_name", name_value)).strip()
        self.name_edit = QLineEdit(name_value)
        self.name_edit.setReadOnly(True)
        self.name_edit.setPlaceholderText("Nome prodotto")
        name_box = QGroupBox("Identificazione")
        name_form = QFormLayout(name_box)
        name_form.setHorizontalSpacing(12)
        name_form.setVerticalSpacing(10)
        name_form.addRow("Nome prodotto", self.name_edit)
        layout.addWidget(name_box)

        env_box = QGroupBox("Ambienti e Versioni")
        env_layout = QGridLayout(env_box)
        env_layout.setHorizontalSpacing(10)
        env_layout.setVerticalSpacing(8)

        env_layout.addWidget(QLabel("Ambienti"), 0, 0, alignment=Qt.AlignmentFlag.AlignTop)
        self.env_list = QListWidget()
        self.env_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.env_list.setMinimumHeight(120)
        self.env_list.setMaximumHeight(160)
        if self.environments:
            for env_name in self.environments:
                self.env_list.addItem(QListWidgetItem(env_name))
        else:
            placeholder = QListWidgetItem("Nessun ambiente associato al prodotto")
            placeholder.setFlags(placeholder.flags() & ~Qt.ItemFlag.ItemIsSelectable)
            self.env_list.addItem(placeholder)
            self.env_list.setEnabled(False)
        self.env_list.itemSelectionChanged.connect(self._rebuild_version_rows)
        env_layout.addWidget(self.env_list, 0, 1)

        self.versions_container = QWidget()
        self.versions_layout = QVBoxLayout(self.versions_container)
        self.versions_layout.setContentsMargins(0, 0, 0, 0)
        self.versions_layout.setSpacing(6)
        env_layout.addWidget(QLabel("Versioni"), 1, 0, alignment=Qt.AlignmentFlag.AlignTop)
        env_layout.addWidget(self.versions_container, 1, 1)
        layout.addWidget(env_box)

        self.ip_edit: QLineEdit | None = None
        if self._flag("flag_ip"):
            self.ip_edit = QLineEdit()
            self.ip_edit.setPlaceholderText("es. 10.0.0.10")
        self.url_edit: QLineEdit | None = None
        if self._flag("flag_url"):
            self.url_edit = QLineEdit()
            self.url_edit.setPlaceholderText("es. https://app.example.com")
        self.host_edit: QLineEdit | None = None
        if self._flag("flag_host"):
            self.host_edit = QLineEdit()
            self.host_edit.setPlaceholderText("es. srv-app-prod")

        endpoint_box = QGroupBox("Endpoint")
        endpoint_form = QFormLayout(endpoint_box)
        endpoint_form.setHorizontalSpacing(12)
        endpoint_form.setVerticalSpacing(10)
        if self.ip_edit is not None:
            endpoint_form.addRow("IP *", self.ip_edit)
        if self.host_edit is not None:
            endpoint_form.addRow("Host *", self.host_edit)
        if self.url_edit is not None:
            endpoint_form.addRow("URL *", self.url_edit)

        self.rdp_edit: QLineEdit | None = None
        if self._flag("flag_preconfigured"):
            self.rdp_edit = QLineEdit()
            self.rdp_edit.setPlaceholderText("Seleziona file .rdp")
            browse_btn = QPushButton("Sfoglia…")
            browse_btn.clicked.connect(self._pick_rdp_path)
            browse_btn.setFixedWidth(90)
            row = QWidget()
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(8)
            row_layout.addWidget(self.rdp_edit, 1)
            row_layout.addWidget(browse_btn)
            endpoint_form.addRow("RDP preconfigurata *", row)

        if self._flag("flag_port"):
            self.port_edit = QLineEdit()
            self.port_edit.setPlaceholderText("1-65535")
            endpoint_form.addRow("Porta *", self.port_edit)
        else:
            self.port_edit = None

        if self.ip_edit is not None or self.host_edit is not None or self.url_edit is not None or self.rdp_edit is not None:
            layout.addWidget(endpoint_box)

        self.domain_edit = QLineEdit()
        self.login_edit = QLineEdit()
        self.username_edit = QLineEdit()
        self.username_edit.setReadOnly(True)
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.domain_edit.setPlaceholderText("es. DOMAIN")
        self.login_edit.setPlaceholderText("es. mario.rossi")
        self.password_edit.setPlaceholderText("Password")

        self.domain_edit.textChanged.connect(self._update_username)
        self.login_edit.textChanged.connect(self._update_username)

        access_box = QGroupBox("Credenziali di accesso")
        access_form = QFormLayout(access_box)
        access_form.setHorizontalSpacing(12)
        access_form.setVerticalSpacing(10)
        access_form.addRow("Dominio *", self.domain_edit)
        access_form.addRow("Nome utente *", self.login_edit)
        access_form.addRow("Username *", self.username_edit)

        password_row = QWidget()
        password_row_layout = QHBoxLayout(password_row)
        password_row_layout.setContentsMargins(0, 0, 0, 0)
        password_row_layout.setSpacing(8)
        password_row_layout.addWidget(self.password_edit, 1)
        self.show_password_chk = QCheckBox("Mostra")
        self.show_password_chk.toggled.connect(
            lambda checked: self.password_edit.setEchoMode(
                QLineEdit.EchoMode.Normal if checked else QLineEdit.EchoMode.Password
            )
        )
        password_row_layout.addWidget(self.show_password_chk)
        copy_btn = QPushButton("Copia")
        copy_btn.setFixedWidth(90)
        copy_btn.clicked.connect(self._copy_password)
        password_row_layout.addWidget(copy_btn)
        access_form.addRow("Password *", password_row)
        layout.addWidget(access_box)

        self.expiry_check = QCheckBox("Password con scadenza")
        self.expiry_check.toggled.connect(self._toggle_expiry_fields)

        self.insert_date_edit = QDateEdit(QDate.currentDate())
        self.insert_date_edit.setCalendarPopup(True)
        self.insert_date_edit.setMinimumWidth(160)
        self.insert_date_edit.dateChanged.connect(self._sync_expiry_fields_from_duration)
        self.duration_spin = QSpinBox()
        self.duration_spin.setRange(0, 3650)
        # Keep the value purely numeric; show unit outside the input.
        self.duration_spin.setSuffix("")
        # We'll show external +/- buttons for easier clicking.
        self.duration_spin.setButtonSymbols(QSpinBox.ButtonSymbols.NoButtons)
        self.duration_spin.setMinimumWidth(140)
        self.duration_spin.valueChanged.connect(self._sync_expiry_fields_from_duration)
        self.end_date_edit = QDateEdit(QDate.currentDate())
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setMinimumWidth(160)
        self.end_date_edit.dateChanged.connect(self._sync_expiry_fields_from_end_date)

        expiry_box = QGroupBox("Scadenza password")
        expiry_layout = QGridLayout(expiry_box)
        expiry_layout.setHorizontalSpacing(10)
        expiry_layout.setVerticalSpacing(8)
        expiry_layout.addWidget(self.expiry_check, 0, 0, 1, 2)
        expiry_layout.addWidget(self._stacked_field("Data inserimento", self.insert_date_edit), 1, 0)
        duration_row = QWidget()
        duration_row_layout = QHBoxLayout(duration_row)
        duration_row_layout.setContentsMargins(0, 0, 0, 0)
        duration_row_layout.setSpacing(8)
        minus_btn = QPushButton("−")
        minus_btn.setFixedWidth(34)
        minus_btn.clicked.connect(lambda: self.duration_spin.stepBy(-1))
        plus_btn = QPushButton("+")
        plus_btn.setFixedWidth(34)
        plus_btn.clicked.connect(lambda: self.duration_spin.stepBy(1))
        duration_row_layout.addWidget(minus_btn, 0)
        duration_row_layout.addWidget(self.duration_spin, 0)
        duration_row_layout.addWidget(plus_btn, 0)
        duration_row_layout.addWidget(QLabel("giorni"), 0)
        duration_row_layout.addStretch(1)
        expiry_layout.addWidget(self._stacked_field("Durata", duration_row), 1, 1)
        expiry_layout.addWidget(self._stacked_field("Data fine", self.end_date_edit), 2, 0, 1, 2)

        self.note_edit = QLineEdit()
        note_box = QGroupBox("Note")
        note_layout = QVBoxLayout(note_box)
        note_layout.setContentsMargins(10, 8, 10, 8)
        note_layout.addWidget(self.note_edit)

        bottom = QWidget()
        bottom_layout = QGridLayout(bottom)
        bottom_layout.setContentsMargins(0, 0, 0, 0)
        bottom_layout.setHorizontalSpacing(16)
        bottom_layout.setVerticalSpacing(12)
        bottom_layout.addWidget(expiry_box, 0, 0)
        bottom_layout.addWidget(note_box, 0, 1)
        layout.addWidget(bottom, 1)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        ok_btn = buttons.button(QDialogButtonBox.StandardButton.Ok)
        if ok_btn is not None:
            ok_btn.setText("Salva modifiche" if self.is_edit_mode else "Crea credenziale")
            ok_btn.setObjectName("primaryButton")
        cancel_btn = buttons.button(QDialogButtonBox.StandardButton.Cancel)
        if cancel_btn is not None:
            cancel_btn.setObjectName("secondaryButton")
        buttons.accepted.connect(self._submit)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self._toggle_expiry_fields(False)
        self._apply_initial_values()
        if not self.is_edit_mode:
            self._apply_credential_suggestion()

    def _copy_password(self) -> None:
        value = self.password_edit.text().strip()
        if not value:
            return
        QGuiApplication.clipboard().setText(value)
        QMessageBox.information(self, "Credenziale", "Password copiata negli appunti.")

    def _flag(self, key: str) -> bool:
        value = self.flags.get(key, 0)
        if isinstance(value, bool):
            return value
        try:
            return int(value) == 1
        except (TypeError, ValueError):
            return False

    def _pick_rdp_path(self) -> None:
        if self.rdp_edit is None:
            return
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleziona file RDP",
            "",
            "RDP files (*.rdp);;All files (*.*)",
        )
        if file_path:
            self.rdp_edit.setText(file_path)

    def _selected_environment_names(self) -> list[str]:
        names: list[str] = []
        for index in range(self.env_list.count()):
            item = self.env_list.item(index)
            if item.isSelected():
                names.append(item.text().strip())
        return names

    def _rebuild_version_rows(self) -> None:
        previous = {
            env_name: combo.currentText().strip()
            for env_name, combo in self.version_combos.items()
        }
        self.version_combos = {}
        while self.versions_layout.count():
            item = self.versions_layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

        for env_name in self._selected_environment_names():
            row = QWidget()
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(8)
            row_layout.addWidget(QLabel(f"Versione [{env_name}] (opzionale)"))
            combo = QComboBox()
            combo.addItem("")
            combo.addItems(self.versions)
            if env_name in previous and previous[env_name]:
                combo.setCurrentText(previous[env_name])
            row_layout.addWidget(combo, 1)
            self.versions_layout.addWidget(row)
            self.version_combos[env_name] = combo

    def _update_username(self) -> None:
        domain = self.domain_edit.text().strip()
        login = self.login_edit.text().strip()
        if domain and login:
            self.username_edit.setText(f"{domain}\\{login}")
            return
        self.username_edit.setText(login or "")

    def _toggle_expiry_fields(self, enabled: bool) -> None:
        self.insert_date_edit.setEnabled(enabled)
        self.duration_spin.setEnabled(enabled)
        self.end_date_edit.setEnabled(enabled)
        if enabled:
            # Ensure fields are consistent when enabling.
            self._sync_expiry_fields_from_duration()

    def _sync_expiry_fields_from_duration(self, *_args) -> None:
        if not self.expiry_check.isChecked():
            return
        base = self.insert_date_edit.date()
        days = int(self.duration_spin.value())
        target = base.addDays(days)
        if self.end_date_edit.date() != target:
            self.end_date_edit.blockSignals(True)
            self.end_date_edit.setDate(target)
            self.end_date_edit.blockSignals(False)

    def _sync_expiry_fields_from_end_date(self, *_args) -> None:
        if not self.expiry_check.isChecked():
            return
        base = self.insert_date_edit.date()
        end = self.end_date_edit.date()
        days = base.daysTo(end)
        if days < 0:
            # Prevent invalid backwards dates by snapping back.
            self.end_date_edit.blockSignals(True)
            self.end_date_edit.setDate(base)
            self.end_date_edit.blockSignals(False)
            days = 0
        if self.duration_spin.value() != days:
            self.duration_spin.blockSignals(True)
            self.duration_spin.setValue(days)
            self.duration_spin.blockSignals(False)

    @staticmethod
    def _is_true(value: object) -> bool:
        if isinstance(value, bool):
            return value
        if isinstance(value, int):
            return value == 1
        text = ("" if value is None else str(value)).strip().lower()
        return text in {"1", "true", "si", "yes", "y"}

    @staticmethod
    def _date_from_iso(value: str | None) -> QDate:
        parsed = QDate.fromString(str(value or "").strip(), "yyyy-MM-dd")
        return parsed if parsed.isValid() else QDate.currentDate()

    @staticmethod
    def _csv_values(value: str | None) -> list[str]:
        chunks = [chunk.strip() for chunk in str(value or "").split(",")]
        out: list[str] = []
        seen: set[str] = set()
        for chunk in chunks:
            key = chunk.lower()
            if chunk and key not in seen:
                seen.add(key)
                out.append(chunk)
        return out

    def _apply_initial_values(self) -> None:
        if not self.is_edit_mode:
            return

        if self.ip_edit is not None:
            self.ip_edit.setText(str(self.credential.get("ip") or "").strip())
        if self.url_edit is not None:
            self.url_edit.setText(str(self.credential.get("url") or "").strip())
        if self.host_edit is not None:
            self.host_edit.setText(str(self.credential.get("host") or "").strip())
        if self.rdp_edit is not None:
            self.rdp_edit.setText(str(self.credential.get("rdp_path") or "").strip())
        if self.port_edit is not None:
            self.port_edit.setText(str(self.credential.get("port") or "").strip())

        self.domain_edit.setText(str(self.credential.get("domain") or "").strip())
        self.login_edit.setText(str(self.credential.get("login_name") or "").strip())
        self._update_username()
        if not self.username_edit.text().strip():
            self.username_edit.setText(str(self.credential.get("username") or "").strip())
        self.password_edit.setText(str(self.credential.get("password") or "").strip())
        self.note_edit.setText(str(self.credential.get("note") or "").strip())

        env_to_release: dict[str, str] = {}
        for entry in self.credential.get("environment_versions", []):
            env_name = ""
            release_name = ""
            if isinstance(entry, tuple) and len(entry) >= 2:
                env_name = str(entry[0] or "").strip()
                release_name = str(entry[1] or "").strip()
            elif isinstance(entry, dict):
                env_name = str(entry.get("environment") or "").strip()
                release_name = str(entry.get("release") or "").strip()
            if env_name:
                env_to_release[env_name] = release_name

        for index in range(self.env_list.count()):
            item = self.env_list.item(index)
            item.setSelected(item.text().strip() in env_to_release)
        self._rebuild_version_rows()
        for env_name, release_name in env_to_release.items():
            combo = self.version_combos.get(env_name)
            if combo is None:
                continue
            if release_name:
                match_index = combo.findText(release_name)
                combo.setCurrentIndex(match_index if match_index >= 0 else 0)

        expiry_enabled = self._is_true(self.credential.get("password_expiry"))
        self.expiry_check.setChecked(expiry_enabled)
        self.insert_date_edit.setDate(
            self._date_from_iso(self.credential.get("password_inserted_at"))
        )
        duration_raw = self.credential.get("password_duration_days")
        try:
            duration_value = int(duration_raw) if duration_raw is not None else 0
        except (TypeError, ValueError):
            duration_value = 0
        self.duration_spin.setValue(max(0, duration_value))

        end_date_value = str(self.credential.get("password_end_date") or "").strip()
        if end_date_value:
            self.end_date_edit.setDate(self._date_from_iso(end_date_value))
            if expiry_enabled:
                self._sync_expiry_fields_from_end_date()
        elif expiry_enabled:
            self._sync_expiry_fields_from_duration()

    def _apply_credential_suggestion(self) -> None:
        """Precompila dominio/IP/Host dall'ultima credenziale dello stesso prodotto (solo suggerimento)."""
        sug = self.credential_suggestion
        if not sug:
            return
        dom = str(sug.get("domain") or "").strip()
        ip = str(sug.get("ip") or "").strip()
        host = str(sug.get("host") or "").strip()
        if dom and not self.domain_edit.text().strip():
            self.domain_edit.setText(dom)
            self._update_username()
        if self.ip_edit is not None and ip and not self.ip_edit.text().strip():
            self.ip_edit.setText(ip)
        if self.host_edit is not None and host and not self.host_edit.text().strip():
            self.host_edit.setText(host)

    def _submit(self) -> None:
        try:
            self._save_credential()
            self.accept()
        except ValueError as exc:
            title = "Modifica credenziale" if self.is_edit_mode else "Nuova credenziale"
            QMessageBox.warning(self, title, str(exc))

    def _require(self, value: str, label: str) -> str:
        cleaned = (value or "").strip()
        if not cleaned:
            raise ValueError(f"{label} obbligatorio.")
        return cleaned

    def _save_credential(self) -> None:
        if not self.environments:
            raise ValueError(
                "Il prodotto non ha ambienti associati. Configura prima gli ambienti in Impostazioni > Prodotti."
            )
        environments = self._selected_environment_names()
        if not environments:
            raise ValueError("Seleziona almeno un ambiente.")

        env_versions: list[tuple[str, str]] = []
        for env_name in environments:
            combo = self.version_combos.get(env_name)
            release_name = combo.currentText().strip() if combo is not None else ""
            env_versions.append((env_name, release_name))

        ip_value = self.ip_edit.text().strip() if self.ip_edit is not None else ""
        url_value = self.url_edit.text().strip() if self.url_edit is not None else ""
        host_value = self.host_edit.text().strip() if self.host_edit is not None else ""
        rdp_value = self.rdp_edit.text().strip() if self.rdp_edit is not None else ""
        port_value = self.port_edit.text().strip() if self.port_edit is not None else ""

        if self.ip_edit is not None:
            self._require(ip_value, "IP")
        if self.url_edit is not None:
            self._require(url_value, "URL")
        if self.host_edit is not None:
            self._require(host_value, "Host")
        if self.rdp_edit is not None:
            self._require(rdp_value, "Directory/file RDP preconfigurata")
        if self.port_edit is not None:
            self._require(port_value, "Porta")

        domain_value = self._require(self.domain_edit.text(), "Dominio")
        login_value = self._require(self.login_edit.text(), "Nome utente")
        username_value = self._require(self.username_edit.text(), "Username")
        password_value = self._require(self.password_edit.text(), "Password")

        expiry_enabled = self.expiry_check.isChecked()
        insert_date = (
            self.insert_date_edit.date().toString("yyyy-MM-dd") if expiry_enabled else None
        )
        duration_days = self.duration_spin.value() if expiry_enabled else None
        end_date = self.end_date_edit.date().toString("yyyy-MM-dd") if expiry_enabled else None
        if expiry_enabled:
            # Duration and end date are kept in sync: 0 days means same day (invalid).
            if duration_days is None or int(duration_days) <= 0:
                raise ValueError(
                    "Scadenza non valida: è stata impostata la stessa data di inserimento.\n"
                    "Imposta una Durata maggiore di 0 oppure una Data fine diversa."
                )

        payload = {
            "credential_name": self.name_edit.text().strip(),
            "environment_versions": env_versions,
            "domain": domain_value,
            "login_name": login_value,
            "username": username_value,
            "password": password_value,
            "ip": ip_value,
            "url": url_value,
            "host": host_value,
            "rdp_path": rdp_value,
            "port": port_value,
            "password_expiry": expiry_enabled,
            "password_inserted_at": insert_date,
            "password_duration_days": duration_days,
            "password_end_date": end_date,
            "note": self.note_edit.text().strip(),
        }

        creds = self.repository.credentials if hasattr(self.repository, "credentials") else self.repository
        if self.is_edit_mode:
            creds.update_product_credential(credential_id=int(self.credential["id"]), **payload)
            return

        creds.create_product_credential(
            client_id=int(self.client["id"]), product_id=int(self.product["id"]), **payload
        )

    @staticmethod
    def _stacked_field(label: str, widget: QWidget) -> QWidget:
        wrapper = QWidget()
        box = QVBoxLayout(wrapper)
        box.setContentsMargins(0, 0, 0, 0)
        box.setSpacing(4)
        box.addWidget(QLabel(label))
        box.addWidget(widget)
        return wrapper


class LinkDialog(QDialog):
    def __init__(
        self,
        tags: list[str] | None = None,
        link: dict | None = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.tags = tags or []
        self.link = link or {}
        self.is_edit = bool(self.link.get("id"))
        self.setWindowTitle("Modifica link" if self.is_edit else "Nuovo link")
        self.resize(440, 220)
        self._build_ui()

    def _build_ui(self) -> None:
        _apply_dialog_style(self)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(12)

        header = _build_header(
            "Modifica link" if self.is_edit else "Nuovo link",
            "Inserisci un nome descrittivo e l'URL. Il Tag è opzionale.",
        )
        layout.addWidget(header)

        self.name_edit = QLineEdit(str(self.link.get("name", "")).strip())
        self.url_edit = QLineEdit(str(self.link.get("url", "")).strip())
        self.name_edit.setPlaceholderText("es. Portale assistenza")
        self.url_edit.setPlaceholderText("es. https://example.com")
        self.tag_combo = QComboBox()
        self.tag_combo.addItem("")
        self.tag_combo.addItems(self.tags)
        current_tag = str(self.link.get("tag_name", "")).strip()
        if current_tag:
            self.tag_combo.setCurrentText(current_tag)

        form_box = QGroupBox("Dettagli link")
        form = QFormLayout(form_box)
        form.setLabelAlignment(Qt.AlignmentFlag.AlignLeft)
        form.setFormAlignment(Qt.AlignmentFlag.AlignTop)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(10)
        form.addRow("Nome *", self.name_edit)
        form.addRow("URL *", self.url_edit)
        form.addRow("Tag", self.tag_combo)
        layout.addWidget(form_box)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        ok_btn = buttons.button(QDialogButtonBox.StandardButton.Ok)
        if ok_btn is not None:
            ok_btn.setText("Salva" if self.is_edit else "Crea")
            ok_btn.setObjectName("primaryButton")
        cancel_btn = buttons.button(QDialogButtonBox.StandardButton.Cancel)
        if cancel_btn is not None:
            cancel_btn.setObjectName("secondaryButton")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def values(self) -> tuple[str, str, str]:
        name = self.name_edit.text().strip()
        url = self.url_edit.text().strip()
        if url and "://" not in url:
            url = f"https://{url}"
        return (name, url, self.tag_combo.currentText().strip())


class ContactDialog(QDialog):
    def __init__(
        self,
        roles: list[str] | None = None,
        contact: dict | None = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.roles = roles or []
        self.contact = contact or {}
        self.is_edit = bool(self.contact.get("id"))
        self.setWindowTitle("Modifica contatto" if self.is_edit else "Nuovo contatto")
        self.resize(520, 260)
        self._build_ui()

    def _build_ui(self) -> None:
        _apply_dialog_style(self)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(12)

        header = _build_header(
            "Modifica contatto" if self.is_edit else "Nuovo contatto",
            "Inserisci i dati del contatto. Email/telefono sono opzionali.",
        )
        layout.addWidget(header)

        self.name_edit = QLineEdit(str(self.contact.get("name", "")).strip())
        self.phone_edit = QLineEdit(str(self.contact.get("phone", "")).strip())
        self.mobile_edit = QLineEdit(str(self.contact.get("mobile", "")).strip())
        self.email_edit = QLineEdit(str(self.contact.get("email", "")).strip())
        self.name_edit.setPlaceholderText("es. Mario Rossi")
        self.phone_edit.setPlaceholderText("es. 06 1234567")
        self.mobile_edit.setPlaceholderText("es. +39 333 1234567")
        self.email_edit.setPlaceholderText("es. mario.rossi@azienda.it")
        self.role_combo = QComboBox()
        self.role_combo.addItem("")
        self.role_combo.addItems(self.roles)
        current_role = str(self.contact.get("role", "")).strip()
        if current_role:
            self.role_combo.setCurrentText(current_role)
        self.note_edit = QLineEdit(str(self.contact.get("note", "")).strip())
        self.note_edit.setPlaceholderText("Note (opzionale)")

        form_box = QGroupBox("Dettagli contatto")
        form = QFormLayout(form_box)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(10)
        form.addRow("Nome *", self.name_edit)
        form.addRow("Telefono", self.phone_edit)
        form.addRow("Cellulare", self.mobile_edit)
        form.addRow("Email", self.email_edit)
        form.addRow("Ruolo", self.role_combo)
        form.addRow("Note", self.note_edit)
        layout.addWidget(form_box)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        ok_btn = buttons.button(QDialogButtonBox.StandardButton.Ok)
        if ok_btn is not None:
            ok_btn.setText("Salva" if self.is_edit else "Crea")
            ok_btn.setObjectName("primaryButton")
        cancel_btn = buttons.button(QDialogButtonBox.StandardButton.Cancel)
        if cancel_btn is not None:
            cancel_btn.setObjectName("secondaryButton")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def values(self) -> dict:
        return {
            "name": self.name_edit.text().strip(),
            "phone": self.phone_edit.text().strip(),
            "mobile": self.mobile_edit.text().strip(),
            "email": self.email_edit.text().strip(),
            "role": self.role_combo.currentText().strip(),
            "note": self.note_edit.text().strip(),
        }
