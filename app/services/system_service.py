from __future__ import annotations

import logging
import os
import subprocess
import sys
import tempfile

from PyQt6.QtCore import QUrl
from PyQt6.QtGui import QDesktopServices

# Evita finestre console/PowerShell visibili su Windows
_SUBPROCESS_FLAGS = (
    subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
)


class SystemService:
    """OS integration: open files/urls, VPN, RDP."""

    def __init__(self) -> None:
        self._log = logging.getLogger(self.__class__.__name__)

    def open_local_file(self, path: str) -> bool:
        clean = str(path or "").strip()
        if not clean:
            return False
        self._log.info("Open local file: %s", clean)
        return QDesktopServices.openUrl(QUrl.fromLocalFile(clean))

    def open_url(self, url: str) -> bool:
        clean = str(url or "").strip()
        if not clean:
            return False
        if "://" not in clean:
            clean = f"https://{clean}"
        qurl = QUrl(clean)
        if not qurl.isValid():
            return False
        self._log.info("Open url: %s", clean)
        return QDesktopServices.openUrl(qurl)

    def start_vpn_windows(self, connection_name: str) -> None:
        name = str(connection_name or "").strip()
        if not name:
            raise ValueError("Nome connessione VPN non valido.")
        self._log.info("Start Windows VPN: %s", name)
        subprocess.Popen(["rasdial", name], creationflags=_SUBPROCESS_FLAGS)

    def start_file(self, path: str) -> None:
        clean = str(path or "").strip()
        if not clean:
            raise ValueError("Percorso non valido.")
        self._log.info("Start file/path: %s", clean)
        if hasattr(os, "startfile"):
            os.startfile(clean)  # type: ignore[attr-defined]
            return
        subprocess.Popen([clean], creationflags=_SUBPROCESS_FLAGS)

    def list_windows_vpn_connections(self) -> list[str]:
        """Returns VPN connection names from Windows, best-effort."""

        try:
            result = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-WindowStyle",
                    "Hidden",
                    "-Command",
                    "Get-VpnConnection | Select-Object -ExpandProperty Name",
                ],
                capture_output=True,
                text=True,
                check=False,
                creationflags=_SUBPROCESS_FLAGS,
            )
        except OSError as exc:
            self._log.warning("Unable to query Windows VPN connections: %s", exc)
            return []
        output = (result.stdout or "").strip()
        if not output:
            return []
        return [line.strip() for line in output.splitlines() if line.strip()]

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

    def connect_rdp(self, target: str, port: str, username: str, password: str) -> None:
        clean_target = str(target or "").strip()
        clean_user = str(username or "").strip()
        clean_password = str(password or "").strip()
        if not clean_target:
            raise ValueError("Target RDP non valorizzato.")
        if not clean_user or not clean_password:
            raise ValueError("Username e Password sono obbligatori per avviare la connessione RDP.")

        self._log.info("Connect RDP: target=%s user=%s", clean_target, clean_user)
        rdp_target = self._normalize_rdp_target(clean_target, port)
        target_variants = self._rdp_target_variants(clean_target, rdp_target)
        cmdkey_targets = [f"TERMSRV/{value}" for value in target_variants] + target_variants

        # Rimuove eventuali credenziali salvate vecchie che possono prevalere.
        for key in cmdkey_targets:
            subprocess.run(
                ["cmdkey", f"/delete:{key}"],
                check=False,
                capture_output=True,
                text=True,
                creationflags=_SUBPROCESS_FLAGS,
            )

        # Salvataggio ridondante per compatibilita tra vari comportamenti cred manager.
        for key in cmdkey_targets:
            subprocess.run(
                [
                    "cmdkey",
                    f"/generic:{key}",
                    f"/user:{clean_user}",
                    f"/pass:{clean_password}",
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
                    f"/user:{clean_user}",
                    f"/pass:{clean_password}",
                ],
                check=False,
                capture_output=True,
                text=True,
                creationflags=_SUBPROCESS_FLAGS,
            )

        rdp_file = self._build_rdp_launch_file(rdp_target, clean_user)
        subprocess.Popen(
            ["mstsc", rdp_file],
            creationflags=_SUBPROCESS_FLAGS,
        )

