from __future__ import annotations

import calendar
import json
import os
import uuid
from pathlib import Path
from typing import Any

from PyQt6.QtCore import QDate, QDateTime, Qt

from app.italian_holidays import fixed_recurring_italian_holidays
from app.services.repository_service import RepositoryService


class SettingsService:
    """CRUD + lookups used by SettingsWindow."""

    def __init__(self, repo: RepositoryService) -> None:
        self.repo = repo

    # lookups
    def list_competences_lookup(self) -> list[dict]:
        return self.repo.list_competences_lookup()

    def list_product_types_lookup(self) -> list[dict]:
        return self.repo.list_product_types_lookup()

    def list_releases_lookup(self) -> list[dict]:
        return self.repo.list_releases_lookup()

    def list_environments_lookup(self) -> list[dict]:
        return self.repo.list_environments_lookup()

    def list_roles_lookup(self) -> list[dict]:
        return self.repo.list_roles_lookup()

    def list_vpns_lookup(self) -> list[dict]:
        return self.repo.list_vpns_lookup()

    def list_resources_lookup(self) -> list[dict]:
        return self.repo.list_resources_lookup()

    def list_clients_lookup(self) -> list[dict]:
        return self.repo.list_clients_lookup()

    def list_tags_lookup(self) -> list[dict]:
        return self.repo.list_tags_lookup()

    # competences
    def list_competences(self) -> list[dict]:
        return self.repo.list_competences()

    def list_competences_with_resources(self) -> list[dict]:
        return self.repo.list_competences_with_resources()

    def assign_resources_to_competence(
        self,
        competence_name: str,
        selected_resource_labels: list[str],
    ) -> None:
        self.repo.assign_resources_to_competence(competence_name, selected_resource_labels)

    def upsert_competence(self, row_id: int | None, name: str) -> int:
        return self.repo.upsert_competence(row_id, name)

    def delete_competence(self, row_id: int) -> None:
        self.repo.delete_competence(row_id)

    # product types
    def list_product_types(self) -> list[dict]:
        return self.repo.list_product_types()

    def upsert_product_type(self, row_id: int | None, **data) -> int:
        return self.repo.upsert_product_type(
            row_id,
            data["name"],
            data["flag_ip"],
            data["flag_host"],
            data["flag_preconfigured"],
            data["flag_url"],
            data["flag_port"],
        )

    def delete_product_type(self, row_id: int) -> None:
        self.repo.delete_product_type(row_id)

    # tags
    def list_tags(self) -> list[dict]:
        return self.repo.list_tags()

    def upsert_tag(self, row_id: int | None, **data) -> int:
        return self.repo.upsert_tag(
            row_id,
            data["name"],
            data["color"],
            data.get("client_name") or "",
        )

    def delete_tag(self, row_id: int) -> None:
        self.repo.delete_tag(row_id)

    # products
    def list_products(self) -> list[dict]:
        return self.repo.list_products()

    def upsert_product(self, row_id: int | None, **data) -> int:
        return self.repo.upsert_product(
            row_id,
            data["name"],
            data["product_type"],
            data.get("clients"),
            data.get("environments"),
        )

    def delete_product(self, row_id: int) -> None:
        self.repo.delete_product(row_id)

    # environments
    def list_environments(self) -> list[dict]:
        return self.repo.list_environments()

    def upsert_environment(self, row_id: int | None, **data) -> int:
        return self.repo.upsert_environment(row_id, data["name"], "")

    def delete_environment(self, row_id: int) -> None:
        self.repo.delete_environment(row_id)

    # releases
    def list_releases(self) -> list[dict]:
        return self.repo.list_releases()

    def upsert_release(self, row_id: int | None, **data) -> int:
        return self.repo.upsert_release(row_id, data["name"], "")

    def delete_release(self, row_id: int) -> None:
        self.repo.delete_release(row_id)

    # clients
    def list_clients(self) -> list[dict]:
        return self.repo.list_clients()

    def upsert_client(self, row_id: int | None, **data) -> int:
        return self.repo.upsert_client(
            row_id,
            data["name"],
            data["location"],
            "",
            data.get("resources"),
            data.get("link") or "",
        )

    def delete_client(self, row_id: int) -> None:
        self.repo.delete_client(row_id)

    # resources
    def list_resources(self) -> list[dict]:
        return self.repo.list_resources()

    def upsert_resource(self, row_id: int | None, **data) -> int:
        return self.repo.upsert_resource(
            row_id,
            data["name"],
            data["surname"],
            data["role_name"],
            data.get("competence_name") or "",
            data.get("phone") or "",
            data.get("email") or "",
            data.get("linkedin") or "",
            data.get("photo_link") or "",
            data.get("note") or "",
        )

    def delete_resource(self, row_id: int) -> None:
        self.repo.delete_resource(row_id)

    # roles
    def list_roles(self) -> list[dict]:
        return self.repo.list_roles()

    def upsert_role(self, row_id: int | None, **data) -> int:
        competence_value = data.get("competence")
        if (not competence_value or not str(competence_value).strip()) and row_id:
            # Se l'UI non espone il campo competenza, preserviamo quella già salvata nel DB.
            try:
                existing = self.repo._query(
                    "SELECT competence FROM roles WHERE id = ?;",
                    (int(row_id),),
                )
                if existing:
                    competence_value = existing[0].get("competence") or ""
            except Exception:
                competence_value = ""

        if not competence_value or not str(competence_value).strip():
            # Fallback per inserimento: mettiamo la prima competenza esistente (se presente).
            try:
                comps = self.repo.list_competences()
                competence_value = comps[0]["name"] if comps else ""
            except Exception:
                competence_value = ""

        return self.repo.upsert_role(
            row_id,
            data["name"],
            data["multi_clients"],
            str(competence_value or ""),
            data.get("display_order"),
        )

    def delete_role(self, row_id: int) -> None:
        self.repo.delete_role(row_id)

    # vpn
    def list_vpns(self) -> list[dict]:
        return self.repo.list_vpns()

    def upsert_vpn(self, row_id: int | None, **data) -> int:
        vpn_type = data["vpn_type"]
        connection_name = (
            data["vpn_windows_name"] if vpn_type == "VPN Windows" else data["connection_name"]
        )
        clients_raw = data.get("clients")
        if clients_raw is None:
            clients_val: str | list[str] = ""
        elif isinstance(clients_raw, list):
            clients_val = clients_raw
        else:
            clients_val = str(clients_raw).strip() if clients_raw else ""
        return self.repo.upsert_vpn(
            row_id,
            connection_name,
            data["server_address"],
            vpn_type,
            data.get("access_info_type") or "",
            data["username"],
            data["password"],
            data.get("vpn_path") or "",
            clients_val,
        )

    def delete_vpn(self, row_id: int) -> None:
        self.repo.delete_vpn(row_id)

    def list_windows_vpn_connections(self) -> list[str]:
        return self.repo.list_windows_vpn_connections()

    # -----------------------
    # UI / App settings
    # -----------------------
    def get_app_setting(self, key: str, default: str = "") -> str:
        conn = self.repo.connection_factory()
        try:
            row = conn.execute(
                "SELECT value FROM app_settings WHERE key = ?;",
                (key,),
            ).fetchone()
            if not row:
                return default
            return str(row["value"] or "")
        finally:
            conn.close()

    def set_app_setting(self, key: str, value: str) -> None:
        conn = self.repo.connection_factory()
        try:
            with conn:
                conn.execute(
                    "INSERT OR REPLACE INTO app_settings(key, value) VALUES (?, ?);",
                    (key, value),
                )
        finally:
            conn.close()

    # Banner image in the main window header (top-right empty space).
    def get_header_banner_png_path(self) -> str:
        return self.get_app_setting("header_banner_png_path", default="")

    def set_header_banner_png_path(self, png_path: str) -> None:
        # Permettiamo stringa vuota per "rimuovi".
        self.set_app_setting("header_banner_png_path", png_path or "")

    def get_header_banner_height_px(self) -> int:
        raw = self.get_app_setting("header_banner_height_px", default="90")
        try:
            value = int(raw)
        except (TypeError, ValueError):
            return 90
        return min(240, max(40, value))

    def set_header_banner_height_px(self, height_px: int) -> None:
        value = min(240, max(40, int(height_px)))
        self.set_app_setting("header_banner_height_px", str(value))

    def get_header_banner_width_px(self) -> int:
        raw = self.get_app_setting("header_banner_width_px", default="360")
        try:
            value = int(raw)
        except (TypeError, ValueError):
            return 360
        return min(1200, max(120, value))

    def set_header_banner_width_px(self, width_px: int) -> None:
        value = min(1200, max(120, int(width_px)))
        self.set_app_setting("header_banner_width_px", str(value))

    def get_header_banner_scale_percent(self) -> int:
        raw = self.get_app_setting("header_banner_scale_percent", default="100")
        try:
            value = int(raw)
        except (TypeError, ValueError):
            return 100
        return min(200, max(10, value))

    def set_header_banner_scale_percent(self, scale_percent: int) -> None:
        value = min(200, max(10, int(scale_percent)))
        self.set_app_setting("header_banner_scale_percent", str(value))

    # -----------------------
    # Strumenti (calcolatrice) + calendario lavorativo
    # -----------------------
    def get_calculator_save_directory(self) -> str:
        return (self.get_app_setting("calculator_save_directory", default="") or "").strip()

    def set_calculator_save_directory(self, path: str) -> None:
        self.set_app_setting("calculator_save_directory", (path or "").strip())

    def resolve_calculator_export_path(self) -> Path | None:
        """Cartella configurata dall'utente, se esiste ed è una directory."""
        raw = self.get_calculator_save_directory()
        if not raw:
            return None
        p = Path(raw).expanduser().resolve()
        if os.path.isdir(str(p)):
            return p
        return None

    def get_calculator_export_base_path(self) -> Path:
        """Directory predefinita per «Salva vista»: impostazione o cartella applicazione."""
        resolved = self.resolve_calculator_export_path()
        if resolved is not None:
            resolved.mkdir(parents=True, exist_ok=True)
            return resolved
        local = os.getenv("LOCALAPPDATA") or os.path.expanduser("~")
        p = Path(local) / "HDManagerDesktop" / "calculator_exports"
        p.mkdir(parents=True, exist_ok=True)
        return p

    def get_agenda_header_widget_enabled(self) -> bool:
        raw = (self.get_app_setting("agenda_header_widget_enabled", default="1") or "").strip().lower()
        return raw not in ("0", "false", "no", "off")

    def set_agenda_header_widget_enabled(self, enabled: bool) -> None:
        self.set_app_setting("agenda_header_widget_enabled", "1" if enabled else "0")

    def get_notes_widget_enabled(self) -> bool:
        raw = (self.get_app_setting("notes_widget_enabled", default="0") or "").strip().lower()
        return raw in ("1", "true", "yes", "on")

    def set_notes_widget_enabled(self, enabled: bool) -> None:
        self.set_app_setting("notes_widget_enabled", "1" if enabled else "0")

    _STICKY_NOTES_KEY = "sticky_notes_v1"

    def get_sticky_notes(self) -> list[dict[str, Any]]:
        """Note attive (non scadute); le scadute vengono rimosse dal salvataggio."""
        raw = (self.get_app_setting(self._STICKY_NOTES_KEY, default="") or "").strip()
        if not raw:
            return []
        try:
            data = json.loads(raw)
            if not isinstance(data, list):
                return []
            items = [x for x in data if isinstance(x, dict)]
        except json.JSONDecodeError:
            return []
        now = QDateTime.currentDateTime()
        kept: list[dict[str, Any]] = []
        for it in items:
            exp = self._sticky_note_expires_qdt(it)
            if exp is None or not exp.isValid():
                continue
            if now >= exp:
                continue
            tid = str(it.get("id") or "").strip()
            text = str(it.get("text") or "").strip()
            if not tid or not text:
                continue
            kept.append(
                {
                    "id": tid,
                    "text": text,
                    "expires_at": exp.toString(Qt.DateFormat.ISODateWithMs),
                }
            )
        kept.sort(key=lambda x: self._sticky_note_expires_qdt(x) or QDateTime())
        if len(kept) != len(items):
            self.set_sticky_notes(kept)
        return kept

    @staticmethod
    def _sticky_note_expires_qdt(note: dict[str, Any]) -> QDateTime | None:
        raw = str(note.get("expires_at") or "").strip()
        if not raw:
            return None
        dt = QDateTime.fromString(raw, Qt.DateFormat.ISODateWithMs)
        if not dt.isValid():
            dt = QDateTime.fromString(raw, Qt.DateFormat.ISODate)
        return dt if dt.isValid() else None

    def sticky_note_expires_at(self, note: dict[str, Any]) -> QDateTime | None:
        return self._sticky_note_expires_qdt(note)

    def set_sticky_notes(self, items: list[dict[str, Any]]) -> None:
        """Salva solo note con id, testo e scadenza valide; ordina per scadenza."""
        now = QDateTime.currentDateTime()
        clean: list[dict[str, Any]] = []
        for it in items:
            if not isinstance(it, dict):
                continue
            tid = str(it.get("id") or "").strip()
            text = str(it.get("text") or "").strip()
            exp = self._sticky_note_expires_qdt(it)
            if not tid or not text or exp is None or not exp.isValid():
                continue
            if now >= exp:
                continue
            clean.append(
                {
                    "id": tid,
                    "text": text,
                    "expires_at": exp.toString(Qt.DateFormat.ISODateWithMs),
                }
            )
        clean.sort(key=lambda x: self._sticky_note_expires_qdt(x) or QDateTime())
        self.set_app_setting(self._STICKY_NOTES_KEY, json.dumps(clean, ensure_ascii=False))

    def upsert_sticky_note(self, note_id: str | None, text: str, expires: QDateTime) -> None:
        tid = (note_id or "").strip() or str(uuid.uuid4())
        items = [x for x in self.get_sticky_notes() if str(x.get("id")) != tid]
        items.append(
            {
                "id": tid,
                "text": (text or "").strip(),
                "expires_at": expires.toString(Qt.DateFormat.ISODateWithMs),
            }
        )
        self.set_sticky_notes(items)

    def delete_sticky_note(self, note_id: str) -> None:
        tid = (note_id or "").strip()
        if not tid:
            return
        items = [x for x in self.get_sticky_notes() if str(x.get("id")) != tid]
        self.set_sticky_notes(items)

    @staticmethod
    def _normalize_hhmm(raw: str, default: str) -> str:
        s = (raw or "").strip().replace(".", ":")
        if not s:
            return default
        parts = s.split(":")
        if len(parts) >= 2:
            try:
                h = int(parts[0])
                m = int(parts[1])
                if 0 <= h <= 23 and 0 <= m <= 59:
                    return f"{h:02d}:{m:02d}"
            except (TypeError, ValueError):
                pass
        return default

    def _default_day_schedule(self, weekday: int) -> dict[str, Any]:
        """weekday 0=Lun … 6=Dom."""
        lavoro = weekday < 5
        return {
            "lavorativo": lavoro,
            "inizio": "09:00",
            "fine": "18:00",
            "pausa_abilitata": lavoro,
            "pausa_inizio": "13:00",
            "pausa_fine": "14:00",
        }

    def _default_schedule(self) -> dict[int, dict[str, Any]]:
        return {d: dict(self._default_day_schedule(d)) for d in range(7)}

    def get_work_schedule(self) -> dict[int, dict[str, Any]]:
        """Orario per giorno: lavorativo, inizio/fine turno, pausa pranzo."""
        raw = (self.get_app_setting("work_schedule_json", default="") or "").strip()
        if raw:
            try:
                data = json.loads(raw)
                if isinstance(data, dict):
                    return self._normalize_schedule(data)
            except json.JSONDecodeError:
                pass
        return self._migrate_legacy_schedule()

    def _normalize_schedule(self, data: dict[str, Any]) -> dict[int, dict[str, Any]]:
        out: dict[int, dict[str, Any]] = {}
        for d in range(7):
            base = self._default_day_schedule(d)
            row = data.get(str(d))
            if not isinstance(row, dict):
                row = {}
            lav = bool(row.get("lavorativo", base["lavorativo"]))
            out[d] = {
                "lavorativo": lav,
                "inizio": self._normalize_hhmm(str(row.get("inizio", "")), base["inizio"]),
                "fine": self._normalize_hhmm(str(row.get("fine", "")), base["fine"]),
                "pausa_abilitata": bool(
                    row.get("pausa_abilitata", row.get("pausa", base["pausa_abilitata"]))
                ),
                "pausa_inizio": self._normalize_hhmm(
                    str(row.get("pausa_inizio", "")), base["pausa_inizio"]
                ),
                "pausa_fine": self._normalize_hhmm(str(row.get("pausa_fine", "")), base["pausa_fine"]),
            }
        return out

    def _migrate_legacy_schedule(self) -> dict[int, dict[str, Any]]:
        start = self._normalize_hhmm(self.get_app_setting("workday_start", "09:00"), "09:00")
        end = self._normalize_hhmm(self.get_app_setting("workday_end", "18:00"), "18:00")
        raw = (self.get_app_setting("work_days", default="0,1,2,3,4") or "").strip()
        days_set: set[int] = set()
        for part in raw.split(","):
            part = part.strip()
            if not part:
                continue
            try:
                d = int(part)
            except ValueError:
                continue
            if 0 <= d <= 6:
                days_set.add(d)
        if not days_set:
            days_set = {0, 1, 2, 3, 4}
        out: dict[int, dict[str, Any]] = {}
        for d in range(7):
            lavoro = d in days_set
            out[d] = {
                "lavorativo": lavoro,
                "inizio": start,
                "fine": end,
                "pausa_abilitata": lavoro,
                "pausa_inizio": "13:00",
                "pausa_fine": "14:00",
            }
        return out

    def set_work_schedule(self, schedule: dict[int, dict[str, Any]]) -> None:
        payload: dict[str, dict[str, Any]] = {}
        for d in range(7):
            row = schedule.get(d) or {}
            base = self._default_day_schedule(d)
            lav = bool(row.get("lavorativo", base["lavorativo"]))
            payload[str(d)] = {
                "lavorativo": lav,
                "inizio": self._normalize_hhmm(str(row.get("inizio", "")), base["inizio"]),
                "fine": self._normalize_hhmm(str(row.get("fine", "")), base["fine"]),
                "pausa_abilitata": bool(row.get("pausa_abilitata", base["pausa_abilitata"])),
                "pausa_inizio": self._normalize_hhmm(
                    str(row.get("pausa_inizio", "")), base["pausa_inizio"]
                ),
                "pausa_fine": self._normalize_hhmm(str(row.get("pausa_fine", "")), base["pausa_fine"]),
            }
        self.set_app_setting("work_schedule_json", json.dumps(payload, ensure_ascii=False))
        self._sync_legacy_from_schedule(payload)

    def _sync_legacy_from_schedule(self, payload: dict[str, dict[str, Any]]) -> None:
        days = [d for d in range(7) if payload[str(d)]["lavorativo"]]
        if not days:
            days = [0, 1, 2, 3, 4]
        self.set_work_days_raw(days)
        ref = payload[str(days[0])]
        self.set_workday_start_raw(ref["inizio"])
        self.set_workday_end_raw(ref["fine"])

    def set_work_days_raw(self, days: list[int]) -> None:
        clean: list[int] = []
        for d in days:
            try:
                v = int(d)
            except (TypeError, ValueError):
                continue
            if 0 <= v <= 6 and v not in clean:
                clean.append(v)
        clean.sort()
        if not clean:
            clean = [0, 1, 2, 3, 4]
        self.set_app_setting("work_days", ",".join(str(d) for d in clean))

    def set_workday_start_raw(self, value: str) -> None:
        self.set_app_setting("workday_start", self._normalize_hhmm(value, "09:00"))

    def set_workday_end_raw(self, value: str) -> None:
        self.set_app_setting("workday_end", self._normalize_hhmm(value, "18:00"))

    def get_workday_start(self) -> str:
        if (self.get_app_setting("work_schedule_json", default="") or "").strip():
            sch = self.get_work_schedule()
            for d in range(7):
                if sch[d]["lavorativo"]:
                    return sch[d]["inizio"]
            return "09:00"
        raw = self.get_app_setting("workday_start", default="09:00")
        return self._normalize_hhmm(raw, "09:00")

    def set_workday_start(self, value: str) -> None:
        self.set_app_setting("workday_start", self._normalize_hhmm(value, "09:00"))

    def get_workday_end(self) -> str:
        if (self.get_app_setting("work_schedule_json", default="") or "").strip():
            sch = self.get_work_schedule()
            for d in range(7):
                if sch[d]["lavorativo"]:
                    return sch[d]["fine"]
            return "18:00"
        raw = self.get_app_setting("workday_end", default="18:00")
        return self._normalize_hhmm(raw, "18:00")

    def set_workday_end(self, value: str) -> None:
        self.set_app_setting("workday_end", self._normalize_hhmm(value, "18:00"))

    def get_work_days(self) -> list[int]:
        """Giorni lavorativi 0=Lunedì … 6=Domenica (come datetime.weekday())."""
        if (self.get_app_setting("work_schedule_json", default="") or "").strip():
            sch = self.get_work_schedule()
            return [d for d in range(7) if sch[d]["lavorativo"]]
        raw = (self.get_app_setting("work_days", default="0,1,2,3,4") or "").strip()
        if not raw:
            return [0, 1, 2, 3, 4]
        out: list[int] = []
        for part in raw.split(","):
            part = part.strip()
            if not part:
                continue
            try:
                d = int(part)
            except ValueError:
                continue
            if 0 <= d <= 6 and d not in out:
                out.append(d)
        return sorted(out) if out else [0, 1, 2, 3, 4]

    def set_work_days(self, days: list[int]) -> None:
        clean: list[int] = []
        for d in days:
            try:
                v = int(d)
            except (TypeError, ValueError):
                continue
            if 0 <= v <= 6 and v not in clean:
                clean.append(v)
        clean.sort()
        if not clean:
            clean = [0, 1, 2, 3, 4]
        self.set_app_setting("work_days", ",".join(str(d) for d in clean))

    # Agenda (finestra dedicata)
    _AGENDA_ITEMS_KEY = "agenda_items_v1"

    def get_agenda_items(self) -> list[dict[str, Any]]:
        raw = (self.get_app_setting(self._AGENDA_ITEMS_KEY, default="") or "").strip()
        if not raw:
            return []
        try:
            data = json.loads(raw)
            if isinstance(data, list):
                return [x for x in data if isinstance(x, dict)]
        except json.JSONDecodeError:
            pass
        return []

    def set_agenda_items(self, items: list[dict[str, Any]]) -> None:
        self.set_app_setting(self._AGENDA_ITEMS_KEY, json.dumps(items, ensure_ascii=False))

    _PUBLIC_HOLIDAYS_KEY = "public_holidays_v1"

    @staticmethod
    def _valid_month_day(month: int, day: int) -> bool:
        if month < 1 or month > 12 or day < 1 or day > 31:
            return False
        _, maxd = calendar.monthrange(2024, month)
        return day <= maxd

    def _normalize_public_holiday_rows(self, rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """Righe {month, day, label}; migrazione da legacy {date, label} (solo giorno/mese)."""
        seen: set[tuple[int, int]] = set()
        out: list[dict[str, Any]] = []
        for r in rows:
            if not isinstance(r, dict):
                continue
            m: int | None = None
            d: int | None = None
            if "month" in r and "day" in r:
                try:
                    m = int(r["month"])
                    d = int(r["day"])
                except (TypeError, ValueError):
                    m = d = None
            if m is None and r.get("date"):
                qd = QDate.fromString(str(r["date"]).strip()[:10], Qt.DateFormat.ISODate)
                if qd.isValid():
                    m, d = qd.month(), qd.day()
            if m is None or d is None or not self._valid_month_day(m, d):
                continue
            if (m, d) in seen:
                continue
            seen.add((m, d))
            lab = str(r.get("label", "")).strip() or "Festività"
            out.append({"month": m, "day": d, "label": lab})
        out.sort(key=lambda x: (x["month"], x["day"]))
        return out

    def _default_public_holiday_rows(self) -> list[dict[str, Any]]:
        return [
            {"month": mo, "day": da, "label": lab}
            for mo, da, lab in fixed_recurring_italian_holidays()
        ]

    def get_public_holidays(self) -> list[dict[str, Any]]:
        """Lista {month, day, label} — ricorrenti ogni anno (senza anno)."""
        raw = (self.get_app_setting(self._PUBLIC_HOLIDAYS_KEY, default="") or "").strip()
        if not raw:
            # Mai salvato: stesso elenco predefinito della vista Agenda.
            return list(self._default_public_holiday_rows())
        try:
            data = json.loads(raw)
            if isinstance(data, list):
                if len(data) == 0:
                    # Salvataggio esplicito di lista vuota.
                    return []
                return self._normalize_public_holiday_rows([x for x in data if isinstance(x, dict)])
        except json.JSONDecodeError:
            pass
        return list(self._default_public_holiday_rows())

    def set_public_holidays(self, rows: list[dict[str, Any]]) -> None:
        clean = self._normalize_public_holiday_rows([r for r in rows if isinstance(r, dict)])
        self.set_app_setting(self._PUBLIC_HOLIDAYS_KEY, json.dumps(clean, ensure_ascii=False))

    # SQL (scheda principale)
    _SQL_ARCHIVE_QUERIES_KEY = "sql_archive_queries_v1"

    def _normalize_sql_archive_row(self, r: dict[str, Any]) -> dict[str, Any]:
        eid = str(r.get("id") or "").strip() or str(uuid.uuid4())
        name = str(r.get("name") or "").strip() or "Senza nome"
        sql_text = str(r.get("sql_text") or "")
        return {"id": eid, "name": name, "sql_text": sql_text}

    def get_sql_archive_queries(self) -> list[dict[str, Any]]:
        """Query SQL salvate dall'utente: id, name, sql_text."""
        raw = (self.get_app_setting(self._SQL_ARCHIVE_QUERIES_KEY, default="") or "").strip()
        if not raw:
            return []
        try:
            data = json.loads(raw)
            if not isinstance(data, list):
                return []
            by_id: dict[str, dict[str, Any]] = {}
            for x in data:
                if not isinstance(x, dict):
                    continue
                row = self._normalize_sql_archive_row(x)
                by_id[row["id"]] = row
            return sorted(by_id.values(), key=lambda z: z["name"].lower())
        except json.JSONDecodeError:
            return []

    def set_sql_archive_queries(self, rows: list[dict[str, Any]]) -> None:
        by_id: dict[str, dict[str, Any]] = {}
        for r in rows:
            if not isinstance(r, dict):
                continue
            row = self._normalize_sql_archive_row(r)
            by_id[row["id"]] = row
        clean = sorted(by_id.values(), key=lambda z: z["name"].lower())
        self.set_app_setting(self._SQL_ARCHIVE_QUERIES_KEY, json.dumps(clean, ensure_ascii=False))

