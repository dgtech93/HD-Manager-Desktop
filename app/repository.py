from __future__ import annotations

import datetime as dt
import os
import sqlite3
import time
from typing import Iterable

from app.database import get_connection
from app.services.secrets_service import SecretsService
from app.services.system_service import SystemService


class Repository:
    VPN_TYPES = {"Vpn Proprietario", "VPN Windows"}
    VPN_ACCESS_INFO_TYPES = frozenset({"Utente/Password", "File Configurato"})

    def __init__(self, connection_factory=get_connection) -> None:
        self.connection_factory = connection_factory
        self._windows_vpn_cache: list[str] = []
        self._windows_vpn_cache_ts: float = 0.0
        self._secrets = SecretsService()
        self._system = SystemService()

    # Generic helpers
    @staticmethod
    def _to_dicts(rows: Iterable[sqlite3.Row]) -> list[dict]:
        return [dict(row) for row in rows]

    @staticmethod
    def _req(value: str, label: str) -> str:
        cleaned = (value or "").strip()
        if not cleaned:
            raise ValueError(f"{label} obbligatorio.")
        return cleaned

    @staticmethod
    def _opt(value: str) -> str:
        return (value or "").strip()

    @staticmethod
    def _parse_bool(value: str | int | bool) -> int:
        if isinstance(value, bool):
            return int(value)
        if isinstance(value, int):
            return 1 if value else 0
        cleaned = (value or "").strip().lower()
        if cleaned in {"1", "si", "s", "true", "yes", "y"}:
            return 1
        if cleaned in {"0", "no", "n", "false", ""}:
            return 0
        raise ValueError(f"Valore booleano non valido: {value}")

    @staticmethod
    def _parse_display_order(value: str | int | None) -> int:
        cleaned = "" if value is None else str(value).strip()
        if not cleaned:
            raise ValueError("Ordine visualizzazione obbligatorio.")
        try:
            current = int(cleaned)
        except ValueError as exc:
            raise ValueError("Ordine visualizzazione deve essere un numero intero (1-20).") from exc
        if current < 1 or current > 20:
            raise ValueError("Ordine visualizzazione deve essere compreso tra 1 e 20.")
        return current

    @staticmethod
    def _date(value: str) -> str:
        cleaned = (value or "").strip()
        if not cleaned:
            return dt.date.today().isoformat()
        try:
            dt.date.fromisoformat(cleaned)
        except ValueError as exc:
            raise ValueError("Data non valida. Usa formato YYYY-MM-DD.") from exc
        return cleaned

    @staticmethod
    def _csv_names(value: str | Iterable[str] | None) -> list[str]:
        if value is None:
            return []
        chunks = value.split(",") if isinstance(value, str) else list(value)
        out: list[str] = []
        seen: set[str] = set()
        for chunk in chunks:
            name = (chunk or "").strip()
            key = name.lower()
            if name and key not in seen:
                out.append(name)
                seen.add(key)
        return out

    @staticmethod
    def _normalize_ids(ids: Iterable[int] | None) -> list[int]:
        out: list[int] = []
        seen: set[int] = set()
        for value in ids or []:
            current = int(value)
            if current not in seen:
                seen.add(current)
                out.append(current)
        return out

    def _query(self, sql: str, params: tuple = ()) -> list[dict]:
        conn = self.connection_factory()
        try:
            return self._to_dicts(conn.execute(sql, params).fetchall())
        finally:
            conn.close()

    def _execute(self, sql: str, params: tuple = (), label: str = "Operazione") -> None:
        conn = self.connection_factory()
        try:
            with conn:
                conn.execute(sql, params)
        except sqlite3.IntegrityError as exc:
            raise ValueError(f"{label} fallita: {exc}") from exc
        finally:
            conn.close()

    def _insert(self, sql: str, params: tuple, label: str = "Inserimento") -> int:
        conn = self.connection_factory()
        try:
            with conn:
                cur = conn.execute(sql, params)
                return int(cur.lastrowid)
        except sqlite3.IntegrityError as exc:
            raise ValueError(f"{label} fallito: {exc}") from exc
        finally:
            conn.close()

    def _lookup_map(self, sql: str, params: tuple = (), key_col: str = "label") -> dict[str, int]:
        rows = self._query(sql, params)
        return {str(r[key_col]).strip().lower(): int(r["id"]) for r in rows}

    def _name_to_id(
        self, value: str, mapping: dict[str, int], label: str, allow_empty: bool = False
    ) -> int | None:
        cleaned = (value or "").strip()
        if not cleaned:
            if allow_empty:
                return None
            raise ValueError(f"{label} obbligatorio.")
        current = mapping.get(cleaned.lower())
        if current is None:
            raise ValueError(f"{label} '{cleaned}' non trovato.")
        return current

    def _names_to_ids(self, value: str | Iterable[str] | None, mapping: dict[str, int], label: str) -> list[int]:
        out: list[int] = []
        for name in self._csv_names(value):
            current = mapping.get(name.lower())
            if current is None:
                raise ValueError(f"{label} '{name}' non trovato.")
            if current not in out:
                out.append(current)
        return out

    # Lookups
    def list_competences_lookup(self) -> list[dict]:
        return self._query("SELECT id, name AS label FROM competences ORDER BY name;")

    def list_product_types_lookup(self) -> list[dict]:
        return self._query("SELECT id, name AS label FROM product_types ORDER BY name;")

    def list_releases_lookup(self) -> list[dict]:
        return self._query("SELECT id, name AS label FROM releases ORDER BY name;")

    def list_environments_lookup(self) -> list[dict]:
        return self._query("SELECT id, name AS label FROM environments ORDER BY name;")

    def list_roles_lookup(self) -> list[dict]:
        return self._query(
            """
            SELECT id, name AS label
            FROM roles
            ORDER BY
                CASE WHEN display_order IS NULL THEN 1 ELSE 0 END,
                display_order,
                name;
            """
        )

    def list_vpns_lookup(self) -> list[dict]:
        return self._query(
            "SELECT id, connection_name AS label FROM vpns ORDER BY connection_name;"
        )

    def list_resources_lookup(self) -> list[dict]:
        return self._query(
            "SELECT id, (name || ' ' || surname) AS label FROM resources ORDER BY name, surname;"
        )

    def list_clients_lookup(self) -> list[dict]:
        return self._query("SELECT id, name AS label FROM clients ORDER BY name;")

    def list_tags_lookup(self) -> list[dict]:
        return self._query("SELECT id, name AS label FROM tags ORDER BY name;")

    # Competenze
    def list_competences(self) -> list[dict]:
        return self._query(
            "SELECT id, ('CMP' || printf('%06d', id)) AS code, name FROM competences ORDER BY name;"
        )

    def list_competences_with_resources(self) -> list[dict]:
        """
        Mostra per ogni competenza l'elenco delle risorse collegate.
        Collegamento: resources.role_id -> roles.id, e roles.competence = competences.name.
        """
        return self._query(
            """
            SELECT
                c.id,
                ('CMP' || printf('%06d', c.id)) AS code,
                c.name,
                COALESCE(
                    (
                        SELECT group_concat(r2.name || ' ' || r2.surname, ', ')
                        FROM resources r2
                        JOIN roles ro2 ON ro2.id = r2.role_id
                        WHERE ro2.competence = c.name
                    ),
                    ''
                ) AS resources
            FROM competences c
            ORDER BY c.name;
            """
        )

    def assign_resources_to_competence(
        self,
        competence_name: str,
        selected_resource_labels: list[str],
    ) -> None:
        """
        Assegna (o disassegna) le risorse alla competenza.

        Logica:
        - Le risorse sono legate a ruoli (resources.role_id -> roles.id)
        - I ruoli appartengono a una competenza (roles.competence)
        - Per assegnare una risorsa a una competenza, la colleghiamo a un ruolo
          della stessa competenza (primo ruolo ordinato per name).
        - Se una risorsa non è selezionata, viene scollegata (role_id=NULL)
          SOLO se era già dentro la competenza.
        """

        competence = (competence_name or "").strip()
        if not competence:
            raise ValueError("Nome competenza non valido.")

        # Lookup: label -> resource_id
        resources_lookup = self._lookup_map(
            "SELECT id, (name || ' ' || surname) AS label FROM resources ORDER BY name, surname;",
            key_col="label",
        )

        selected_ids: set[int] = set()
        for label in selected_resource_labels:
            cleaned = (label or "").strip()
            if not cleaned:
                continue
            res_id = resources_lookup.get(cleaned.lower())
            if res_id is None:
                raise ValueError(f"Risorsa non trovata: '{cleaned}'.")
            selected_ids.add(int(res_id))

        role_rows = self._query(
            """
            SELECT id
            FROM roles
            WHERE competence = ?
            ORDER BY name;
            """,
            (competence,),
        )
        role_ids = [int(r["id"]) for r in role_rows]
        default_role_id = role_ids[0] if role_ids else None

        # Risorse già in questa competenza (quindi con ruoli della competenza)
        current_rows = self._query(
            """
            SELECT r.id
            FROM resources r
            JOIN roles ro ON ro.id = r.role_id
            WHERE ro.competence = ?;
            """,
            (competence,),
        )
        current_ids: set[int] = {int(r["id"]) for r in current_rows}

        conn = self.connection_factory()
        try:
            # Disassegna quelli che erano nella competenza ma non sono più selezionati
            to_unassign = current_ids - selected_ids
            if to_unassign:
                placeholders = ",".join("?" for _ in to_unassign)
                conn.execute(
                    f"UPDATE resources SET role_id=NULL WHERE id IN ({placeholders});",
                    tuple(to_unassign),
                )

            # Assegna quelli selezionati che non sono già nella competenza
            to_assign = selected_ids - current_ids
            if to_assign:
                if default_role_id is None:
                    raise ValueError(
                        "Non esiste alcun Ruolo per questa competenza. Crea almeno un ruolo prima di assegnare risorse."
                    )
                placeholders = ",".join("?" for _ in to_assign)
                conn.execute(
                    f"UPDATE resources SET role_id=? WHERE id IN ({placeholders});",
                    (default_role_id, *tuple(to_assign)),
                )
            conn.commit()
        finally:
            conn.close()

    def upsert_competence(self, row_id: int | None, name: str) -> int:
        clean = self._req(name, "Nome competenza")
        if row_id:
            self._execute(
                "UPDATE competences SET name = ? WHERE id = ?;",
                (clean, int(row_id)),
                "Aggiornamento competenza",
            )
            return int(row_id)
        return self._insert(
            "INSERT INTO competences(name) VALUES (?);",
            (clean,),
            "Inserimento competenza",
        )

    def delete_competence(self, row_id: int) -> None:
        self._execute(
            "DELETE FROM competences WHERE id = ?;",
            (int(row_id),),
            "Eliminazione competenza",
        )

    # Tag
    def list_tags(self) -> list[dict]:
        return self._query(
            """
            SELECT t.id,
                   ('TAG' || printf('%06d', t.id)) AS code,
                   t.name,
                   t.color,
                   c.name AS client_name
            FROM tags t
            LEFT JOIN clients c ON c.id = t.client_id
            ORDER BY t.name;
            """
        )

    @staticmethod
    def _normalize_color(value: str | None) -> str:
        cleaned = (value or "").strip()
        if not cleaned:
            return "#0f766e"
        if not cleaned.startswith("#"):
            cleaned = f"#{cleaned}"
        if len(cleaned) != 7:
            return "#0f766e"
        return cleaned.lower()

    def upsert_tag(
        self,
        row_id: int | None,
        name: str,
        color: str,
        client_name: str = "",
    ) -> int:
        clean_name = self._req(name, "Nome tag")
        clean_color = self._normalize_color(color)
        client_id = None
        clean_client = (client_name or "").strip()
        if clean_client:
            client_map = self._lookup_map("SELECT id, name AS label FROM clients ORDER BY name;")
            client_id = self._name_to_id(clean_client, client_map, "Cliente")
        if row_id:
            self._execute(
                "UPDATE tags SET name=?, color=?, client_id=? WHERE id=?;",
                (clean_name, clean_color, client_id, int(row_id)),
                "Aggiornamento tag",
            )
            return int(row_id)
        existing_id = self._tag_id_from_name(clean_name)
        if existing_id is not None:
            return int(existing_id)
        return self._insert(
            "INSERT INTO tags(name, color, client_id) VALUES (?, ?, ?);",
            (clean_name, clean_color, client_id),
            "Inserimento tag",
        )

    def delete_tag(self, row_id: int) -> None:
        self._execute(
            "DELETE FROM tags WHERE id = ?;",
            (int(row_id),),
            "Eliminazione tag",
        )

    def list_tags_for_client(self, client_id: int) -> list[dict]:
        return self._query(
            """
            SELECT id, name, color
            FROM tags
            WHERE client_id = ?
            ORDER BY name;
            """,
            (int(client_id),),
        )

    # Tipi prodotto
    def list_product_types(self) -> list[dict]:
        return self._query(
            """
            SELECT id,
                   ('TPR' || printf('%06d', id)) AS code,
                   name,
                   flag_ip,
                   flag_host,
                   flag_preconfigured,
                   flag_url,
                   flag_port
            FROM product_types
            ORDER BY name;
            """
        )

    def upsert_product_type(
        self,
        row_id: int | None,
        name: str,
        flag_ip: str | int | bool,
        flag_host: str | int | bool,
        flag_preconfigured: str | int | bool,
        flag_url: str | int | bool,
        flag_port: str | int | bool,
    ) -> int:
        clean = self._req(name, "Nome tipo prodotto")
        payload = (
            clean,
            self._parse_bool(flag_ip),
            self._parse_bool(flag_host),
            self._parse_bool(flag_preconfigured),
            self._parse_bool(flag_url),
            self._parse_bool(flag_port),
        )
        if row_id:
            self._execute(
                """
                UPDATE product_types
                SET name=?, flag_ip=?, flag_host=?, flag_preconfigured=?, flag_url=?, flag_port=?
                WHERE id=?;
                """,
                payload + (int(row_id),),
                "Aggiornamento tipo prodotto",
            )
            return int(row_id)
        return self._insert(
            """
            INSERT INTO product_types(name, flag_ip, flag_host, flag_preconfigured, flag_url, flag_port)
            VALUES (?, ?, ?, ?, ?, ?);
            """,
            payload,
            "Inserimento tipo prodotto",
        )

    def delete_product_type(self, row_id: int) -> None:
        self._execute(
            "DELETE FROM product_types WHERE id = ?;",
            (int(row_id),),
            "Eliminazione tipo prodotto",
        )

    # Release
    def list_releases(self) -> list[dict]:
        return self._query(
            "SELECT id, ('RLS' || printf('%06d', id)) AS code, name, release_date FROM releases ORDER BY release_date DESC, name;"
        )

    def upsert_release(self, row_id: int | None, name: str, release_date: str) -> int:
        clean_name = self._req(name, "Nome release")
        clean_date = self._date(release_date)
        if row_id:
            self._execute(
                "UPDATE releases SET name=?, release_date=? WHERE id=?;",
                (clean_name, clean_date, int(row_id)),
                "Aggiornamento release",
            )
            return int(row_id)
        return self._insert(
            "INSERT INTO releases(name, release_date) VALUES (?, ?);",
            (clean_name, clean_date),
            "Inserimento release",
        )

    def delete_release(self, row_id: int) -> None:
        self._execute(
            "DELETE FROM releases WHERE id = ?;",
            (int(row_id),),
            "Eliminazione release",
        )

    # Ambienti
    def list_environments(self) -> list[dict]:
        return self._query(
            """
            SELECT e.id,
                   ('AMB' || printf('%06d', e.id)) AS code,
                   e.name,
                   r.name AS release_name,
                   r.release_date
            FROM environments e
            LEFT JOIN releases r ON r.id = e.release_id
            ORDER BY e.name;
            """
        )

    def _release_id(self, release_name: str) -> int | None:
        release_map = self._lookup_map("SELECT id, name AS label FROM releases ORDER BY name;")
        return self._name_to_id(release_name, release_map, "Release", allow_empty=True)

    def upsert_environment(self, row_id: int | None, name: str, release_name: str = "") -> int:
        clean_name = self._req(name, "Nome ambiente")
        release_id = self._release_id(release_name)
        if row_id:
            self._execute(
                "UPDATE environments SET name=?, release_id=? WHERE id=?;",
                (clean_name, release_id, int(row_id)),
                "Aggiornamento ambiente",
            )
            return int(row_id)
        return self._insert(
            "INSERT INTO environments(name, release_id) VALUES (?, ?);",
            (clean_name, release_id),
            "Inserimento ambiente",
        )

    def delete_environment(self, row_id: int) -> None:
        self._execute(
            "DELETE FROM environments WHERE id = ?;",
            (int(row_id),),
            "Eliminazione ambiente",
        )

    # Ruoli
    def list_roles(self) -> list[dict]:
        return self._query(
            """
            SELECT id,
                   ('ROL' || printf('%06d', id)) AS code,
                   name,
                   competence,
                   multi_clients,
                   display_order
            FROM roles
            ORDER BY
                CASE WHEN display_order IS NULL THEN 1 ELSE 0 END,
                display_order,
                name;
            """
        )

    def _assert_role_display_order_available(
        self, display_order: int, row_id: int | None = None
    ) -> None:
        if row_id:
            rows = self._query(
                "SELECT id FROM roles WHERE display_order=? AND id<>?;",
                (display_order, int(row_id)),
            )
        else:
            rows = self._query(
                "SELECT id FROM roles WHERE display_order=?;",
                (display_order,),
            )
        if rows:
            raise ValueError(
                f"Ordine visualizzazione {display_order} gia utilizzato da un altro ruolo."
            )

    def upsert_role(
        self,
        row_id: int | None,
        name: str,
        multi_clients: str | int | bool,
        competence: str = "",
        display_order: str | int | None = None,
    ) -> int:
        clean_name = self._req(name, "Nome ruolo")
        clean_comp = self._opt(competence)
        clean_multi = self._parse_bool(multi_clients)
        cleaned_order = ("" if display_order is None else str(display_order)).strip()
        if cleaned_order:
            clean_order = self._parse_display_order(display_order)
        else:
            clean_order = self._next_role_display_order()
        self._assert_role_display_order_available(clean_order, row_id)
        if row_id:
            self._execute(
                "UPDATE roles SET name=?, competence=?, multi_clients=?, display_order=? WHERE id=?;",
                (clean_name, clean_comp, clean_multi, clean_order, int(row_id)),
                "Aggiornamento ruolo",
            )
            return int(row_id)
        return self._insert(
            "INSERT INTO roles(name, competence, multi_clients, display_order) VALUES (?, ?, ?, ?);",
            (clean_name, clean_comp, clean_multi, clean_order),
            "Inserimento ruolo",
        )

    def delete_role(self, row_id: int) -> None:
        self._execute(
            "DELETE FROM roles WHERE id = ?;",
            (int(row_id),),
            "Eliminazione ruolo",
        )

    # VPN
    def list_vpns(self) -> list[dict]:
        rows = self._query(
            """
            SELECT id,
                   ('VPN' || printf('%06d', id)) AS code,
                   connection_name,
                   server_address,
                   vpn_type,
                   access_info_type,
                   username,
                   password,
                   password_ref,
                   vpn_path,
                   COALESCE(
                       (
                           SELECT group_concat(c.name, ', ')
                           FROM clients c
                           WHERE c.vpn_id = vpns.id
                       ),
                       ''
                   ) AS clients
            FROM vpns
            ORDER BY connection_name;
            """
        )
        # Resolve secret if stored in keyring
        for row in rows:
            ref = str(row.get("password_ref") or "").strip()
            if ref and self._secrets.available:
                secret = self._secrets.get_secret(self._secrets.vpn_password_ref(str(row.get("connection_name") or "")))
                if secret:
                    row["password"] = secret
        return rows

    def list_windows_vpn_connections(self) -> list[str]:
        now = time.monotonic()
        if self._windows_vpn_cache and (now - self._windows_vpn_cache_ts) < 15:
            return list(self._windows_vpn_cache)
        connections = self._system.list_windows_vpn_connections()
        self._windows_vpn_cache = connections
        self._windows_vpn_cache_ts = now
        return list(connections)

    def upsert_vpn(
        self,
        row_id: int | None,
        connection_name: str,
        server_address: str,
        vpn_type: str,
        access_info_type: str,
        username: str,
        password: str,
        vpn_path: str = "",
        clients: str | Iterable[str] | None = None,
    ) -> int:
        clean_vpn_type = self._req(vpn_type, "Tipo VPN")
        if clean_vpn_type not in self.VPN_TYPES:
            allowed = ", ".join(sorted(self.VPN_TYPES))
            raise ValueError(f"Tipo VPN non valido. Valori ammessi: {allowed}.")

        if clean_vpn_type == "VPN Windows":
            windows_vpns = self.list_windows_vpn_connections()
            if not windows_vpns:
                raise ValueError("Nessuna VPN Windows configurata trovata nel sistema.")
            clean_name = self._req(connection_name, "Nome connessione")
            if clean_name not in windows_vpns:
                raise ValueError(
                    "Nome connessione non trovato tra le VPN Windows configurate."
                )
            clean_path = ""
        else:
            clean_name = self._req(connection_name, "Nome connessione")
            clean_path = self._req(vpn_path, "Percorso VPN")

        clean_access = self._opt(access_info_type)
        if clean_access and clean_access not in self.VPN_ACCESS_INFO_TYPES:
            raise ValueError(
                "Tipo Info Accesso non valido. Valori ammessi: Utente/Password, File Configurato."
            )

        payload = (
            clean_name,
            self._req(server_address, "Indirizzo server"),
            clean_vpn_type,
            clean_access,
            self._req(username, "Nome utente"),
            self._req(password, "Password"),
            self._opt(clean_path),
        )
        client_ids = self._client_ids(clients)
        password_ref_value: str | None = None
        if self._secrets.available:
            ref = self._secrets.vpn_password_ref(clean_name)
            self._secrets.set_secret(ref, self._req(password, "Password"))
            password_ref_value = ref.serialize()

        conn = self.connection_factory()
        try:
            with conn:
                if row_id:
                    conn.execute(
                        """
                        UPDATE vpns
                        SET connection_name=?, server_address=?, vpn_type=?, access_info_type=?, username=?, password=?, password_ref=?, vpn_path=?
                        WHERE id=?;
                        """,
                        payload[:-1] + (password_ref_value, payload[-1]) + (int(row_id),),
                    )
                    vpn_id = int(row_id)
                else:
                    cur = conn.execute(
                        """
                        INSERT INTO vpns(connection_name, server_address, vpn_type, access_info_type, username, password, password_ref, vpn_path)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?);
                        """,
                        payload[:-1] + (password_ref_value, payload[-1]),
                    )
                    vpn_id = int(cur.lastrowid)

                # Reset previous assignments, then set selected clients to this VPN.
                conn.execute("UPDATE clients SET vpn_id=NULL WHERE vpn_id=?;", (vpn_id,))
                for client_id in client_ids:
                    conn.execute(
                        "UPDATE clients SET vpn_id=? WHERE id=?;",
                        (vpn_id, int(client_id)),
                    )
                return vpn_id
        except sqlite3.IntegrityError as exc:
            action = "aggiornare" if row_id else "inserire"
            raise ValueError(f"Impossibile {action} VPN: {exc}") from exc
        finally:
            conn.close()

    def delete_vpn(self, row_id: int) -> None:
        self._execute(
            "DELETE FROM vpns WHERE id = ?;",
            (int(row_id),),
            "Eliminazione VPN",
        )

    # Risorse
    def list_resources(self) -> list[dict]:
        return self._query(
            """
            SELECT rs.id,
                   ('RES' || printf('%06d', rs.id)) AS code,
                   rs.name,
                   rs.surname,
                   rl.name AS role_name,
                   cm.name AS competence_name,
                   rs.phone,
                   rs.email,
                   rs.linkedin,
                   rs.photo_link,
                   rs.note
            FROM resources rs
            LEFT JOIN roles rl ON rl.id = rs.role_id
            LEFT JOIN competences cm ON cm.id = rs.competence_id
            ORDER BY rs.name, rs.surname;
            """
        )

    def _role_id(self, role_name: str) -> int | None:
        role_map = self._lookup_map("SELECT id, name AS label FROM roles ORDER BY name;")
        return self._name_to_id(role_name, role_map, "Ruolo", allow_empty=False)

    def _competence_id(self, competence_name: str) -> int | None:
        mapping = self._lookup_map("SELECT id, name AS label FROM competences ORDER BY name;")
        return self._name_to_id(competence_name, mapping, "Competenza", allow_empty=True)

    def upsert_resource(
        self,
        row_id: int | None,
        name: str,
        surname: str,
        role_name: str = "",
        competence_name: str = "",
        phone: str = "",
        email: str = "",
        linkedin: str = "",
        photo_link: str = "",
        note: str = "",
    ) -> int:
        payload = (
            self._req(name, "Nome risorsa"),
            self._req(surname, "Cognome risorsa"),
            self._role_id(role_name),
            self._competence_id(competence_name),
            self._opt(phone),
            self._opt(email),
            self._opt(linkedin),
            self._opt(photo_link),
            self._opt(note),
        )
        if row_id:
            self._execute(
                """
                UPDATE resources
                SET name=?, surname=?, role_id=?, competence_id=?, phone=?, email=?, linkedin=?, photo_link=?, note=?
                WHERE id=?;
                """,
                payload + (int(row_id),),
                "Aggiornamento risorsa",
            )
            return int(row_id)
        return self._insert(
            """
            INSERT INTO resources(name, surname, role_id, competence_id, phone, email, linkedin, photo_link, note)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);
            """,
            payload,
            "Inserimento risorsa",
        )

    def delete_resource(self, row_id: int) -> None:
        self._execute(
            "DELETE FROM resources WHERE id = ?;",
            (int(row_id),),
            "Eliminazione risorsa",
        )

    # Clienti
    def list_clients(self) -> list[dict]:
        return self._query(
            """
            SELECT c.id,
                   ('CLI' || printf('%06d', c.id)) AS code,
                   c.name,
                   c.location,
                   c.link,
                   v.connection_name AS vpn_name,
                   COALESCE(
                       (
                           SELECT group_concat(r.name || ' ' || r.surname, ', ')
                           FROM client_resources cr
                           JOIN resources r ON r.id = cr.resource_id
                           WHERE cr.client_id = c.id
                       ),
                       ''
                   ) AS resources
            FROM clients c
            LEFT JOIN vpns v ON v.id = c.vpn_id
            ORDER BY c.name;
            """
        )

    def _vpn_id(self, vpn_name: str) -> int | None:
        vpn_map = self._lookup_map(
            "SELECT id, connection_name AS label FROM vpns ORDER BY connection_name;"
        )
        return self._name_to_id(vpn_name, vpn_map, "VPN associata", allow_empty=True)

    def _resource_ids(self, resources: str | Iterable[str] | None) -> list[int]:
        resources_map = self._lookup_map(
            "SELECT id, (name || ' ' || surname) AS label FROM resources ORDER BY name, surname;"
        )
        return self._names_to_ids(resources, resources_map, "Risorsa")

    def upsert_client(
        self,
        row_id: int | None,
        name: str,
        location: str,
        vpn_name: str = "",
        resources: str | Iterable[str] | None = None,
        link: str = "",
    ) -> int:
        clean_name = self._req(name, "Nome cliente")
        clean_location = self._req(location, "Localita")
        clean_link = self._opt(link)
        vpn_id = self._vpn_id(vpn_name)
        resource_ids = self._resource_ids(resources)

        conn = self.connection_factory()
        try:
            with conn:
                if row_id:
                    conn.execute(
                        "UPDATE clients SET name=?, location=?, link=?, vpn_id=? WHERE id=?;",
                        (clean_name, clean_location, clean_link, vpn_id, int(row_id)),
                    )
                    client_id = int(row_id)
                    conn.execute(
                        "DELETE FROM client_resources WHERE client_id=?;",
                        (client_id,),
                    )
                else:
                    cur = conn.execute(
                        "INSERT INTO clients(name, location, link, vpn_id) VALUES (?, ?, ?, ?);",
                        (clean_name, clean_location, clean_link, vpn_id),
                    )
                    client_id = int(cur.lastrowid)

                for resource_id in resource_ids:
                    conn.execute(
                        "INSERT INTO client_resources(client_id, resource_id) VALUES (?, ?);",
                        (client_id, resource_id),
                    )
                return client_id
        except sqlite3.IntegrityError as exc:
            action = "aggiornare" if row_id else "inserire"
            raise ValueError(f"Impossibile {action} cliente: {exc}") from exc
        finally:
            conn.close()

    def delete_client(self, row_id: int) -> None:
        self._execute(
            "DELETE FROM clients WHERE id = ?;",
            (int(row_id),),
            "Eliminazione cliente",
        )

    # Prodotti
    def list_products(self) -> list[dict]:
        return self._query(
            """
            SELECT p.id,
                   ('PRD' || printf('%06d', p.id)) AS code,
                   p.name,
                   pt.name AS product_type,
                   COALESCE(
                       (
                           SELECT group_concat(c.name, ', ')
                           FROM product_clients pc
                           JOIN clients c ON c.id = pc.client_id
                           WHERE pc.product_id = p.id
                       ),
                       ''
                   ) AS clients,
                   COALESCE(
                       (
                           SELECT group_concat(e.name, ', ')
                           FROM product_environments pe
                           JOIN environments e ON e.id = pe.environment_id
                           WHERE pe.product_id = p.id
                       ),
                       ''
                   ) AS environments
            FROM products p
            JOIN product_types pt ON pt.id = p.product_type_id
            ORDER BY p.name;
            """
        )

    def list_client_product_environment_releases(self, client_id: int) -> list[dict]:
        products = self._query(
            """
            SELECT p.id,
                   p.name,
                   COALESCE(
                       (
                           SELECT group_concat(e.name, ', ')
                           FROM product_environments pe
                           JOIN environments e ON e.id = pe.environment_id
                           WHERE pe.product_id = p.id
                       ),
                       ''
                   ) AS environments
            FROM products p
            JOIN product_clients pc ON pc.product_id = p.id
            WHERE pc.client_id = ?
            ORDER BY p.name;
            """,
            (int(client_id),),
        )
        if not products:
            return []

        credential_rows = self._query(
            """
            SELECT pc.product_id,
                   e.name AS environment_name,
                   r.name AS release_name
            FROM product_credentials pc
            JOIN product_credential_environments pce ON pce.credential_id = pc.id
            JOIN environments e ON e.id = pce.environment_id
            LEFT JOIN releases r ON r.id = pce.release_id
            WHERE pc.client_id = ?
            ORDER BY pc.product_id, e.name, r.name;
            """,
            (int(client_id),),
        )

        releases_by_product: dict[int, dict[str, set[str]]] = {}
        for row in credential_rows:
            product_id = int(row["product_id"])
            env_name = str(row.get("environment_name") or "").strip()
            release_name = str(row.get("release_name") or "").strip()
            if not env_name:
                continue
            product_map = releases_by_product.setdefault(product_id, {})
            env_releases = product_map.setdefault(env_name, set())
            if release_name:
                env_releases.add(release_name)

        out: list[dict] = []
        for product in products:
            product_id = int(product["id"])
            configured_envs = self._csv_names(product.get("environments"))
            credential_env_map = releases_by_product.get(product_id, {})

            env_order: list[str] = []
            seen: set[str] = set()
            for env_name in configured_envs + sorted(credential_env_map.keys(), key=str.lower):
                key = env_name.lower()
                if key in seen:
                    continue
                seen.add(key)
                env_order.append(env_name)

            pairs: list[str] = []
            for env_name in env_order:
                releases = sorted(credential_env_map.get(env_name, set()), key=str.lower)
                release_text = " / ".join(releases) if releases else "-"
                pairs.append(f"{env_name} - {release_text}")

            out.append(
                {
                    "product_id": product_id,
                    "product_name": product.get("name", ""),
                    "pairs": pairs,
                    "summary": ", ".join(pairs) if pairs else "Nessun ambiente associato",
                }
            )
        return out

    def _product_type_id(self, product_type_name: str) -> int:
        product_types_map = self._lookup_map(
            "SELECT id, name AS label FROM product_types ORDER BY name;"
        )
        return int(self._name_to_id(product_type_name, product_types_map, "Tipo prodotto"))

    def _client_ids(self, clients: str | Iterable[str] | None) -> list[int]:
        clients_map = self._lookup_map("SELECT id, name AS label FROM clients ORDER BY name;")
        return self._names_to_ids(clients, clients_map, "Cliente")

    def _environment_ids(self, environments: str | Iterable[str] | None) -> list[int]:
        env_map = self._lookup_map("SELECT id, name AS label FROM environments ORDER BY name;")
        return self._names_to_ids(environments, env_map, "Ambiente")

    def _opt_date(self, value: str) -> str | None:
        cleaned = (value or "").strip()
        if not cleaned:
            return None
        return self._date(cleaned)

    def upsert_product(
        self,
        row_id: int | None,
        name: str,
        product_type: str,
        clients: str | Iterable[str] | None = None,
        environments: str | Iterable[str] | None = None,
    ) -> int:
        clean_name = self._req(name, "Nome prodotto")
        product_type_id = self._product_type_id(product_type)
        client_ids = self._client_ids(clients)
        environment_ids = self._environment_ids(environments)

        conn = self.connection_factory()
        try:
            with conn:
                if row_id:
                    conn.execute(
                        "UPDATE products SET name=?, product_type_id=? WHERE id=?;",
                        (clean_name, product_type_id, int(row_id)),
                    )
                    product_id = int(row_id)
                    conn.execute(
                        "DELETE FROM product_clients WHERE product_id=?;",
                        (product_id,),
                    )
                    conn.execute(
                        "DELETE FROM product_environments WHERE product_id=?;",
                        (product_id,),
                    )
                else:
                    cur = conn.execute(
                        "INSERT INTO products(name, product_type_id) VALUES (?, ?);",
                        (clean_name, product_type_id),
                    )
                    product_id = int(cur.lastrowid)

                for client_id in client_ids:
                    conn.execute(
                        "INSERT INTO product_clients(product_id, client_id) VALUES (?, ?);",
                        (product_id, client_id),
                    )
                for environment_id in environment_ids:
                    conn.execute(
                        "INSERT INTO product_environments(product_id, environment_id) VALUES (?, ?);",
                        (product_id, environment_id),
                    )
                return product_id
        except sqlite3.IntegrityError as exc:
            action = "aggiornare" if row_id else "inserire"
            raise ValueError(f"Impossibile {action} prodotto: {exc}") from exc
        finally:
            conn.close()

    def delete_product(self, row_id: int) -> None:
        self._execute(
            "DELETE FROM products WHERE id = ?;",
            (int(row_id),),
            "Eliminazione prodotto",
        )

    def get_product_type_flags_for_product(self, product_id: int) -> dict:
        rows = self._query(
            """
            SELECT p.id,
                   p.name,
                   pt.name AS product_type_name,
                   pt.flag_ip,
                   pt.flag_url,
                   pt.flag_host,
                   pt.flag_preconfigured,
                   pt.flag_port
            FROM products p
            JOIN product_types pt ON pt.id = p.product_type_id
            WHERE p.id = ?;
            """,
            (int(product_id),),
        )
        if not rows:
            raise ValueError("Prodotto non trovato.")
        return rows[0]

    def _normalize_product_credential_payload(
        self,
        credential_name: str,
        environment_versions: list[tuple[str, str]],
        domain: str,
        login_name: str,
        username: str,
        password: str,
        ip: str = "",
        url: str = "",
        host: str = "",
        rdp_path: str = "",
        port: str | int | None = None,
        password_expiry: bool = False,
        password_inserted_at: str | None = None,
        password_duration_days: str | int | None = None,
        password_end_date: str | None = None,
        note: str = "",
    ) -> dict:
        clean_name = self._req(credential_name, "Nome credenziale")
        clean_domain = self._req(domain, "Dominio")
        clean_login = self._req(login_name, "Nome utente")
        clean_user = self._req(username, "Username")
        clean_password = self._req(password, "Password")

        if not environment_versions:
            raise ValueError("Seleziona almeno un ambiente.")

        env_map = self._lookup_map("SELECT id, name AS label FROM environments ORDER BY name;")
        rel_map = self._lookup_map("SELECT id, name AS label FROM releases ORDER BY name;")

        env_payload: list[tuple[int, int | None]] = []
        seen_env: set[int] = set()
        for environment_name, release_name in environment_versions:
            environment_id = self._name_to_id(environment_name, env_map, "Ambiente")
            release_id = self._name_to_id(release_name, rel_map, "Versione", allow_empty=True)
            assert environment_id is not None
            if environment_id in seen_env:
                continue
            seen_env.add(environment_id)
            env_payload.append((environment_id, release_id))

        clean_port: int | None = None
        if port is not None and str(port).strip():
            try:
                clean_port = int(str(port).strip())
            except ValueError as exc:
                raise ValueError("Porta non valida.") from exc
            if clean_port < 1 or clean_port > 65535:
                raise ValueError("Porta deve essere compresa tra 1 e 65535.")

        clean_expiry = 1 if password_expiry else 0
        clean_inserted_at: str | None = None
        clean_duration: int | None = None
        clean_end_date: str | None = None
        if clean_expiry:
            clean_inserted_at = self._date(password_inserted_at or "")
            if password_duration_days is not None and str(password_duration_days).strip():
                try:
                    clean_duration = int(str(password_duration_days).strip())
                except ValueError as exc:
                    raise ValueError("Durata password non valida.") from exc
                if clean_duration <= 0:
                    raise ValueError("Durata password deve essere maggiore di 0.")
            clean_end_date = self._opt_date(password_end_date or "")
            if clean_duration is None and not clean_end_date:
                raise ValueError(
                    "Con Password con scadenza attiva devi compilare Durata o Data fine."
                )

        return {
            "credential_name": clean_name,
            "domain": clean_domain,
            "login_name": clean_login,
            "username": clean_user,
            "password": clean_password,
            "password_ref": None,
            "ip": self._opt(ip),
            "url": self._opt(url),
            "host": self._opt(host),
            "rdp_path": self._opt(rdp_path),
            "port": clean_port,
            "password_expiry": clean_expiry,
            "password_inserted_at": clean_inserted_at,
            "password_duration_days": clean_duration,
            "password_end_date": clean_end_date,
            "note": self._opt(note),
            "environment_payload": env_payload,
        }

    def create_product_credential(
        self,
        client_id: int,
        product_id: int,
        credential_name: str,
        environment_versions: list[tuple[str, str]],
        domain: str,
        login_name: str,
        username: str,
        password: str,
        ip: str = "",
        url: str = "",
        host: str = "",
        rdp_path: str = "",
        port: str | int | None = None,
        password_expiry: bool = False,
        password_inserted_at: str | None = None,
        password_duration_days: str | int | None = None,
        password_end_date: str | None = None,
        note: str = "",
    ) -> int:
        payload = self._normalize_product_credential_payload(
            credential_name=credential_name,
            environment_versions=environment_versions,
            domain=domain,
            login_name=login_name,
            username=username,
            password=password,
            ip=ip,
            url=url,
            host=host,
            rdp_path=rdp_path,
            port=port,
            password_expiry=password_expiry,
            password_inserted_at=password_inserted_at,
            password_duration_days=password_duration_days,
            password_end_date=password_end_date,
            note=note,
        )
        if self._secrets.available:
            ref = self._secrets.credential_password_ref(int(client_id), int(product_id), payload["credential_name"])
            self._secrets.set_secret(ref, payload["password"])
            payload["password_ref"] = ref.serialize()

        conn = self.connection_factory()
        try:
            with conn:
                cur = conn.execute(
                    """
                    INSERT INTO product_credentials(
                        client_id,
                        product_id,
                        credential_name,
                        ip,
                        url,
                        host,
                        rdp_path,
                        domain,
                        login_name,
                        username,
                        password,
                        password_ref,
                        port,
                        password_expiry,
                        password_inserted_at,
                        password_duration_days,
                        password_end_date,
                        note
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
                    """,
                    (
                        int(client_id),
                        int(product_id),
                        payload["credential_name"],
                        payload["ip"],
                        payload["url"],
                        payload["host"],
                        payload["rdp_path"],
                        payload["domain"],
                        payload["login_name"],
                        payload["username"],
                        payload["password"],
                        payload["password_ref"],
                        payload["port"],
                        payload["password_expiry"],
                        payload["password_inserted_at"],
                        payload["password_duration_days"],
                        payload["password_end_date"],
                        payload["note"],
                    ),
                )
                credential_id = int(cur.lastrowid)

                for environment_id, release_id in payload["environment_payload"]:
                    conn.execute(
                        """
                        INSERT INTO product_credential_environments(
                            credential_id,
                            environment_id,
                            release_id
                        ) VALUES (?, ?, ?);
                        """,
                        (credential_id, environment_id, release_id),
                    )
                return credential_id
        except sqlite3.IntegrityError as exc:
            raise ValueError(f"Inserimento credenziale fallito: {exc}") from exc
        finally:
            conn.close()

    def get_product_credential_detail(self, credential_id: int) -> dict:
        rows = self._query(
            """
            SELECT id,
                   client_id,
                   product_id,
                   credential_name,
                   ip,
                   url,
                   host,
                   rdp_path,
                   domain,
                   login_name,
                   username,
                   password,
                   port,
                   password_expiry,
                   password_inserted_at,
                   password_duration_days,
                   password_end_date,
                   note
            FROM product_credentials
            WHERE id = ?;
            """,
            (int(credential_id),),
        )
        if not rows:
            raise ValueError("Credenziale non trovata.")

        row = rows[0]
        # Resolve password if stored in keyring
        if self._secrets.available:
            ref_val = str(row.get("password_ref") or "").strip()
            if ref_val:
                secret = self._secrets.get_secret(
                    self._secrets.credential_password_ref(
                        int(row.get("client_id")), int(row.get("product_id")), str(row.get("credential_name") or "")
                    )
                )
                if secret:
                    row["password"] = secret
        env_rows = self._query(
            """
            SELECT e.name AS environment,
                   r.name AS release
            FROM product_credential_environments pce
            JOIN environments e ON e.id = pce.environment_id
            LEFT JOIN releases r ON r.id = pce.release_id
            WHERE pce.credential_id = ?
            ORDER BY e.name;
            """,
            (int(credential_id),),
        )
        row["environment_versions"] = [
            (str(item.get("environment") or "").strip(), str(item.get("release") or "").strip())
            for item in env_rows
            if str(item.get("environment") or "").strip()
        ]
        return row

    def update_product_credential(
        self,
        credential_id: int,
        credential_name: str,
        environment_versions: list[tuple[str, str]],
        domain: str,
        login_name: str,
        username: str,
        password: str,
        ip: str = "",
        url: str = "",
        host: str = "",
        rdp_path: str = "",
        port: str | int | None = None,
        password_expiry: bool = False,
        password_inserted_at: str | None = None,
        password_duration_days: str | int | None = None,
        password_end_date: str | None = None,
        note: str = "",
    ) -> None:
        payload = self._normalize_product_credential_payload(
            credential_name=credential_name,
            environment_versions=environment_versions,
            domain=domain,
            login_name=login_name,
            username=username,
            password=password,
            ip=ip,
            url=url,
            host=host,
            rdp_path=rdp_path,
            port=port,
            password_expiry=password_expiry,
            password_inserted_at=password_inserted_at,
            password_duration_days=password_duration_days,
            password_end_date=password_end_date,
            note=note,
        )
        if self._secrets.available:
            ref = self._secrets.credential_password_ref(
                int(self.get_product_credential_detail(int(credential_id))["client_id"]),
                int(self.get_product_credential_detail(int(credential_id))["product_id"]),
                payload["credential_name"],
            )
            self._secrets.set_secret(ref, payload["password"])
            payload["password_ref"] = ref.serialize()

        conn = self.connection_factory()
        try:
            with conn:
                cur = conn.execute(
                    """
                    UPDATE product_credentials
                    SET credential_name=?,
                        ip=?,
                        url=?,
                        host=?,
                        rdp_path=?,
                        domain=?,
                        login_name=?,
                        username=?,
                        password=?,
                        password_ref=?,
                        port=?,
                        password_expiry=?,
                        password_inserted_at=?,
                        password_duration_days=?,
                        password_end_date=?,
                        note=?
                    WHERE id=?;
                    """,
                    (
                        payload["credential_name"],
                        payload["ip"],
                        payload["url"],
                        payload["host"],
                        payload["rdp_path"],
                        payload["domain"],
                        payload["login_name"],
                        payload["username"],
                        payload["password"],
                        payload["password_ref"],
                        payload["port"],
                        payload["password_expiry"],
                        payload["password_inserted_at"],
                        payload["password_duration_days"],
                        payload["password_end_date"],
                        payload["note"],
                        int(credential_id),
                    ),
                )
                if cur.rowcount == 0:
                    raise ValueError("Credenziale non trovata.")

                conn.execute(
                    "DELETE FROM product_credential_environments WHERE credential_id=?;",
                    (int(credential_id),),
                )
                for environment_id, release_id in payload["environment_payload"]:
                    conn.execute(
                        """
                        INSERT INTO product_credential_environments(
                            credential_id,
                            environment_id,
                            release_id
                        ) VALUES (?, ?, ?);
                        """,
                        (int(credential_id), environment_id, release_id),
                    )
        except sqlite3.IntegrityError as exc:
            raise ValueError(f"Aggiornamento credenziale fallito: {exc}") from exc
        finally:
            conn.close()

    def delete_product_credential(self, credential_id: int) -> None:
        self._execute(
            "DELETE FROM product_credentials WHERE id = ?;",
            (int(credential_id),),
            "Eliminazione credenziale prodotto",
        )

    def list_product_credentials(self, client_id: int, product_id: int) -> list[dict]:
        rows = self._query(
            """
            SELECT id,
                   credential_name,
                   ip,
                   url,
                   host,
                   rdp_path,
                   domain,
                   login_name,
                   username,
                   password,
                   password_ref,
                   port,
                   password_expiry,
                   password_inserted_at,
                   password_duration_days,
                   password_end_date,
                   COALESCE(
                       (
                           SELECT group_concat(
                               e.name || CASE
                                   WHEN r.name IS NOT NULL THEN ' [' || r.name || ']'
                                   ELSE ''
                               END,
                               ', '
                           )
                           FROM product_credential_environments pce
                           JOIN environments e ON e.id = pce.environment_id
                           LEFT JOIN releases r ON r.id = pce.release_id
                           WHERE pce.credential_id = product_credentials.id
                       ),
                       ''
                   ) AS environments_versions,
                   note
            FROM product_credentials
            WHERE client_id = ? AND product_id = ?
            ORDER BY id DESC;
            """,
            (int(client_id), int(product_id)),
        )
        if self._secrets.available:
            for row in rows:
                ref_val = str(row.get("password_ref") or "").strip()
                if not ref_val:
                    continue
                secret = self._secrets.get_secret(
                    self._secrets.credential_password_ref(int(client_id), int(product_id), str(row.get("credential_name") or ""))
                )
                if secret:
                    row["password"] = secret
        return rows

    # Archivio cartelle e file
    def list_archive_folders(self) -> list[dict]:
        return self._query(
            "SELECT id, name, parent_id FROM archive_folders ORDER BY name;"
        )

    def add_archive_folder(self, name: str, parent_id: int | None = None) -> int:
        clean_name = self._req(name, "Nome cartella")
        return self._insert(
            "INSERT INTO archive_folders(name, parent_id) VALUES (?, ?);",
            (clean_name, parent_id),
            "Inserimento cartella",
        )

    def delete_archive_folder(self, folder_id: int) -> None:
        """Delete a folder and all nested folders/files/links."""
        folder_id = int(folder_id)
        conn = self.connection_factory()
        try:
            with conn:
                # Collect all descendant folder IDs (including self).
                rows = conn.execute(
                    """
                    WITH RECURSIVE folders AS (
                        SELECT id
                        FROM archive_folders
                        WHERE id = ?
                        UNION ALL
                        SELECT af.id
                        FROM archive_folders af
                        JOIN folders f ON af.parent_id = f.id
                    )
                    SELECT id FROM folders;
                    """,
                    (folder_id,),
                ).fetchall()
                ids = [int(row[0]) for row in rows]
                if ids:
                    placeholders = ",".join("?" for _ in ids)
                    conn.execute(
                        f"DELETE FROM archive_files WHERE folder_id IN ({placeholders});",
                        tuple(ids),
                    )
                    conn.execute(
                        f"DELETE FROM archive_links WHERE folder_id IN ({placeholders});",
                        tuple(ids),
                    )
                    conn.execute(
                        f"DELETE FROM archive_folders WHERE id IN ({placeholders});",
                        tuple(ids),
                    )
        except sqlite3.IntegrityError as exc:
            raise ValueError(f"Eliminazione cartella fallita: {exc}") from exc
        finally:
            conn.close()

    def _tag_id_from_name(self, tag_name: str | None) -> int | None:
        cleaned = (tag_name or "").strip()
        if not cleaned:
            return None
        lookup = self._lookup_map("SELECT id, name AS label FROM tags ORDER BY name;")
        return lookup.get(cleaned.lower())

    def list_archive_files(self, folder_id: int | None) -> list[dict]:
        if folder_id is None:
            where_clause = "af.folder_id IS NULL"
            params: tuple = ()
        else:
            where_clause = "af.folder_id = ?"
            params = (int(folder_id),)
        return self._query(
            f"""
            SELECT af.id,
                   af.folder_id,
                   af.name,
                   af.file_type,
                   af.last_modified,
                   af.file_size,
                   af.extension,
                   af.path,
                   t.name AS tag_name,
                   t.color AS tag_color
            FROM archive_files af
            LEFT JOIN tags t ON t.id = af.tag_id
            WHERE {where_clause}
            ORDER BY af.name;
            """,
            params,
        )

    def list_archive_file_extensions(self, folder_id: int | None) -> list[str]:
        if folder_id is None:
            where_clause = "folder_id IS NULL"
            params: tuple = ()
        else:
            where_clause = "folder_id = ?"
            params = (int(folder_id),)
        rows = self._query(
            f"""
            SELECT DISTINCT lower(trim(extension)) AS ext
            FROM archive_files
            WHERE {where_clause} AND extension IS NOT NULL AND trim(extension) <> ''
            ORDER BY ext;
            """,
            params,
        )
        return [str(row.get("ext") or "").strip() for row in rows if str(row.get("ext") or "").strip()]

    def list_archive_files_filtered(
        self,
        folder_id: int | None,
        *,
        extension: str | None = None,
        name_contains: str | None = None,
    ) -> list[dict]:
        where_parts: list[str] = []
        params: list[object] = []
        if folder_id is None:
            where_parts.append("af.folder_id IS NULL")
        else:
            where_parts.append("af.folder_id = ?")
            params.append(int(folder_id))

        clean_ext = (extension or "").strip().lower()
        if clean_ext and clean_ext != "tutte":
            where_parts.append("lower(af.extension) = ?")
            params.append(clean_ext)

        clean_name = (name_contains or "").strip().lower()
        if clean_name:
            where_parts.append("lower(af.name) LIKE ?")
            params.append(f"%{clean_name}%")

        where_clause = " AND ".join(where_parts) if where_parts else "1=1"
        return self._query(
            f"""
            SELECT af.id,
                   af.folder_id,
                   af.name,
                   af.file_type,
                   af.last_modified,
                   af.file_size,
                   af.extension,
                   af.path,
                   t.name AS tag_name,
                   t.color AS tag_color
            FROM archive_files af
            LEFT JOIN tags t ON t.id = af.tag_id
            WHERE {where_clause}
            ORDER BY af.name;
            """,
            tuple(params),
        )

    def list_archive_files_all(self) -> list[dict]:
        return self._query(
            """
            SELECT af.id,
                   af.folder_id,
                   af.name,
                   af.file_type,
                   af.last_modified,
                   af.file_size,
                   af.extension,
                   af.path,
                   t.name AS tag_name,
                   t.color AS tag_color
            FROM archive_files af
            LEFT JOIN tags t ON t.id = af.tag_id
            ORDER BY af.name;
            """
        )

    def add_archive_file(self, folder_id: int | None, file_path: str, tag_name: str | None = None) -> int:
        clean_path = self._req(file_path, "Percorso file")
        if not os.path.exists(clean_path):
            raise ValueError("File non trovato.")

        name = os.path.basename(clean_path)
        extension = os.path.splitext(clean_path)[1].lstrip(".")
        try:
            stats = os.stat(clean_path)
            last_modified = dt.datetime.fromtimestamp(stats.st_mtime).strftime("%Y-%m-%d %H:%M")
            file_size = int(stats.st_size)
        except OSError:
            last_modified = ""
            file_size = None

        tag_id = self._tag_id_from_name(tag_name)

        return self._insert(
            """
            INSERT INTO archive_files(
                folder_id,
                name,
                file_type,
                last_modified,
                file_size,
                extension,
                path,
                tag_id
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?);
            """,
            (
                folder_id,
                name,
                "File",
                last_modified,
                file_size,
                extension,
                clean_path,
                tag_id,
            ),
            "Inserimento file",
        )

    def update_archive_file_tag(self, file_id: int, tag_name: str | None) -> None:
        tag_id = self._tag_id_from_name(tag_name)
        self._execute(
            "UPDATE archive_files SET tag_id=? WHERE id=?;",
            (tag_id, int(file_id)),
            "Aggiornamento tag file",
        )

    def delete_archive_file(self, file_id: int) -> None:
        self._execute(
            "DELETE FROM archive_files WHERE id = ?;",
            (int(file_id),),
            "Eliminazione file",
        )

    def move_archive_file(self, file_id: int, folder_id: int) -> None:
        file_id = int(file_id)
        folder_id = int(folder_id)

        # Validate destination folder exists.
        folder_exists = self._query(
            "SELECT id FROM archive_folders WHERE id = ?;",
            (folder_id,),
        )
        if not folder_exists:
            raise ValueError("Cartella destinazione non trovata.")

        self._execute(
            "UPDATE archive_files SET folder_id = ? WHERE id = ?;",
            (folder_id, file_id),
            "Spostamento file archivio",
        )

    # Archivio link
    def list_archive_links(self, folder_id: int | None = None) -> list[dict]:
        if folder_id is None:
            where_clause = "al.folder_id IS NULL"
            params: tuple = ()
        else:
            where_clause = "al.folder_id = ?"
            params = (int(folder_id),)
        return self._query(
            f"""
            SELECT al.id,
                   al.folder_id,
                   al.name,
                   al.url,
                   t.name AS tag_name,
                   t.color AS tag_color
            FROM archive_links al
            LEFT JOIN tags t ON t.id = al.tag_id
            WHERE {where_clause}
            ORDER BY al.name;
            """,
            params,
        )

    def list_archive_links_filtered(
        self,
        folder_id: int | None,
        *,
        tag_name: str | None = None,
        name_contains: str | None = None,
    ) -> list[dict]:
        where_parts: list[str] = []
        params: list[object] = []
        if folder_id is None:
            where_parts.append("al.folder_id IS NULL")
        else:
            where_parts.append("al.folder_id = ?")
            params.append(int(folder_id))

        clean_tag = (tag_name or "").strip()
        if clean_tag and clean_tag != "Tutti":
            where_parts.append("t.name = ?")
            params.append(clean_tag)

        clean_name = (name_contains or "").strip().lower()
        if clean_name:
            where_parts.append("lower(al.name) LIKE ?")
            params.append(f"%{clean_name}%")

        where_clause = " AND ".join(where_parts) if where_parts else "1=1"
        return self._query(
            f"""
            SELECT al.id,
                   al.folder_id,
                   al.name,
                   al.url,
                   t.name AS tag_name,
                   t.color AS tag_color
            FROM archive_links al
            LEFT JOIN tags t ON t.id = al.tag_id
            WHERE {where_clause}
            ORDER BY al.name;
            """,
            tuple(params),
        )

    def list_archive_links_all(self) -> list[dict]:
        return self._query(
            """
            SELECT al.id,
                   al.folder_id,
                   al.name,
                   al.url,
                   t.name AS tag_name,
                   t.color AS tag_color
            FROM archive_links al
            LEFT JOIN tags t ON t.id = al.tag_id
            ORDER BY al.name;
            """
        )

    def upsert_archive_link(
        self,
        row_id: int | None,
        name: str,
        url: str,
        folder_id: int | None = None,
        tag_name: str | None = None,
    ) -> int:
        clean_name = self._req(name, "Nome link")
        clean_url = self._req(url, "URL")
        tag_id = self._tag_id_from_name(tag_name)
        if row_id:
            self._execute(
                "UPDATE archive_links SET name=?, url=?, folder_id=?, tag_id=? WHERE id=?;",
                (clean_name, clean_url, folder_id, tag_id, int(row_id)),
                "Aggiornamento link",
            )
            return int(row_id)
        return self._insert(
            "INSERT INTO archive_links(name, url, folder_id, tag_id) VALUES (?, ?, ?, ?);",
            (clean_name, clean_url, folder_id, tag_id),
            "Inserimento link",
        )

    def delete_archive_link(self, row_id: int) -> None:
        self._execute(
            "DELETE FROM archive_links WHERE id = ?;",
            (int(row_id),),
            "Eliminazione link",
        )

    # Preferiti archivio
    def list_archive_favorites(self) -> list[dict]:
        return self._query(
            """
            SELECT f.id AS favorite_id,
                   f.item_type,
                   f.item_id,
                   CASE
                       WHEN f.item_type='file' THEN af.name
                       WHEN f.item_type='link' THEN al.name
                       ELSE ''
                   END AS name,
                   CASE
                       WHEN f.item_type='file' THEN af.path
                       WHEN f.item_type='link' THEN al.url
                       ELSE ''
                   END AS location
            FROM archive_favorites f
            LEFT JOIN archive_files af ON af.id = f.item_id AND f.item_type='file'
            LEFT JOIN archive_links al ON al.id = f.item_id AND f.item_type='link'
            ORDER BY f.created_at DESC;
            """
        )

    def add_archive_favorite(self, item_type: str, item_id: int) -> None:
        if item_type not in {"file", "link"}:
            raise ValueError("Tipo preferito non valido.")
        self._execute(
            "INSERT INTO archive_favorites(item_type, item_id) VALUES (?, ?);",
            (item_type, int(item_id)),
            "Inserimento preferito",
        )

    def remove_archive_favorite(self, item_type: str, item_id: int) -> None:
        if item_type not in {"file", "link"}:
            raise ValueError("Tipo preferito non valido.")
        self._execute(
            "DELETE FROM archive_favorites WHERE item_type=? AND item_id=?;",
            (item_type, int(item_id)),
            "Rimozione preferito",
        )

    # Rubrica clienti
    def list_client_contacts(self, client_id: int) -> list[dict]:
        return self._query(
            """
            SELECT id, name, phone, mobile, email, role, note
            FROM client_contacts
            WHERE client_id=?
            ORDER BY name;
            """,
            (int(client_id),),
        )

    def upsert_client_contact(
        self,
        contact_id: int | None,
        client_id: int,
        name: str,
        phone: str = "",
        mobile: str = "",
        email: str = "",
        role: str = "",
        note: str = "",
    ) -> int:
        clean_name = self._req(name, "Nome contatto")
        payload = (
            int(client_id),
            clean_name,
            self._opt(phone),
            self._opt(mobile),
            self._opt(email),
            self._opt(role),
            self._opt(note),
        )
        if contact_id:
            self._execute(
                """
                UPDATE client_contacts
                SET name=?, phone=?, mobile=?, email=?, role=?, note=?
                WHERE id=?;
                """,
                (
                    clean_name,
                    self._opt(phone),
                    self._opt(mobile),
                    self._opt(email),
                    self._opt(role),
                    self._opt(note),
                    int(contact_id),
                ),
                "Aggiornamento contatto",
            )
            return int(contact_id)
        return self._insert(
            """
            INSERT INTO client_contacts(client_id, name, phone, mobile, email, role, note)
            VALUES (?, ?, ?, ?, ?, ?, ?);
            """,
            payload,
            "Inserimento contatto",
        )

    def delete_client_contact(self, contact_id: int) -> None:
        self._execute(
            "DELETE FROM client_contacts WHERE id=?;",
            (int(contact_id),),
            "Eliminazione contatto",
        )

    # Note cliente
    def list_client_notes(self, client_id: int) -> list[dict]:
        return self._query(
            """
            SELECT id, title, content_type, content, created_at
            FROM client_notes
            WHERE client_id=?
            ORDER BY created_at DESC;
            """,
            (int(client_id),),
        )

    def save_client_note(
        self,
        note_id: int | None,
        client_id: int,
        title: str,
        content_type: str,
        content: str,
    ) -> int:
        clean_title = self._req(title, "Titolo nota")
        if content_type not in ("text", "table"):
            raise ValueError("content_type deve essere 'text' o 'table'.")
        if note_id:
            self._execute(
                """
                UPDATE client_notes SET title=?, content_type=?, content=? WHERE id=?;
                """,
                (clean_title, content_type, content, int(note_id)),
                "Aggiornamento nota",
            )
            return int(note_id)
        return self._insert(
            """
            INSERT INTO client_notes(client_id, title, content_type, content)
            VALUES (?, ?, ?, ?);
            """,
            (int(client_id), clean_title, content_type, content),
            "Inserimento nota",
        )

    def delete_client_note(self, note_id: int) -> None:
        self._execute(
            "DELETE FROM client_notes WHERE id=?;",
            (int(note_id),),
            "Eliminazione nota",
        )

    def get_client_note(self, note_id: int) -> dict | None:
        rows = self._query(
            "SELECT id, client_id, title, content_type, content, created_at FROM client_notes WHERE id=?;",
            (int(note_id),),
        )
        return dict(rows[0]) if rows else None

    def get_or_create_tag_for_client(self, client_id: int, client_name: str) -> int:
        """Restituisce l'id del tag con nome=client_name associato al cliente, creandolo se necessario."""
        client_id = int(client_id)
        clean_name = (client_name or "").strip()
        if not clean_name:
            raise ValueError("Nome cliente obbligatorio per il tag.")
        rows = self._query("SELECT id, client_id FROM tags WHERE name=?;", (clean_name,))
        if rows:
            row = rows[0]
            tag_id = int(row["id"])
            if row["client_id"] is None:
                self._execute(
                    "UPDATE tags SET client_id=? WHERE id=?;",
                    (client_id, tag_id),
                    "Associazione tag cliente",
                )
            return tag_id
        return self._insert(
            "INSERT INTO tags(name, color, client_id) VALUES (?, ?, ?);",
            (clean_name, "#0f766e", client_id),
            "Creazione tag cliente",
        )

    # Compatibilita API precedente (used by smoke test / old calls)
    def add_competence(self, name: str) -> int:
        return self.upsert_competence(None, name)

    def add_product_type(
        self,
        name: str,
        flag_ip: bool,
        flag_host: bool,
        flag_preconfigured: bool,
        flag_url: bool,
        flag_port: bool,
    ) -> int:
        return self.upsert_product_type(
            None, name, flag_ip, flag_host, flag_preconfigured, flag_url, flag_port
        )

    def add_release(self, name: str, release_date: str) -> int:
        return self.upsert_release(None, name, release_date)

    def add_environment(self, name: str, release_id: int | None) -> int:
        release_name = ""
        if release_id:
            row = self._query("SELECT name FROM releases WHERE id = ?;", (int(release_id),))
            if row:
                release_name = row[0]["name"]
        return self.upsert_environment(None, name, release_name)

    def add_role(self, name: str, competence: str, multi_clients: bool) -> int:
        next_order = self._next_role_display_order()
        return self.upsert_role(None, name, multi_clients, competence, next_order)

    def _next_role_display_order(self) -> int:
        used_rows = self._query(
            "SELECT display_order FROM roles WHERE display_order BETWEEN 1 AND 20;"
        )
        used = {
            int(row["display_order"])
            for row in used_rows
            if row.get("display_order") is not None
        }
        next_order = next((value for value in range(1, 21) if value not in used), None)
        if next_order is None:
            raise ValueError(
                "Nessun ordine visualizzazione disponibile (1-20) per i ruoli."
            )
        return next_order

    def add_vpn(
        self,
        connection_name: str,
        server_address: str,
        vpn_type: str,
        access_info_type: str,
        username: str,
        password: str,
        vpn_path: str = "",
    ) -> int:
        return self.upsert_vpn(
            None,
            connection_name,
            server_address,
            vpn_type,
            access_info_type,
            username,
            password,
            vpn_path,
        )

    def add_resource(
        self,
        name: str,
        surname: str,
        role_id: int | None,
        phone: str,
        email: str,
        note: str,
        linkedin: str = "",
        photo_link: str = "",
    ) -> int:
        role_name = ""
        if role_id:
            row = self._query("SELECT name FROM roles WHERE id = ?;", (int(role_id),))
            if row:
                role_name = row[0]["name"]
        return self.upsert_resource(
            None,
            name,
            surname,
            role_name,
            "",
            phone,
            email,
            linkedin,
            photo_link,
            note,
        )

    def add_client(
        self,
        name: str,
        location: str,
        vpn_id: int | None,
        resource_ids: Iterable[int] | None,
    ) -> int:
        vpn_name = ""
        if vpn_id:
            row = self._query("SELECT connection_name FROM vpns WHERE id = ?;", (int(vpn_id),))
            if row:
                vpn_name = row[0]["connection_name"]
        resources: list[str] = []
        for resource_id in self._normalize_ids(resource_ids):
            row = self._query(
                "SELECT (name || ' ' || surname) AS label FROM resources WHERE id = ?;",
                (resource_id,),
            )
            if row:
                resources.append(row[0]["label"])
        return self.upsert_client(None, name, location, vpn_name, resources)

    def add_product(
        self,
        name: str,
        client_ids: Iterable[int] | None,
        environment_ids: Iterable[int] | None,
        product_type_id: int,
    ) -> int:
        product_type_rows = self._query(
            "SELECT name FROM product_types WHERE id = ?;",
            (int(product_type_id),),
        )
        if not product_type_rows:
            raise ValueError("Tipo prodotto obbligatorio.")
        product_type = product_type_rows[0]["name"]

        clients: list[str] = []
        for client_id in self._normalize_ids(client_ids):
            row = self._query("SELECT name FROM clients WHERE id = ?;", (client_id,))
            if row:
                clients.append(row[0]["name"])

        environments: list[str] = []
        for environment_id in self._normalize_ids(environment_ids):
            row = self._query("SELECT name FROM environments WHERE id = ?;", (environment_id,))
            if row:
                environments.append(row[0]["name"])

        return self.upsert_product(None, name, product_type, clients, environments)
