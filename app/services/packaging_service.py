from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from app.repository import Repository


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _split_csv_items(value: Any) -> list[str]:
    """
    Repository usa spesso group_concat con separatore ', '.
    Gestiamo stringhe vuote/None in modo robusto.
    """
    if value is None:
        return []
    text = str(value).strip()
    if not text:
        return []
    return [part.strip() for part in text.split(",") if part.strip()]


def _resource_label(name: str, surname: str) -> str:
    name = (name or "").strip()
    surname = (surname or "").strip()
    full = f"{name} {surname}".strip()
    return full


@dataclass(frozen=True, slots=True)
class ExportPackage:
    format: str
    version: int
    package_type: str
    generated_at: str
    payload: dict[str, Any]


class PackagingService:
    """
    Export/Import pacchetti (Core, Risorse, VPN).
    In import del Core usiamo una tabella di "pending relazioni"
    per risolvere automaticamente Clienti->Risorse/VPN dopo che
    importano anche i pacchetti dipendenti.
    """

    CORE_TYPE = "core"
    RESOURCES_TYPE = "resources"
    VPNS_TYPE = "vpns"

    def __init__(self, repository: Repository) -> None:
        self.repo = repository

    def _write_package(self, path: str | Path, package_type: str, payload: dict[str, Any]) -> None:
        pkg = ExportPackage(
            format="HDManagerPackage",
            version=1,
            package_type=package_type,
            generated_at=_now_iso(),
            payload=payload,
        )
        out_path = Path(path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(pkg.__dict__, ensure_ascii=False, indent=2), encoding="utf-8")

    def _read_package(self, path: str | Path) -> dict[str, Any]:
        in_path = Path(path)
        raw = in_path.read_text(encoding="utf-8")
        return json.loads(raw)

    def _resolve_pending_client_relations(self) -> int:
        """
        Prova a risolvere le relazioni pendenti.
        Ritorna numero di client risolti (e pending rimossi).
        """
        conn = self.repo.connection_factory()
        resolved_count = 0
        try:
            vpn_map = {row["connection_name"]: int(row["id"]) for row in self.repo.list_vpns()}
            resources = self.repo.list_resources()
            resource_map: dict[str, int] = {}
            for r in resources:
                label = _resource_label(r.get("name"), r.get("surname"))
                if label and label not in resource_map:
                    resource_map[label] = int(r["id"])

            pending_rows = conn.execute(
                "SELECT client_id, vpn_connection_name, resources_json FROM client_pending_relations;"
            ).fetchall()

            with conn:
                for row in pending_rows:
                    client_id = int(row["client_id"])
                    vpn_connection_name = str(row["vpn_connection_name"] or "").strip()
                    resources_json = str(row["resources_json"] or "[]").strip()

                    try:
                        desired_labels = json.loads(resources_json)
                    except json.JSONDecodeError:
                        desired_labels = []

                    desired_labels = [str(x).strip() for x in desired_labels if str(x).strip()]
                    missing_resources = [lab for lab in desired_labels if lab not in resource_map]
                    if missing_resources:
                        continue

                    vpn_id: int | None = None
                    vpn_resolved = False
                    if vpn_connection_name:
                        vpn_id = vpn_map.get(vpn_connection_name)
                        vpn_resolved = vpn_id is not None
                    else:
                        # Se nella configurazione non c'è VPN, è "risolto" già lato VPN.
                        vpn_resolved = True

                    # Aggiorniamo sempre le risorse quando sono disponibili,
                    # anche se la VPN non è ancora stata importata.
                    if vpn_resolved:
                        conn.execute("UPDATE clients SET vpn_id=? WHERE id=?;", (vpn_id, client_id))

                    # Ripopola client_resources solo se nel pacchetto sono presenti risorse.
                    # Se la lista è vuota, consideriamo che l'utente non vuole
                    # sovrascrivere/aggiornare la configurazione locale delle risorse.
                    if desired_labels:
                        conn.execute("DELETE FROM client_resources WHERE client_id=?;", (client_id,))
                        for lab in desired_labels:
                            resource_id = resource_map[lab]
                            conn.execute(
                                "INSERT OR IGNORE INTO client_resources(client_id, resource_id) VALUES (?, ?);",
                                (client_id, resource_id),
                            )

                    if vpn_resolved:
                        conn.execute(
                            "DELETE FROM client_pending_relations WHERE client_id=?;",
                            (client_id,),
                        )
                        resolved_count += 1
        finally:
            conn.close()
        return resolved_count

    # -----------------------
    # EXPORT
    # -----------------------
    def export_core(self, path: str | Path) -> None:
        competences = [{"name": r["name"]} for r in self.repo.list_competences()]
        product_types = [
            {
                "name": r["name"],
                "flag_ip": r.get("flag_ip", 0),
                "flag_host": r.get("flag_host", 0),
                "flag_preconfigured": r.get("flag_preconfigured", 0),
                "flag_url": r.get("flag_url", 0),
                "flag_port": r.get("flag_port", 0),
            }
            for r in self.repo.list_product_types()
        ]
        releases = [{"name": r["name"], "release_date": r.get("release_date") or ""} for r in self.repo.list_releases()]
        environments = [{"name": r["name"], "release_name": r.get("release_name") or ""} for r in self.repo.list_environments()]
        roles = [
            {
                "name": r["name"],
                "competence": r.get("competence") or "",
                "multi_clients": r.get("multi_clients", 0),
                "display_order": r.get("display_order"),
            }
            for r in self.repo.list_roles()
        ]

        # Clienti: esportiamo solo dati base + VPN associata (via connection_name).
        # Le risorse possono variare da utente a utente: non le includiamo nel Core.
        clients_payload = []
        for c in self.repo.list_clients():
            clients_payload.append(
                {
                    "name": c["name"],
                    "location": c.get("location") or "",
                    "vpn_connection_name": (c.get("vpn_name") or "").strip(),
                }
            )

        products = []
        for p in self.repo.list_products():
            clients_list = _split_csv_items(p.get("clients") or "")
            envs_list = _split_csv_items(p.get("environments") or "")
            products.append(
                {
                    "name": p["name"],
                    "product_type": p.get("product_type") or "",
                    "clients": clients_list,
                    "environments": envs_list,
                }
            )

        payload = {
            "competences": competences,
            "product_types": product_types,
            "releases": releases,
            "environments": environments,
            "roles": roles,
            "clients": clients_payload,
            "products": products,
        }

        self._write_package(path, self.CORE_TYPE, payload)

    def export_resources(self, path: str | Path) -> None:
        resources = []
        for r in self.repo.list_resources():
            resources.append(
                {
                    "name": r["name"],
                    "surname": r["surname"],
                    "role_name": r.get("role_name") or "",
                    "phone": r.get("phone") or "",
                    "email": r.get("email") or "",
                    "linkedin": r.get("linkedin") or "",
                    "photo_link": r.get("photo_link") or "",
                    "note": r.get("note") or "",
                }
            )
        payload = {"resources": resources}
        self._write_package(path, self.RESOURCES_TYPE, payload)

    def export_vpns(self, path: str | Path) -> None:
        vpns = []
        for v in self.repo.list_vpns():
            clients = _split_csv_items(v.get("clients") or "")
            vpns.append(
                {
                    "connection_name": v["connection_name"],
                    "server_address": v.get("server_address") or "",
                    "vpn_type": v.get("vpn_type") or "",
                    "access_info_type": v.get("access_info_type") or "",
                    "username": v.get("username") or "",
                    "password": v.get("password") or "",
                    "vpn_path": v.get("vpn_path") or "",
                    # includiamo anche i client perchè l'import imposta vpn_id.
                    "clients": clients,
                }
            )
        payload = {"vpns": vpns}
        self._write_package(path, self.VPNS_TYPE, payload)

    # -----------------------
    # IMPORT
    # -----------------------
    def import_package(self, path: str | Path) -> None:
        data = self._read_package(path)
        pkg_type = str(data.get("package_type") or "").strip()
        if pkg_type == self.CORE_TYPE:
            self.import_core(data)
        elif pkg_type == self.RESOURCES_TYPE:
            self.import_resources(data)
        elif pkg_type == self.VPNS_TYPE:
            self.import_vpns(data)
        else:
            raise ValueError(f"Package non riconosciuto: {pkg_type}")

    def import_core(self, package: dict[str, Any]) -> None:
        payload = package.get("payload") or {}

        # Competences
        comp_by_name = {r["name"]: int(r["id"]) for r in self.repo.list_competences()}
        for row in payload.get("competences", []):
            name = str(row.get("name") or "").strip()
            if not name:
                continue
            row_id = comp_by_name.get(name)
            comp_id = self.repo.upsert_competence(row_id, name)
            comp_by_name[name] = comp_id

        # Product types
        pt_by_name = {r["name"]: int(r["id"]) for r in self.repo.list_product_types()}
        for row in payload.get("product_types", []):
            name = str(row.get("name") or "").strip()
            if not name:
                continue
            row_id = pt_by_name.get(name)
            pt_id = self.repo.upsert_product_type(
                row_id,
                name=name,
                flag_ip=row.get("flag_ip", 0),
                flag_host=row.get("flag_host", 0),
                flag_preconfigured=row.get("flag_preconfigured", 0),
                flag_url=row.get("flag_url", 0),
                flag_port=row.get("flag_port", 0),
            )
            pt_by_name[name] = pt_id

        # Releases
        rel_by_name = {r["name"]: int(r["id"]) for r in self.repo.list_releases()}
        for row in payload.get("releases", []):
            name = str(row.get("name") or "").strip()
            if not name:
                continue
            row_id = rel_by_name.get(name)
            rel_id = self.repo.upsert_release(row_id, name, str(row.get("release_date") or ""))
            rel_by_name[name] = rel_id

        # Environments
        env_by_name = {r["name"]: int(r["id"]) for r in self.repo.list_environments()}
        for row in payload.get("environments", []):
            name = str(row.get("name") or "").strip()
            if not name:
                continue
            row_id = env_by_name.get(name)
            env_id = self.repo.upsert_environment(
                row_id,
                name=name,
                release_name=str(row.get("release_name") or ""),
            )
            env_by_name[name] = env_id

        # Roles
        role_by_name = {r["name"]: int(r["id"]) for r in self.repo.list_roles()}
        for row in payload.get("roles", []):
            name = str(row.get("name") or "").strip()
            if not name:
                continue
            row_id = role_by_name.get(name)
            multi_clients = row.get("multi_clients", 0)
            competence = str(row.get("competence") or "")
            display_order = row.get("display_order")

            try:
                role_id = self.repo.upsert_role(
                    row_id,
                    name=name,
                    multi_clients=multi_clients,
                    competence=competence,
                    display_order=display_order,
                )
            except ValueError:
                # Se l'ordine visualizzazione collide, garantiamo almeno il ruolo inserito.
                role_id = self.repo.upsert_role(
                    row_id,
                    name=name,
                    multi_clients=multi_clients,
                    competence=competence,
                    display_order=None,
                )
            role_by_name[name] = role_id

        # Clients: importiamo solo dati base (nome/localita) e VPN associata.
        # Le risorse restano gestite dall'utente locale.
        existing_clients = self.repo.list_clients()
        clients_by_name = {r["name"]: int(r["id"]) for r in existing_clients}
        existing_by_name = {r["name"]: r for r in existing_clients}
        for row in payload.get("clients", []):
            name = str(row.get("name") or "").strip()
            location = str(row.get("location") or "").strip()
            vpn_connection_name = str(row.get("vpn_connection_name") or "").strip()

            row_id = clients_by_name.get(name)

            # Se il cliente esiste già, preserviamo le sue risorse locali
            # (evitiamo che upsert_client le cancelli).
            resources_to_keep: list[str] | None = None
            if row_id:
                resources_to_keep = _split_csv_items(existing_by_name.get(name, {}).get("resources") or "")

            client_id = self.repo.upsert_client(
                row_id=row_id,
                name=name,
                location=location,
                vpn_name="",  # pending si occupa di vpn.
                resources=resources_to_keep,
            )
            clients_by_name[name] = client_id

            if vpn_connection_name:
                conn = self.repo.connection_factory()
                try:
                    with conn:
                        conn.execute(
                            """
                            INSERT OR REPLACE INTO client_pending_relations(client_id, vpn_connection_name, resources_json)
                            VALUES (?, ?, ?);
                            """,
                            (
                                client_id,
                                vpn_connection_name,
                                json.dumps([], ensure_ascii=False),
                            ),
                        )
                finally:
                    conn.close()

        # Products
        products_by_name = {r["name"]: int(r["id"]) for r in self.repo.list_products()}
        for row in payload.get("products", []):
            name = str(row.get("name") or "").strip()
            if not name:
                continue
            row_id = products_by_name.get(name)
            clients_list = row.get("clients") or []
            envs_list = row.get("environments") or []
            if isinstance(clients_list, str):
                clients_list = _split_csv_items(clients_list)
            if isinstance(envs_list, str):
                envs_list = _split_csv_items(envs_list)
            clients_list = [str(x).strip() for x in clients_list if str(x).strip()]
            envs_list = [str(x).strip() for x in envs_list if str(x).strip()]

            self.repo.upsert_product(
                row_id=row_id,
                name=name,
                product_type=str(row.get("product_type") or "").strip(),
                clients=clients_list,
                environments=envs_list,
            )

        # Prova a risolvere subito se nel DB sono già presenti VPN.
        self._resolve_pending_client_relations()

    def import_resources(self, package: dict[str, Any]) -> None:
        payload = package.get("payload") or {}
        resources = payload.get("resources") or []

        resources_by_key: dict[str, int] = {}
        for r in self.repo.list_resources():
            key = f"{(r.get('name') or '').strip()}|{(r.get('surname') or '').strip()}"
            resources_by_key[key] = int(r["id"])

        for r in resources:
            name = str(r.get("name") or "").strip()
            surname = str(r.get("surname") or "").strip()
            if not name or not surname:
                continue
            role_name = str(r.get("role_name") or "").strip()
            row_id = resources_by_key.get(f"{name}|{surname}")
            resource_id = self.repo.upsert_resource(
                row_id=row_id,
                name=name,
                surname=surname,
                role_name=role_name,
                phone=str(r.get("phone") or ""),
                email=str(r.get("email") or ""),
                linkedin=str(r.get("linkedin") or ""),
                photo_link=str(r.get("photo_link") or ""),
                note=str(r.get("note") or ""),
            )
            resources_by_key[f"{name}|{surname}"] = resource_id

        self._resolve_pending_client_relations()

    def import_vpns(self, package: dict[str, Any]) -> None:
        payload = package.get("payload") or {}
        vpns = payload.get("vpns") or []

        vpns_by_name = {r["connection_name"]: int(r["id"]) for r in self.repo.list_vpns()}

        for v in vpns:
            connection_name = str(v.get("connection_name") or "").strip()
            if not connection_name:
                continue
            row_id = vpns_by_name.get(connection_name)
            clients_list = v.get("clients") or []
            if isinstance(clients_list, str):
                clients_list = _split_csv_items(clients_list)
            clients_list = [str(x).strip() for x in clients_list if str(x).strip()]

            self.repo.upsert_vpn(
                row_id=row_id,
                connection_name=connection_name,
                server_address=str(v.get("server_address") or ""),
                vpn_type=str(v.get("vpn_type") or ""),
                access_info_type=str(v.get("access_info_type") or ""),
                username=str(v.get("username") or ""),
                password=str(v.get("password") or ""),
                vpn_path=str(v.get("vpn_path") or ""),
                clients=clients_list,
            )

        self._resolve_pending_client_relations()

