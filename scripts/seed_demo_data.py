from __future__ import annotations

import datetime as dt

from app.database import init_db
from app.repository import Repository


def _map_by(rows: list[dict], key: str) -> dict[str, dict]:
    out: dict[str, dict] = {}
    for row in rows:
        value = str(row.get(key, "")).strip()
        if value:
            out[value.lower()] = row
    return out


def _full_name(name: str, surname: str) -> str:
    return f"{name.strip()} {surname.strip()}".strip()


def _find_free_role_order(repo: Repository, preferred: int, current_id: int | None = None) -> int:
    roles = repo.list_roles()
    used: set[int] = set()
    for row in roles:
        row_id = int(row["id"])
        if current_id is not None and row_id == current_id:
            continue
        raw = row.get("display_order")
        try:
            value = int(raw) if raw is not None else 0
        except (TypeError, ValueError):
            value = 0
        if 1 <= value <= 20:
            used.add(value)

    if 1 <= preferred <= 20 and preferred not in used:
        return preferred
    for value in range(1, 21):
        if value not in used:
            return value
    raise ValueError("Nessun ordine visualizzazione libero (1-20) per i ruoli demo.")


def _ensure_credential(
    repo: Repository,
    client_name: str,
    product_name: str,
    environment_versions: list[tuple[str, str]],
    domain: str,
    login_name: str,
    password: str,
    ip: str = "",
    url: str = "",
    host: str = "",
    rdp_path: str = "",
    port: str | int | None = None,
    password_expiry: bool = False,
    password_duration_days: int | None = None,
    note: str = "",
) -> str:
    clients = _map_by(repo.list_clients(), "name")
    products = _map_by(repo.list_products(), "name")
    client_row = clients.get(client_name.lower())
    product_row = products.get(product_name.lower())
    if client_row is None or product_row is None:
        raise ValueError(f"Associazione credenziale non trovata: {client_name} / {product_name}")

    client_id = int(client_row["id"])
    product_id = int(product_row["id"])
    username = f"{domain}\\{login_name}" if domain and login_name else login_name
    payload = {
        "credential_name": product_name,
        "environment_versions": environment_versions,
        "domain": domain,
        "login_name": login_name,
        "username": username,
        "password": password,
        "ip": ip,
        "url": url,
        "host": host,
        "rdp_path": rdp_path,
        "port": port,
        "password_expiry": password_expiry,
        "password_inserted_at": dt.date.today().isoformat() if password_expiry else None,
        "password_duration_days": password_duration_days if password_expiry else None,
        "password_end_date": None,
        "note": note,
    }

    existing = repo.list_product_credentials(client_id, product_id)
    target = next(
        (
            row
            for row in existing
            if str(row.get("credential_name", "")).strip().lower() == product_name.lower()
        ),
        None,
    )
    if target is not None:
        repo.update_product_credential(credential_id=int(target["id"]), **payload)
        return "updated"

    repo.create_product_credential(client_id=client_id, product_id=product_id, **payload)
    return "created"


def seed_demo_data() -> None:
    init_db()
    repo = Repository()

    # Competenze
    competence_names = [
        "DEMO - Gestione Progetto",
        "DEMO - Sistemistica",
        "DEMO - Applicativo",
    ]
    competences = _map_by(repo.list_competences(), "name")
    for name in competence_names:
        row = competences.get(name.lower())
        repo.upsert_competence(int(row["id"]) if row else None, name)
    competences = _map_by(repo.list_competences(), "name")

    # Tipi prodotto
    product_types = [
        ("DEMO Web App", 0, 0, 0, 1, 0),
        ("DEMO Server RDP", 1, 1, 1, 0, 1),
        ("DEMO API", 1, 1, 0, 1, 1),
    ]
    existing_product_types = _map_by(repo.list_product_types(), "name")
    for name, flag_ip, flag_host, flag_preconfigured, flag_url, flag_port in product_types:
        row = existing_product_types.get(name.lower())
        repo.upsert_product_type(
            int(row["id"]) if row else None,
            name,
            flag_ip,
            flag_host,
            flag_preconfigured,
            flag_url,
            flag_port,
        )
    existing_product_types = _map_by(repo.list_product_types(), "name")

    # Release
    releases = [
        ("DEMO 2026.1", "2026-01-15"),
        ("DEMO 2026.2", "2026-05-20"),
        ("DEMO 2026.3", "2026-09-10"),
    ]
    existing_releases = _map_by(repo.list_releases(), "name")
    for name, release_date in releases:
        row = existing_releases.get(name.lower())
        repo.upsert_release(int(row["id"]) if row else None, name, release_date)
    existing_releases = _map_by(repo.list_releases(), "name")

    # Ambienti
    environments = [
        ("DEMO Produzione", "DEMO 2026.2"),
        ("DEMO Pre-Produzione", "DEMO 2026.2"),
        ("DEMO Collaudo", "DEMO 2026.3"),
    ]
    existing_environments = _map_by(repo.list_environments(), "name")
    for name, release_name in environments:
        row = existing_environments.get(name.lower())
        repo.upsert_environment(int(row["id"]) if row else None, name, release_name)
    existing_environments = _map_by(repo.list_environments(), "name")

    # Ruoli
    roles = [
        ("DEMO PM", 0, "DEMO - Gestione Progetto", 1),
        ("DEMO Consulente", 1, "DEMO - Applicativo", 2),
        ("DEMO SysAdmin", 1, "DEMO - Sistemistica", 3),
        ("DEMO Helpdesk", 1, "DEMO - Applicativo", 4),
    ]
    existing_roles = _map_by(repo.list_roles(), "name")
    for name, multi_clients, competence, preferred_order in roles:
        row = existing_roles.get(name.lower())
        row_id = int(row["id"]) if row else None
        order = _find_free_role_order(repo, preferred_order, current_id=row_id)
        repo.upsert_role(row_id, name, multi_clients, competence, order)
    existing_roles = _map_by(repo.list_roles(), "name")

    # VPN
    windows_vpns = repo.list_windows_vpn_connections()
    vpn_alfa_type = "VPN Windows" if windows_vpns else "Vpn Proprietario"
    vpn_alfa_name = windows_vpns[0] if windows_vpns else "DEMO VPN Alfa"
    vpns = [
        (
            vpn_alfa_name,
            "vpn-alfa.demo.local",
            vpn_alfa_type,
            "Credenziali AD",
            "vpn_alfa",
            "DemoAlfa#2026",
            r"C:\Program Files\DemoVpn\DemoVpn.exe" if vpn_alfa_type != "VPN Windows" else "",
            [],
        ),
        (
            "DEMO VPN Beta",
            "vpn-beta.demo.local",
            "Vpn Proprietario",
            "Token + password",
            "vpn_beta",
            "DemoBeta#2026",
            r"C:\Program Files\DemoVpn\DemoVpn.exe",
            [],
        ),
    ]
    existing_vpns = _map_by(repo.list_vpns(), "connection_name")
    for connection_name, server, vpn_type, access_info, username, password, vpn_path, clients in vpns:
        row = existing_vpns.get(connection_name.lower())
        repo.upsert_vpn(
            int(row["id"]) if row else None,
            connection_name,
            server,
            vpn_type,
            access_info,
            username,
            password,
            vpn_path,
            clients,
        )
    existing_vpns = _map_by(repo.list_vpns(), "connection_name")

    # Risorse
    resources = [
        ("Luca", "Bianchi", "DEMO PM", "+39 333 1000001", "luca.bianchi@demo.it", "Referente principale"),
        (
            "Sara",
            "Verdi",
            "DEMO Consulente",
            "+39 333 1000002",
            "sara.verdi@demo.it",
            "Analisi processi",
        ),
        (
            "Marco",
            "Neri",
            "DEMO SysAdmin",
            "+39 333 1000003",
            "marco.neri@demo.it",
            "Gestione infrastruttura",
        ),
        (
            "Elisa",
            "Gallo",
            "DEMO Helpdesk",
            "+39 333 1000004",
            "elisa.gallo@demo.it",
            "Supporto utenti",
        ),
        (
            "Paolo",
            "Riva",
            "DEMO Consulente",
            "+39 333 1000005",
            "paolo.riva@demo.it",
            "Consulenza applicativa",
        ),
    ]
    existing_resources = {
        _full_name(row.get("name", ""), row.get("surname", "")).lower(): row
        for row in repo.list_resources()
    }
    for name, surname, role_name, phone, email, note in resources:
        key = _full_name(name, surname).lower()
        row = existing_resources.get(key)
        repo.upsert_resource(
            int(row["id"]) if row else None,
            name,
            surname,
            role_name,
            phone,
            email,
            note,
        )
    existing_resources = {
        _full_name(row.get("name", ""), row.get("surname", "")).lower(): row
        for row in repo.list_resources()
    }

    # Clienti
    clients = [
        (
            "DEMO Cliente Alfa",
            "Milano",
            "DEMO VPN Alfa",
            [
                _full_name("Luca", "Bianchi"),
                _full_name("Sara", "Verdi"),
                _full_name("Marco", "Neri"),
            ],
        ),
        (
            "DEMO Cliente Beta",
            "Roma",
            "DEMO VPN Beta",
            [
                _full_name("Sara", "Verdi"),
                _full_name("Elisa", "Gallo"),
            ],
        ),
        (
            "DEMO Cliente Gamma",
            "Torino",
            "",
            [
                _full_name("Paolo", "Riva"),
                _full_name("Marco", "Neri"),
            ],
        ),
    ]
    existing_clients = _map_by(repo.list_clients(), "name")
    for name, location, vpn_name, resource_names in clients:
        row = existing_clients.get(name.lower())
        repo.upsert_client(
            int(row["id"]) if row else None,
            name,
            location,
            vpn_name,
            ", ".join(resource_names),
        )
    existing_clients = _map_by(repo.list_clients(), "name")

    # Aggiorna le VPN con i clienti demo associati.
    vpn_assignments = {
        vpn_alfa_name: ["DEMO Cliente Alfa"],
        "DEMO VPN Beta": ["DEMO Cliente Beta"],
    }
    for vpn_name, client_names in vpn_assignments.items():
        row = existing_vpns.get(vpn_name.lower())
        if row is None:
            continue
        repo.upsert_vpn(
            int(row["id"]),
            row["connection_name"],
            row["server_address"],
            row["vpn_type"],
            row.get("access_info_type") or "",
            row["username"],
            row["password"],
            ", ".join(client_names),
        )

    # Prodotti
    products = [
        (
            "DEMO Portale HR",
            "DEMO Web App",
            ["DEMO Cliente Alfa", "DEMO Cliente Beta"],
            ["DEMO Produzione", "DEMO Collaudo"],
        ),
        (
            "DEMO Gestionale ERP",
            "DEMO Server RDP",
            ["DEMO Cliente Alfa", "DEMO Cliente Gamma"],
            ["DEMO Produzione", "DEMO Pre-Produzione"],
        ),
        (
            "DEMO CRM Service",
            "DEMO API",
            ["DEMO Cliente Beta", "DEMO Cliente Gamma"],
            ["DEMO Produzione", "DEMO Pre-Produzione", "DEMO Collaudo"],
        ),
    ]
    existing_products = _map_by(repo.list_products(), "name")
    for name, product_type, client_names, environment_names in products:
        row = existing_products.get(name.lower())
        repo.upsert_product(
            int(row["id"]) if row else None,
            name,
            product_type,
            ", ".join(client_names),
            ", ".join(environment_names),
        )
    existing_products = _map_by(repo.list_products(), "name")

    # Credenziali prodotto per cliente
    credential_results: list[str] = []
    credential_results.append(
        _ensure_credential(
            repo,
            "DEMO Cliente Alfa",
            "DEMO Portale HR",
            [
                ("DEMO Produzione", "DEMO 2026.2"),
                ("DEMO Collaudo", "DEMO 2026.3"),
            ],
            domain="ALFA",
            login_name="portale.hr",
            password="AlfaHr#2026",
            url="https://hr.alfa.demo.local",
            note="Accesso principale HR",
        )
    )
    credential_results.append(
        _ensure_credential(
            repo,
            "DEMO Cliente Beta",
            "DEMO Portale HR",
            [
                ("DEMO Produzione", "DEMO 2026.2"),
                ("DEMO Collaudo", "DEMO 2026.3"),
            ],
            domain="BETA",
            login_name="portale.hr",
            password="BetaHr#2026",
            url="https://hr.beta.demo.local",
            password_expiry=True,
            password_duration_days=90,
            note="Utente con scadenza password",
        )
    )
    credential_results.append(
        _ensure_credential(
            repo,
            "DEMO Cliente Alfa",
            "DEMO Gestionale ERP",
            [
                ("DEMO Produzione", "DEMO 2026.2"),
                ("DEMO Pre-Produzione", "DEMO 2026.2"),
            ],
            domain="ALFA",
            login_name="erp.admin",
            password="AlfaErp#2026",
            ip="10.10.1.20",
            host="erp-prod-alfa",
            rdp_path=r"C:\RDP\DEMO_ERP_ALFA.rdp",
            port=3389,
            note="Server ERP primario",
        )
    )
    credential_results.append(
        _ensure_credential(
            repo,
            "DEMO Cliente Gamma",
            "DEMO Gestionale ERP",
            [
                ("DEMO Produzione", "DEMO 2026.2"),
                ("DEMO Pre-Produzione", "DEMO 2026.2"),
            ],
            domain="GAMMA",
            login_name="erp.admin",
            password="GammaErp#2026",
            ip="10.20.1.30",
            host="erp-prod-gamma",
            rdp_path=r"C:\RDP\DEMO_ERP_GAMMA.rdp",
            port=3390,
            note="Server ERP cliente gamma",
        )
    )
    credential_results.append(
        _ensure_credential(
            repo,
            "DEMO Cliente Beta",
            "DEMO CRM Service",
            [
                ("DEMO Produzione", "DEMO 2026.2"),
                ("DEMO Pre-Produzione", "DEMO 2026.2"),
                ("DEMO Collaudo", "DEMO 2026.3"),
            ],
            domain="BETA",
            login_name="crm.service",
            password="BetaCrm#2026",
            ip="172.16.1.40",
            host="crm-beta",
            url="https://api.crm.beta.demo.local",
            port=8443,
            note="API CRM beta",
        )
    )
    credential_results.append(
        _ensure_credential(
            repo,
            "DEMO Cliente Gamma",
            "DEMO CRM Service",
            [
                ("DEMO Produzione", "DEMO 2026.2"),
                ("DEMO Pre-Produzione", "DEMO 2026.2"),
                ("DEMO Collaudo", "DEMO 2026.3"),
            ],
            domain="GAMMA",
            login_name="crm.service",
            password="GammaCrm#2026",
            ip="172.16.2.40",
            host="crm-gamma",
            url="https://api.crm.gamma.demo.local",
            port=9443,
            password_expiry=True,
            password_duration_days=120,
            note="API CRM gamma",
        )
    )

    demo_counts = {
        "competenze": len([x for x in repo.list_competences() if str(x.get("name", "")).startswith("DEMO ")]),
        "tipi_prodotto": len(
            [x for x in repo.list_product_types() if str(x.get("name", "")).startswith("DEMO ")]
        ),
        "release": len([x for x in repo.list_releases() if str(x.get("name", "")).startswith("DEMO ")]),
        "ambienti": len([x for x in repo.list_environments() if str(x.get("name", "")).startswith("DEMO ")]),
        "ruoli": len([x for x in repo.list_roles() if str(x.get("name", "")).startswith("DEMO ")]),
        "vpn": len([x for x in repo.list_vpns() if str(x.get("connection_name", "")).startswith("DEMO ")]),
        "risorse": len(
            [
                x
                for x in repo.list_resources()
                if _full_name(x.get("name", ""), x.get("surname", "")).startswith("Luca ")
                or _full_name(x.get("name", ""), x.get("surname", "")).startswith("Sara ")
                or _full_name(x.get("name", ""), x.get("surname", "")).startswith("Marco ")
                or _full_name(x.get("name", ""), x.get("surname", "")).startswith("Elisa ")
                or _full_name(x.get("name", ""), x.get("surname", "")).startswith("Paolo ")
            ]
        ),
        "clienti": len([x for x in repo.list_clients() if str(x.get("name", "")).startswith("DEMO Cliente ")]),
        "prodotti": len([x for x in repo.list_products() if str(x.get("name", "")).startswith("DEMO ")]),
    }

    created = sum(1 for result in credential_results if result == "created")
    updated = sum(1 for result in credential_results if result == "updated")

    print("Seed dati demo completato.")
    print("Riepilogo:")
    for key, value in demo_counts.items():
        print(f"- {key}: {value}")
    print(f"- credenziali create: {created}")
    print(f"- credenziali aggiornate: {updated}")


if __name__ == "__main__":
    seed_demo_data()
