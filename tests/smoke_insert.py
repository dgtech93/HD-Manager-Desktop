from pathlib import Path
import sys

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from app.database import DB_PATH, init_db
from app.repository import Repository


def main() -> None:
    if DB_PATH.exists():
        DB_PATH.unlink()

    init_db()
    repository = Repository()

    competence_id = repository.add_competence("Infrastruttura")
    release_id = repository.add_release("RLS-2026.03", "2026-03-13")
    environment_id = repository.add_environment("Produzione", release_id)
    role_id = repository.add_role("Tecnico", "Infrastruttura", True)
    vpn_id = repository.add_vpn(
        connection_name="VPN Cliente Demo",
        server_address="vpn.demo.local",
        vpn_type="Vpn Proprietario",
        access_info_type="Utente/Password",
        username="demo_user",
        password="demo_pwd",
        vpn_path=r"C:\Program Files\DemoVpn\DemoVpn.exe",
    )
    resource_id = repository.add_resource(
        name="Mario",
        surname="Rossi",
        role_id=role_id,
        phone="3331234567",
        email="mario.rossi@example.com",
        note="Referente tecnico",
    )
    client_id = repository.add_client(
        name="Cliente Demo",
        location="Roma",
        vpn_id=vpn_id,
        resource_ids=[resource_id],
    )
    product_type_id = repository.add_product_type(
        name="CRM",
        flag_ip=True,
        flag_host=True,
        flag_preconfigured=False,
        flag_url=True,
        flag_port=True,
    )
    product_id = repository.add_product(
        name="CRM Enterprise",
        client_ids=[client_id],
        environment_ids=[environment_id],
        product_type_id=product_type_id,
    )

    assert competence_id > 0
    assert release_id > 0
    assert environment_id > 0
    assert role_id > 0
    assert vpn_id > 0
    assert resource_id > 0
    assert client_id > 0
    assert product_type_id > 0
    assert product_id > 0

    assert len(repository.list_competences()) == 1
    assert len(repository.list_releases()) == 1
    assert len(repository.list_environments()) == 1
    assert len(repository.list_roles()) == 1
    assert len(repository.list_vpns()) == 1
    assert len(repository.list_resources()) == 1
    assert len(repository.list_clients()) == 1
    assert len(repository.list_product_types()) == 1
    assert len(repository.list_products()) == 1

    print("Smoke test inserimento completato con successo.")


if __name__ == "__main__":
    main()
