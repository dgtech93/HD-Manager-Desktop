from __future__ import annotations

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
            data.get("phone") or "",
            data.get("email") or "",
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

