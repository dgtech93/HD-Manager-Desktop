from __future__ import annotations

from dataclasses import dataclass

from app.services.repository_service import RepositoryService


@dataclass(frozen=True, slots=True)
class ClientSelection:
    client_id: int | None
    product_id: int | None


class ClientsService:
    def __init__(self, repo: RepositoryService) -> None:
        self.repo = repo

    def list_clients(self) -> list[dict]:
        return self.repo.list_clients()

    def list_products(self) -> list[dict]:
        return self.repo.list_products()

    def list_resources(self) -> list[dict]:
        return self.repo.list_resources()

    def list_roles(self) -> list[dict]:
        return self.repo.list_roles()

    def list_vpns(self) -> list[dict]:
        return self.repo.list_vpns()

    def list_tags(self) -> list[dict]:
        return self.repo.list_tags()

    def list_client_contacts(self, client_id: int) -> list[dict]:
        return self.repo.list_client_contacts(client_id)

    def upsert_client_contact(self, contact_id: int | None, client_id: int, **payload) -> int:
        return self.repo.upsert_client_contact(contact_id, client_id, **payload)

    def delete_client_contact(self, contact_id: int) -> None:
        self.repo.delete_client_contact(contact_id)

    def list_tags_for_client(self, client_id: int) -> list[dict]:
        return self.repo.list_tags_for_client(client_id)

    # lookups
    def list_roles_lookup(self) -> list[dict]:
        return self.repo.list_roles_lookup()

    # client/product aggregations
    def list_client_product_environment_releases(self, client_id: int) -> list[dict]:
        return self.repo.list_client_product_environment_releases(client_id)

