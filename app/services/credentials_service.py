from __future__ import annotations

from app.services.repository_service import RepositoryService


class CredentialsService:
    def __init__(self, repo: RepositoryService) -> None:
        self.repo = repo

    def list_product_credentials(self, client_id: int, product_id: int) -> list[dict]:
        return self.repo.list_product_credentials(client_id, product_id)

    def get_credential_detail(self, credential_id: int) -> dict:
        return self.repo.get_product_credential_detail(credential_id)

    def get_product_type_flags_for_product(self, product_id: int) -> dict:
        return self.repo.get_product_type_flags_for_product(product_id)

    def create_credential(self, client_id: int, product_id: int, **payload) -> int:
        return self.repo.create_product_credential(client_id=client_id, product_id=product_id, **payload)

    def update_credential(self, credential_id: int, **payload) -> None:
        self.repo.update_product_credential(credential_id=credential_id, **payload)

    def delete_credential(self, credential_id: int) -> None:
        self.repo.delete_product_credential(credential_id)

    # legacy-compatible aliases (used by existing dialogs/mixins during migration)
    def create_product_credential(self, client_id: int, product_id: int, **payload) -> int:
        return self.create_credential(client_id, product_id, **payload)

    def update_product_credential(self, credential_id: int, **payload) -> None:
        self.update_credential(credential_id, **payload)

