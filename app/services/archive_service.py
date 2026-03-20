from __future__ import annotations

from app.services.repository_service import RepositoryService


class ArchiveService:
    def __init__(self, repo: RepositoryService) -> None:
        self.repo = repo

    # folders
    def list_folders(self) -> list[dict]:
        return self.repo.list_archive_folders()

    def add_folder(self, name: str, parent_id: int | None) -> int:
        return self.repo.add_archive_folder(name, parent_id)

    def delete_folder(self, folder_id: int) -> None:
        self.repo.delete_archive_folder(folder_id)

    # files
    def list_files(self, folder_id: int | None) -> list[dict]:
        return self.repo.list_archive_files(folder_id)

    def list_file_extensions(self, folder_id: int | None) -> list[str]:
        return self.repo.list_archive_file_extensions(folder_id)

    def list_files_filtered(
        self,
        folder_id: int | None,
        *,
        extension: str | None = None,
        name_contains: str | None = None,
    ) -> list[dict]:
        return self.repo.list_archive_files_filtered(
            folder_id, extension=extension, name_contains=name_contains
        )

    def list_files_all(self) -> list[dict]:
        return self.repo.list_archive_files_all()

    def add_file(self, folder_id: int | None, file_path: str, tag_name: str | None) -> int:
        return self.repo.add_archive_file(folder_id, file_path, tag_name)

    def update_file_tag(self, file_id: int, tag_name: str | None) -> None:
        self.repo.update_archive_file_tag(file_id, tag_name)

    def delete_file(self, file_id: int) -> None:
        self.repo.delete_archive_file(file_id)

    def move_file(self, file_id: int, folder_id: int) -> None:
        self.repo.move_archive_file(file_id, folder_id)

    # links
    def list_links(self, folder_id: int | None) -> list[dict]:
        return self.repo.list_archive_links(folder_id)

    def list_links_filtered(
        self,
        folder_id: int | None,
        *,
        tag_name: str | None = None,
        name_contains: str | None = None,
    ) -> list[dict]:
        return self.repo.list_archive_links_filtered(
            folder_id, tag_name=tag_name, name_contains=name_contains
        )

    def list_links_all(self) -> list[dict]:
        return self.repo.list_archive_links_all()

    def upsert_link(
        self,
        row_id: int | None,
        name: str,
        url: str,
        folder_id: int | None,
        tag_name: str | None,
    ) -> int:
        return self.repo.upsert_archive_link(row_id, name, url, folder_id, tag_name)

    def delete_link(self, row_id: int) -> None:
        self.repo.delete_archive_link(row_id)

    # favorites
    def list_favorites(self) -> list[dict]:
        return self.repo.list_archive_favorites()

    def add_favorite(self, item_type: str, item_id: int) -> None:
        self.repo.add_archive_favorite(item_type, item_id)

    def remove_favorite(self, item_type: str, item_id: int) -> None:
        self.repo.remove_archive_favorite(item_type, item_id)

