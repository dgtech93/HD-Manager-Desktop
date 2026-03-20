from __future__ import annotations

from dataclasses import dataclass

from app.services.archive_service import ArchiveService
from app.services.clients_service import ClientsService
from app.services.credentials_service import CredentialsService
from app.services.settings_service import SettingsService
from app.services.system_service import SystemService


@dataclass(slots=True)
class MainCache:
    clients: list[dict] = None  # type: ignore[assignment]
    products: list[dict] = None  # type: ignore[assignment]
    resources: list[dict] = None  # type: ignore[assignment]
    vpns: list[dict] = None  # type: ignore[assignment]
    roles: list[dict] = None  # type: ignore[assignment]
    tags: list[dict] = None  # type: ignore[assignment]
    archive_folders: list[dict] = None  # type: ignore[assignment]
    archive_links: list[dict] = None  # type: ignore[assignment]

    def __post_init__(self) -> None:
        # Use lists rather than None for simpler view code.
        self.clients = self.clients or []
        self.products = self.products or []
        self.resources = self.resources or []
        self.vpns = self.vpns or []
        self.roles = self.roles or []
        self.tags = self.tags or []
        self.archive_folders = self.archive_folders or []
        self.archive_links = self.archive_links or []


class MainController:
    """Use-cases for the main window (Clients + Archive)."""

    def __init__(
        self,
        *,
        clients: ClientsService,
        archive: ArchiveService,
        credentials: CredentialsService,
        system: SystemService,
    ) -> None:
        self.clients = clients
        self.archive = archive
        self.credentials = credentials
        self.system = system
        # Legacy compatibility: some existing views still call `repository.*`.
        self.repository = clients.repo
        self.settings = SettingsService(self.repository)

        self.cache = MainCache()

    def refresh_cache(self) -> None:
        self.cache.clients = self.clients.list_clients()
        self.cache.products = self.clients.list_products()
        self.cache.resources = self.clients.list_resources()
        self.cache.vpns = self.clients.list_vpns()
        self.cache.roles = self.clients.list_roles()
        self.cache.tags = self.clients.list_tags()
        self.cache.archive_folders = self.archive.list_folders()
        self.cache.archive_links = self.archive.list_links_all()

