from __future__ import annotations

from app.repository import Repository


class RepositoryService(Repository):
    """Service wrapper around data-access repository.

    For now it subclasses the existing `Repository` to keep behavior unchanged,
    while giving the app a clearer Services layer in MVC.
    """

    pass

