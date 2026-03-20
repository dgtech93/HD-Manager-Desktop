from __future__ import annotations

from app import database


class DbService:
    """Database lifecycle and access.

    Keeps `main.py` free of direct DB module calls.
    """

    def init_db(self) -> None:
        database.init_db()

    def get_connection(self):
        return database.get_connection()

