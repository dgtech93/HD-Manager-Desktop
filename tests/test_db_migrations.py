from __future__ import annotations

import sqlite3
from pathlib import Path

import pytest


@pytest.fixture()
def temp_db(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> Path:
    # Patch app.database globals so init_db writes into tmp_path
    import app.database as db

    data_dir = tmp_path / "data"
    db_path = data_dir / "app.db"
    monkeypatch.setattr(db, "DATA_DIR", data_dir, raising=True)
    monkeypatch.setattr(db, "DB_PATH", db_path, raising=True)
    return db_path


def _table_columns(conn: sqlite3.Connection, table: str) -> set[str]:
    rows = conn.execute(f"PRAGMA table_info({table});").fetchall()
    return {row[1] for row in rows}  # name


def _indexes(conn: sqlite3.Connection) -> set[str]:
    rows = conn.execute("SELECT name FROM sqlite_master WHERE type='index';").fetchall()
    return {row[0] for row in rows}


def test_init_db_adds_password_ref_columns_and_indexes(temp_db: Path) -> None:
    from app.database import init_db, get_connection

    init_db()
    conn = get_connection()
    try:
        vpn_cols = _table_columns(conn, "vpns")
        assert "password_ref" in vpn_cols

        cred_cols = _table_columns(conn, "product_credentials")
        assert "password_ref" in cred_cols

        idx = _indexes(conn)
        assert "idx_archive_files_folder_id" in idx
        assert "idx_archive_files_tag_id" in idx
        assert "idx_archive_files_extension" in idx
        assert "idx_archive_links_folder_id" in idx
        assert "idx_archive_links_tag_id" in idx
        assert "idx_product_credentials_client_product" in idx
    finally:
        conn.close()

