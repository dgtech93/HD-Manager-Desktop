from __future__ import annotations

from pathlib import Path

import pytest


@pytest.fixture()
def repo(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    import app.database as db
    from app.database import init_db
    from app.repository import Repository

    data_dir = tmp_path / "data"
    db_path = data_dir / "app.db"
    monkeypatch.setattr(db, "DATA_DIR", data_dir, raising=True)
    monkeypatch.setattr(db, "DB_PATH", db_path, raising=True)

    init_db()
    return Repository(connection_factory=db.get_connection)


def test_archive_files_filtered_and_extensions(repo) -> None:
    # Create folder + files with extensions
    folder_id = repo.add_archive_folder("Docs", None)

    # Insert directly (avoid filesystem dependency from add_archive_file)
    conn = repo.connection_factory()
    try:
        with conn:
            conn.execute(
                """
                INSERT INTO archive_files(folder_id, name, file_type, last_modified, file_size, extension, path, tag_id)
                VALUES (?, ?, 'File', '2026-03-18 10:00', 123, ?, ?, NULL);
                """,
                (folder_id, "alpha.txt", "txt", r"C:\tmp\alpha.txt"),
            )
            conn.execute(
                """
                INSERT INTO archive_files(folder_id, name, file_type, last_modified, file_size, extension, path, tag_id)
                VALUES (?, ?, 'File', '2026-03-18 10:00', 456, ?, ?, NULL);
                """,
                (folder_id, "beta.pdf", "pdf", r"C:\tmp\beta.pdf"),
            )
            conn.execute(
                """
                INSERT INTO archive_files(folder_id, name, file_type, last_modified, file_size, extension, path, tag_id)
                VALUES (?, ?, 'File', '2026-03-18 10:00', 789, ?, ?, NULL);
                """,
                (folder_id, "gamma.PDF", "PDF", r"C:\tmp\gamma.PDF"),
            )
    finally:
        conn.close()

    exts = repo.list_archive_file_extensions(folder_id)
    assert exts == ["pdf", "txt"]

    pdf_rows = repo.list_archive_files_filtered(folder_id, extension="pdf", name_contains=None)
    assert {r["name"] for r in pdf_rows} == {"beta.pdf", "gamma.PDF"}

    name_rows = repo.list_archive_files_filtered(folder_id, extension="Tutte", name_contains="alp")
    assert [r["name"] for r in name_rows] == ["alpha.txt"]


def test_archive_links_filtered(repo) -> None:
    folder_id = repo.add_archive_folder("Links", None)
    tag_id = repo.upsert_tag(None, "DOC", "#0f766e", "")

    conn = repo.connection_factory()
    try:
        with conn:
            conn.execute(
                """
                INSERT INTO archive_links(folder_id, name, url, tag_id)
                VALUES (?, ?, ?, ?);
                """,
                (folder_id, "Wiki", "https://example.local/wiki", tag_id),
            )
            conn.execute(
                """
                INSERT INTO archive_links(folder_id, name, url, tag_id)
                VALUES (?, ?, ?, NULL);
                """,
                (folder_id, "Portal", "https://example.local",),
            )
    finally:
        conn.close()

    all_rows = repo.list_archive_links_filtered(folder_id, tag_name="Tutti", name_contains=None)
    assert {r["name"] for r in all_rows} == {"Wiki", "Portal"}

    doc_rows = repo.list_archive_links_filtered(folder_id, tag_name="DOC", name_contains=None)
    assert [r["name"] for r in doc_rows] == ["Wiki"]

    name_rows = repo.list_archive_links_filtered(folder_id, tag_name="Tutti", name_contains="por")
    assert [r["name"] for r in name_rows] == ["Portal"]

