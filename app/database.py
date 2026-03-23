import os
import sys
from pathlib import Path
import shutil
import sqlite3


ROOT_DIR = Path(__file__).resolve().parent.parent


def _resolve_data_dir() -> Path:
    # In esecuzione installata/frozen, Program Files non è scrivibile:
    # usiamo LocalAppData per DB e dati runtime.
    if getattr(sys, "frozen", False):
        local_appdata = os.getenv("LOCALAPPDATA")
        if local_appdata:
            new_dir = Path(local_appdata) / "HDManagerDesktop" / "data"

            # Migrazione "robustezza": se esiste un vecchio DB dentro la cartella applicativa
            # (es. Program Files / _internal/data), e il nuovo DB non esiste ancora,
            # copiamo il file una volta sola.
            try:
                new_db = new_dir / "app.db"
                old_db = (ROOT_DIR / "data" / "app.db")
                if (not new_db.exists()) and old_db.exists():
                    new_dir.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(old_db, new_db)
            except Exception:
                # Se la copia fallisce, procediamo comunque: verrà creato un DB nuovo.
                pass

            return new_dir

    # Fallback sviluppo locale
    return ROOT_DIR / "data"


DATA_DIR = _resolve_data_dir()
DB_PATH = DATA_DIR / "app.db"


SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS product_types (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    flag_ip INTEGER NOT NULL DEFAULT 0 CHECK (flag_ip IN (0, 1)),
    flag_host INTEGER NOT NULL DEFAULT 0 CHECK (flag_host IN (0, 1)),
    flag_preconfigured INTEGER NOT NULL DEFAULT 0 CHECK (flag_preconfigured IN (0, 1)),
    flag_url INTEGER NOT NULL DEFAULT 0 CHECK (flag_url IN (0, 1)),
    flag_port INTEGER NOT NULL DEFAULT 0 CHECK (flag_port IN (0, 1)),
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS competences (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS releases (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    release_date TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS environments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    release_id INTEGER REFERENCES releases(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS roles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    competence TEXT NOT NULL,
    multi_clients INTEGER NOT NULL DEFAULT 0 CHECK (multi_clients IN (0, 1)),
    display_order INTEGER CHECK (display_order BETWEEN 1 AND 20)
);

CREATE TABLE IF NOT EXISTS vpns (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    connection_name TEXT NOT NULL UNIQUE,
    server_address TEXT NOT NULL,
    vpn_type TEXT NOT NULL,
    access_info_type TEXT NOT NULL,
    username TEXT NOT NULL,
    password TEXT NOT NULL,
    password_ref TEXT,
    vpn_path TEXT
);

CREATE TABLE IF NOT EXISTS resources (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    surname TEXT NOT NULL,
    role_id INTEGER REFERENCES roles(id) ON DELETE SET NULL,
    phone TEXT,
    email TEXT,
    note TEXT
);

CREATE TABLE IF NOT EXISTS clients (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    location TEXT NOT NULL,
    link TEXT,
    vpn_id INTEGER REFERENCES vpns(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS client_resources (
    client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    resource_id INTEGER NOT NULL REFERENCES resources(id) ON DELETE CASCADE,
    PRIMARY KEY (client_id, resource_id)
);

CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    product_type_id INTEGER NOT NULL REFERENCES product_types(id) ON DELETE RESTRICT
);

CREATE TABLE IF NOT EXISTS product_clients (
    product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
    client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    PRIMARY KEY (product_id, client_id)
);

CREATE TABLE IF NOT EXISTS product_environments (
    product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
    environment_id INTEGER NOT NULL REFERENCES environments(id) ON DELETE CASCADE,
    PRIMARY KEY (product_id, environment_id)
);

CREATE TABLE IF NOT EXISTS product_credentials (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
    credential_name TEXT NOT NULL,
    ip TEXT,
    url TEXT,
    host TEXT,
    rdp_path TEXT,
    domain TEXT,
    login_name TEXT,
    username TEXT,
    password TEXT,
    password_ref TEXT,
    port INTEGER,
    password_expiry INTEGER NOT NULL DEFAULT 0 CHECK (password_expiry IN (0, 1)),
    password_inserted_at TEXT,
    password_duration_days INTEGER,
    password_end_date TEXT,
    note TEXT,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS product_credential_environments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    credential_id INTEGER NOT NULL REFERENCES product_credentials(id) ON DELETE CASCADE,
    environment_id INTEGER NOT NULL REFERENCES environments(id) ON DELETE CASCADE,
    release_id INTEGER REFERENCES releases(id) ON DELETE SET NULL,
    UNIQUE (credential_id, environment_id)
);

CREATE TABLE IF NOT EXISTS tags (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    color TEXT NOT NULL,
    client_id INTEGER REFERENCES clients(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS archive_folders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    parent_id INTEGER REFERENCES archive_folders(id) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS archive_files (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    folder_id INTEGER REFERENCES archive_folders(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    file_type TEXT,
    last_modified TEXT,
    file_size INTEGER,
    extension TEXT,
    path TEXT NOT NULL,
    tag_id INTEGER REFERENCES tags(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS archive_links (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    folder_id INTEGER REFERENCES archive_folders(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    url TEXT NOT NULL,
    tag_id INTEGER REFERENCES tags(id) ON DELETE SET NULL
);

CREATE TABLE IF NOT EXISTS archive_favorites (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    item_type TEXT NOT NULL CHECK (item_type IN ('file', 'link')),
    item_id INTEGER NOT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (item_type, item_id)
);

CREATE TABLE IF NOT EXISTS client_contacts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    phone TEXT,
    mobile TEXT,
    email TEXT,
    role TEXT,
    note TEXT
);

CREATE TABLE IF NOT EXISTS client_notes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
    title TEXT NOT NULL,
    content_type TEXT NOT NULL CHECK (content_type IN ('text', 'table')),
    content TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS client_pending_relations (
    client_id INTEGER PRIMARY KEY,
    vpn_connection_name TEXT,
    resources_json TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS app_settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
"""


def get_connection() -> sqlite3.Connection:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    connection.execute("PRAGMA foreign_keys = ON;")
    return connection


def _migrate_roles_competence_column(connection: sqlite3.Connection) -> None:
    columns = connection.execute("PRAGMA table_info(roles);").fetchall()
    if not columns:
        return

    column_names = {row["name"] for row in columns}
    if "competence" in column_names:
        return

    if "competence_id" not in column_names:
        return

    with connection:
        connection.execute("PRAGMA foreign_keys = OFF;")
        connection.execute("ALTER TABLE roles RENAME TO roles_old;")
        connection.execute(
            """
            CREATE TABLE roles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                competence TEXT NOT NULL,
                multi_clients INTEGER NOT NULL DEFAULT 0 CHECK (multi_clients IN (0, 1)),
                display_order INTEGER CHECK (display_order BETWEEN 1 AND 20)
            );
            """
        )
        connection.execute(
            """
            INSERT INTO roles(id, name, competence, multi_clients, display_order)
            SELECT ro.id,
                   ro.name,
                   COALESCE(c.name, ''),
                   ro.multi_clients,
                   NULL
            FROM roles_old ro
            LEFT JOIN competences c ON c.id = ro.competence_id;
            """
        )
        connection.execute("DROP TABLE roles_old;")
        connection.execute("PRAGMA foreign_keys = ON;")


def _migrate_roles_display_order_column(connection: sqlite3.Connection) -> None:
    columns = connection.execute("PRAGMA table_info(roles);").fetchall()
    if not columns:
        return

    column_names = {row["name"] for row in columns}
    with connection:
        if "display_order" not in column_names:
            connection.execute(
                "ALTER TABLE roles ADD COLUMN display_order INTEGER CHECK (display_order BETWEEN 1 AND 20);"
            )

        rows = connection.execute(
            "SELECT id, display_order FROM roles ORDER BY id;"
        ).fetchall()

        used: set[int] = set()
        to_assign: list[int] = []
        for row in rows:
            raw_value = row["display_order"]
            role_id = int(row["id"])
            order_value: int | None

            try:
                order_value = int(raw_value) if raw_value is not None else None
            except (TypeError, ValueError):
                order_value = None

            if (
                order_value is None
                or order_value < 1
                or order_value > 20
                or order_value in used
            ):
                to_assign.append(role_id)
            else:
                used.add(order_value)

        free_values = [value for value in range(1, 21) if value not in used]
        for role_id in to_assign:
            next_value = free_values.pop(0) if free_values else None
            connection.execute(
                "UPDATE roles SET display_order=? WHERE id=?;",
                (next_value, role_id),
            )

        connection.execute(
            """
            CREATE UNIQUE INDEX IF NOT EXISTS idx_roles_display_order_unique
            ON roles(display_order)
            WHERE display_order IS NOT NULL;
            """
        )


def _migrate_product_credentials_schema(connection: sqlite3.Connection) -> None:
    table_exists = connection.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='product_credentials';"
    ).fetchone()
    if not table_exists:
        return

    columns = connection.execute("PRAGMA table_info(product_credentials);").fetchall()
    column_names = {row["name"] for row in columns}
    required_columns: list[tuple[str, str]] = [
        ("ip", "TEXT"),
        ("url", "TEXT"),
        ("host", "TEXT"),
        ("rdp_path", "TEXT"),
        ("domain", "TEXT"),
        ("login_name", "TEXT"),
        ("port", "INTEGER"),
        ("password_expiry", "INTEGER NOT NULL DEFAULT 0 CHECK (password_expiry IN (0, 1))"),
        ("password_inserted_at", "TEXT"),
        ("password_duration_days", "INTEGER"),
        ("password_end_date", "TEXT"),
        ("created_at", "TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP"),
    ]

    with connection:
        for column_name, definition in required_columns:
            if column_name in column_names:
                continue
            connection.execute(
                f"ALTER TABLE product_credentials ADD COLUMN {column_name} {definition};"
            )

        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS product_credential_environments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                credential_id INTEGER NOT NULL REFERENCES product_credentials(id) ON DELETE CASCADE,
                environment_id INTEGER NOT NULL REFERENCES environments(id) ON DELETE CASCADE,
                release_id INTEGER REFERENCES releases(id) ON DELETE SET NULL,
                UNIQUE (credential_id, environment_id)
            );
            """
        )


def _migrate_vpn_path_column(connection: sqlite3.Connection) -> None:
    columns = connection.execute("PRAGMA table_info(vpns);").fetchall()
    if not columns:
        return
    column_names = {row["name"] for row in columns}
    if "vpn_path" in column_names:
        return
    with connection:
        connection.execute("ALTER TABLE vpns ADD COLUMN vpn_path TEXT;")


def _migrate_password_ref_columns(connection: sqlite3.Connection) -> None:
    vpn_columns = connection.execute("PRAGMA table_info(vpns);").fetchall()
    if vpn_columns:
        vpn_names = {row["name"] for row in vpn_columns}
        with connection:
            if "password_ref" not in vpn_names:
                connection.execute("ALTER TABLE vpns ADD COLUMN password_ref TEXT;")

    cred_columns = connection.execute("PRAGMA table_info(product_credentials);").fetchall()
    if cred_columns:
        cred_names = {row["name"] for row in cred_columns}
        with connection:
            if "password_ref" not in cred_names:
                connection.execute("ALTER TABLE product_credentials ADD COLUMN password_ref TEXT;")


def _migrate_clients_link_column(connection: sqlite3.Connection) -> None:
    columns = connection.execute("PRAGMA table_info(clients);").fetchall()
    if not columns:
        return
    column_names = {row["name"] for row in columns}
    if "link" in column_names:
        return
    with connection:
        connection.execute("ALTER TABLE clients ADD COLUMN link TEXT;")


def _migrate_indexes(connection: sqlite3.Connection) -> None:
    # Archive
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_archive_files_folder_id ON archive_files(folder_id);"
    )
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_archive_files_tag_id ON archive_files(tag_id);"
    )
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_archive_files_extension ON archive_files(extension);"
    )
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_archive_links_folder_id ON archive_links(folder_id);"
    )
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_archive_links_tag_id ON archive_links(tag_id);"
    )
    # Credentials access patterns
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_product_credentials_client_product ON product_credentials(client_id, product_id);"
    )


def _migrate_archive_links_columns(connection: sqlite3.Connection) -> None:
    columns = connection.execute("PRAGMA table_info(archive_links);").fetchall()
    if not columns:
        return
    column_names = {row["name"] for row in columns}
    with connection:
        if "folder_id" not in column_names:
            connection.execute(
                "ALTER TABLE archive_links ADD COLUMN folder_id INTEGER REFERENCES archive_folders(id) ON DELETE SET NULL;"
            )
        if "tag_id" not in column_names:
            connection.execute(
                "ALTER TABLE archive_links ADD COLUMN tag_id INTEGER REFERENCES tags(id) ON DELETE SET NULL;"
            )


def _foreign_key_on_delete_action(connection: sqlite3.Connection, table: str, ref_table: str) -> str | None:
    rows = connection.execute(f"PRAGMA foreign_key_list({table});").fetchall()
    for row in rows:
        # row fields: (id, seq, table, from, to, on_update, on_delete, match)
        if str(row[2]) == ref_table:
            return str(row[6] or "").upper()
    return None


def _migrate_archive_folder_delete_behavior(connection: sqlite3.Connection) -> None:
    """Ensure deleting a folder deletes contained files/links.

    Old schema used ON DELETE SET NULL, which moves items to root.
    """

    files_on_delete = _foreign_key_on_delete_action(connection, "archive_files", "archive_folders")
    links_on_delete = _foreign_key_on_delete_action(connection, "archive_links", "archive_folders")

    if files_on_delete == "CASCADE" and links_on_delete == "CASCADE":
        return

    with connection:
        connection.execute("PRAGMA foreign_keys = OFF;")

        # archive_files
        if files_on_delete != "CASCADE":
            connection.execute("ALTER TABLE archive_files RENAME TO archive_files_old;")
            connection.execute(
                """
                CREATE TABLE archive_files (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    folder_id INTEGER REFERENCES archive_folders(id) ON DELETE CASCADE,
                    name TEXT NOT NULL,
                    file_type TEXT,
                    last_modified TEXT,
                    file_size INTEGER,
                    extension TEXT,
                    path TEXT NOT NULL,
                    tag_id INTEGER REFERENCES tags(id) ON DELETE SET NULL
                );
                """
            )
            connection.execute(
                """
                INSERT INTO archive_files(id, folder_id, name, file_type, last_modified, file_size, extension, path, tag_id)
                SELECT id, folder_id, name, file_type, last_modified, file_size, extension, path, tag_id
                FROM archive_files_old;
                """
            )
            connection.execute("DROP TABLE archive_files_old;")

        # archive_links
        if links_on_delete != "CASCADE":
            connection.execute("ALTER TABLE archive_links RENAME TO archive_links_old;")
            connection.execute(
                """
                CREATE TABLE archive_links (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    folder_id INTEGER REFERENCES archive_folders(id) ON DELETE CASCADE,
                    name TEXT NOT NULL,
                    url TEXT NOT NULL,
                    tag_id INTEGER REFERENCES tags(id) ON DELETE SET NULL
                );
                """
            )
            connection.execute(
                """
                INSERT INTO archive_links(id, folder_id, name, url, tag_id)
                SELECT id, folder_id, name, url, tag_id
                FROM archive_links_old;
                """
            )
            connection.execute("DROP TABLE archive_links_old;")

        connection.execute("PRAGMA foreign_keys = ON;")


def _migrate_tags_client_column(connection: sqlite3.Connection) -> None:
    columns = connection.execute("PRAGMA table_info(tags);").fetchall()
    if not columns:
        return
    column_names = {row["name"] for row in columns}
    if "client_id" in column_names:
        return
    with connection:
        connection.execute(
            "ALTER TABLE tags ADD COLUMN client_id INTEGER REFERENCES clients(id) ON DELETE SET NULL;"
        )


def _migrate_archive_favorites_table(connection: sqlite3.Connection) -> None:
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS archive_favorites (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_type TEXT NOT NULL CHECK (item_type IN ('file', 'link')),
            item_id INTEGER NOT NULL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE (item_type, item_id)
        );
        """
    )


def _migrate_client_contacts_table(connection: sqlite3.Connection) -> None:
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS client_contacts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
            name TEXT NOT NULL,
            phone TEXT,
            mobile TEXT,
            email TEXT,
            role TEXT,
            note TEXT
        );
        """
    )


def _migrate_client_notes_table(connection: sqlite3.Connection) -> None:
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS client_notes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            client_id INTEGER NOT NULL REFERENCES clients(id) ON DELETE CASCADE,
            title TEXT NOT NULL,
            content_type TEXT NOT NULL CHECK (content_type IN ('text', 'table')),
            content TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );
        """
    )


def _migrate_client_pending_relations_table(connection: sqlite3.Connection) -> None:
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS client_pending_relations (
            client_id INTEGER PRIMARY KEY,
            vpn_connection_name TEXT,
            resources_json TEXT NOT NULL,
            FOREIGN KEY (client_id) REFERENCES clients(id) ON DELETE CASCADE
        );
        """
    )


def init_db() -> None:
    connection = get_connection()
    try:
        connection.executescript(SCHEMA_SQL)
        _migrate_roles_competence_column(connection)
        _migrate_roles_display_order_column(connection)
        _migrate_product_credentials_schema(connection)
        _migrate_vpn_path_column(connection)
        _migrate_password_ref_columns(connection)
        _migrate_clients_link_column(connection)
        _migrate_archive_links_columns(connection)
        _migrate_archive_folder_delete_behavior(connection)
        _migrate_tags_client_column(connection)
        _migrate_archive_favorites_table(connection)
        _migrate_client_contacts_table(connection)
        _migrate_client_notes_table(connection)
        _migrate_client_pending_relations_table(connection)
        _migrate_indexes(connection)
        connection.commit()
    finally:
        connection.close()
