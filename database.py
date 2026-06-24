import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "data", "contracts.db")


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")  # better concurrency for 5 users
    return conn


def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    with get_db() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS templates (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                name       TEXT NOT NULL,
                code       TEXT NOT NULL UNIQUE,
                filename   TEXT NOT NULL,
                file_path  TEXT NOT NULL,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS batches (
                id              TEXT PRIMARY KEY,
                excel_filename  TEXT NOT NULL,
                contract_count  INTEGER DEFAULT 0,
                created_at      DATETIME DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS contracts (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                name          TEXT NOT NULL,
                template_id   INTEGER,
                template_name TEXT,
                template_code TEXT,
                pdf_path      TEXT,
                row_data      TEXT,
                batch_id      TEXT,
                created_at    DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (template_id) REFERENCES templates(id)
            );
        """)


# ── Templates ──

def get_templates():
    with get_db() as conn:
        return conn.execute(
            "SELECT * FROM templates ORDER BY created_at DESC"
        ).fetchall()


def get_template(id):
    with get_db() as conn:
        return conn.execute(
            "SELECT * FROM templates WHERE id = ?", (id,)
        ).fetchone()


def get_template_by_code(code):
    with get_db() as conn:
        return conn.execute(
            "SELECT * FROM templates WHERE code = ?", (code,)
        ).fetchone()


def add_template(name, code, filename, file_path):
    with get_db() as conn:
        conn.execute(
            "INSERT INTO templates (name, code, filename, file_path) VALUES (?,?,?,?)",
            (name, code, filename, file_path)
        )


def delete_template(id):
    with get_db() as conn:
        conn.execute("DELETE FROM templates WHERE id = ?", (id,))


# ── Contracts ──

def get_contracts(search="", template_code=""):
    sql = "SELECT * FROM contracts WHERE 1=1"
    params = []
    if search:
        sql += " AND name LIKE ?"
        params.append(f"%{search}%")
    if template_code:
        sql += " AND template_code = ?"
        params.append(template_code)
    sql += " ORDER BY created_at DESC"
    with get_db() as conn:
        return conn.execute(sql, params).fetchall()


def get_contract(id):
    with get_db() as conn:
        return conn.execute(
            "SELECT * FROM contracts WHERE id = ?", (id,)
        ).fetchone()


def add_contract(name, template_id, template_name, template_code, pdf_path, row_data, batch_id):
    with get_db() as conn:
        conn.execute(
            """INSERT INTO contracts
               (name, template_id, template_name, template_code, pdf_path, row_data, batch_id)
               VALUES (?,?,?,?,?,?,?)""",
            (name, template_id, template_name, template_code, pdf_path, row_data, batch_id)
        )


def delete_contract(id):
    with get_db() as conn:
        conn.execute("DELETE FROM contracts WHERE id = ?", (id,))


# ── Batches ──

def add_batch(id, excel_filename, contract_count):
    with get_db() as conn:
        conn.execute(
            "INSERT OR IGNORE INTO batches (id, excel_filename, contract_count) VALUES (?,?,?)",
            (id, excel_filename, contract_count)
        )
