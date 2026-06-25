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

def get_contracts(search="", template_code="", page=1, per_page=20):
    where = "WHERE 1=1"
    params = []
    if search:
        where += " AND (name LIKE ? OR row_data LIKE ?)"
        params.extend([f"%{search}%", f"%{search}%"])
    if template_code:
        where += " AND template_code = ?"
        params.append(template_code)

    offset = (page - 1) * per_page
    with get_db() as conn:
        total = conn.execute(
            f"SELECT COUNT(*) FROM contracts {where}", params
        ).fetchone()[0]
        rows = conn.execute(
            f"SELECT * FROM contracts {where} ORDER BY created_at DESC LIMIT ? OFFSET ?",
            params + [per_page, offset]
        ).fetchall()
    return rows, total


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


def get_contracts_by_batch(batch_id):
    with get_db() as conn:
        return conn.execute(
            "SELECT * FROM contracts WHERE batch_id = ? ORDER BY template_code",
            (batch_id,)
        ).fetchall()


def get_contracts_same_customer(batch_id, customer_name):
    """Lấy tất cả HĐ cùng khách hàng trong 1 batch (HDVV, HDTC...)"""
    with get_db() as conn:
        return conn.execute(
            "SELECT * FROM contracts WHERE batch_id = ? AND name LIKE ? ORDER BY template_code",
            (batch_id, f"{customer_name}%")
        ).fetchall()


def update_contracts_row_data(ids, row_data_json):
    with get_db() as conn:
        for cid in ids:
            conn.execute("UPDATE contracts SET row_data = ? WHERE id = ?",
                         (row_data_json, cid))


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
