import sqlite3
import os
from datetime import datetime

def _now():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

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

            CREATE TABLE IF NOT EXISTS orders (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                status      TEXT DEFAULT 'draft',
                data        TEXT,
                created_at  DATETIME DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS holidays (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                ngay       TEXT NOT NULL,
                loai       TEXT NOT NULL,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS banks (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                bank_code   TEXT UNIQUE NOT NULL,
                medium_name TEXT,
                bank_name   TEXT,
                short_name  TEXT
            );

            CREATE TABLE IF NOT EXISTS policies (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                ten_chinh_sach TEXT NOT NULL,
                pct_the_chap   REAL DEFAULT 0,
                status         TEXT DEFAULT 'draft',
                created_at     DATETIME DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS policy_tenors (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                policy_id   INTEGER NOT NULL,
                ky_han      INTEGER NOT NULL,
                loi_tuc     REAL DEFAULT 0,
                created_at  DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (policy_id) REFERENCES policies(id)
            );

            CREATE TABLE IF NOT EXISTS settings (
                key   TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS bond_lots (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                ma_vk       TEXT NOT NULL,
                ma_vsdc     TEXT,
                to_chuc_ph  TEXT,
                ngay_ph     TEXT,
                ngay_dh     TEXT,
                don_gia     REAL DEFAULT 0,
                created_at  DATETIME DEFAULT CURRENT_TIMESTAMP
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
        # Migration: thêm cột mới vào bảng đã tồn tại
        try:
            conn.execute("ALTER TABLE bond_lots ADD COLUMN don_gia REAL DEFAULT 0")
        except Exception:
            pass
        try:
            conn.execute("ALTER TABLE contracts ADD COLUMN order_id INTEGER")
        except Exception:
            pass
        try:
            conn.execute("ALTER TABLE orders ADD COLUMN review_status TEXT DEFAULT 'chua_kiem_tra'")
        except Exception:
            pass
        # Khởi tạo sequence counter nếu chưa có
        conn.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('next_order_seq', '10000')")


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
            "INSERT INTO templates (name, code, filename, file_path, created_at) VALUES (?,?,?,?,?)",
            (name, code, filename, file_path, _now())
        )


def delete_template(id):
    with get_db() as conn:
        conn.execute("DELETE FROM templates WHERE id = ?", (id,))


# ── Contracts ──

def get_contracts(search="", template_code="", page=1, per_page=20, date_from="", date_to=""):
    where = "WHERE 1=1"
    params = []
    if search:
        where += " AND (name LIKE ? OR row_data LIKE ?)"
        params.extend([f"%{search}%", f"%{search}%"])
    if template_code:
        where += " AND template_code = ?"
        params.append(template_code)

    # Lọc ngày giao dịch — data lưu dd/mm/yyyy, chuyển sang yyyy-mm-dd để so sánh
    if date_from or date_to:
        ngd = ("json_extract(row_data, '$.\"Ngày giao dịch\"')")
        iso = (f"CASE WHEN {ngd} IS NOT NULL AND length({ngd})=10 "
               f"THEN substr({ngd},7,4)||'-'||substr({ngd},4,2)||'-'||substr({ngd},1,2) "
               f"ELSE '' END")
        if date_from:
            where += f" AND ({iso}) >= ?"
            params.append(date_from)
        if date_to:
            where += f" AND ({iso}) <= ?"
            params.append(date_to)

    # Mỗi khách hàng chỉ hiện 1 dòng — đại diện bởi bản ghi có id nhỏ nhất
    group_expr = ("CASE WHEN INSTR(name, ' - ') > 0 "
                  "THEN SUBSTR(name, 1, INSTR(name, ' - ') - 1) "
                  "ELSE name END")
    group_by = f"GROUP BY batch_id, {group_expr}"

    offset = (page - 1) * per_page
    with get_db() as conn:
        total = conn.execute(
            f"SELECT COUNT(*) FROM (SELECT 1 FROM contracts {where} {group_by})", params
        ).fetchone()[0]
        rows = conn.execute(
            f"SELECT *, MIN(id) as id FROM contracts {where} "
            f"{group_by} ORDER BY MAX(created_at) DESC LIMIT ? OFFSET ?",
            params + [per_page, offset]
        ).fetchall()
    return rows, total


def get_contract(id):
    with get_db() as conn:
        return conn.execute(
            "SELECT * FROM contracts WHERE id = ?", (id,)
        ).fetchone()


def add_contract(name, template_id, template_name, template_code, pdf_path, row_data, batch_id, order_id=None):
    with get_db() as conn:
        conn.execute(
            """INSERT INTO contracts
               (name, template_id, template_name, template_code, pdf_path, row_data, batch_id, order_id, created_at)
               VALUES (?,?,?,?,?,?,?,?,?)""",
            (name, template_id, template_name, template_code, pdf_path, row_data, batch_id, order_id, _now())
        )


def get_contracts_by_order(order_id):
    with get_db() as conn:
        return conn.execute(
            "SELECT * FROM contracts WHERE order_id = ? ORDER BY created_at DESC",
            (order_id,)
        ).fetchall()


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


# ── Orders ──

def get_orders(search="", page=1, per_page=15, date_from="", date_to=""):
    where = "WHERE 1=1"
    params = []
    if search:
        where += " AND data LIKE ?"
        params.append(f"%{search}%")
    # Ngày giao dịch stored as DD/MM/YYYY → convert to YYYY-MM-DD for comparison
    _ngay = (
        "substr(json_extract(data,'$.Ngày giao dịch'),7,4)||'-'||"
        "substr(json_extract(data,'$.Ngày giao dịch'),4,2)||'-'||"
        "substr(json_extract(data,'$.Ngày giao dịch'),1,2)"
    )
    if date_from:
        where += f" AND {_ngay} >= ?"
        params.append(date_from)
    if date_to:
        where += f" AND {_ngay} <= ?"
        params.append(date_to)
    offset = (page - 1) * per_page
    with get_db() as conn:
        total = conn.execute(f"SELECT COUNT(*) FROM orders {where}", params).fetchone()[0]
        rows  = conn.execute(
            f"SELECT * FROM orders {where} ORDER BY created_at DESC LIMIT ? OFFSET ?",
            params + [per_page, offset]
        ).fetchall()
    return rows, total


def get_order(id):
    with get_db() as conn:
        return conn.execute("SELECT * FROM orders WHERE id = ?", (id,)).fetchone()


def find_order_by_so_hd(val):
    """Tìm lệnh có Số HĐ vay vốn hoặc Số HĐ Thế chấp khớp với val."""
    like = f"%{val}%"
    with get_db() as conn:
        return conn.execute(
            "SELECT * FROM orders WHERE data LIKE ? LIMIT 1", (like,)
        ).fetchone()


def add_order(data_json):
    with get_db() as conn:
        cur = conn.execute("INSERT INTO orders (data, created_at) VALUES (?,?)", (data_json, _now()))
        return cur.lastrowid


def update_order(id, data_json, status=None):
    with get_db() as conn:
        if status:
            conn.execute("UPDATE orders SET data=?, status=? WHERE id=?", (data_json, status, id))
        else:
            conn.execute("UPDATE orders SET data=? WHERE id=?", (data_json, id))


def update_order_review(id, review_status):
    with get_db() as conn:
        conn.execute("UPDATE orders SET review_status=? WHERE id=?", (review_status, id))


def delete_order(id):
    with get_db() as conn:
        conn.execute("DELETE FROM orders WHERE id=?", (id,))


def bulk_delete_orders(ids):
    with get_db() as conn:
        conn.executemany("DELETE FROM orders WHERE id=?", [(i,) for i in ids])


HOLIDAY_SEED = [
    ("2026-01-01","Tết Dương lịch"),("2026-01-02","Tết Dương lịch"),
    ("2026-02-16","Tết Âm lịch"),("2026-02-17","Tết Âm lịch"),
    ("2026-02-18","Tết Âm lịch"),("2026-02-19","Tết Âm lịch"),
    ("2026-02-20","Tết Âm lịch"),
    ("2026-04-27","Nghi lễ"),("2026-04-30","Nghi lễ"),("2026-05-01","Nghi lễ"),
    ("2026-08-31","Nghi lễ"),("2026-09-01","Nghi lễ"),
    ("2026-11-23","Nghi lễ"),("2026-11-24","Nghi lễ"),
    ("2027-01-01","Tết Dương lịch"),
    ("2027-02-03","Tết Âm lịch"),("2027-02-04","Tết Âm lịch"),
    ("2027-02-05","Tết Âm lịch"),("2027-02-08","Tết Âm lịch"),
    ("2027-02-09","Tết Âm lịch"),("2027-02-10","Tết Âm lịch"),
    ("2027-04-16","Giỗ Tổ Hùng Vương"),
    ("2027-04-30","Nghi lễ"),("2027-05-03","Nghi lễ"),
    ("2027-09-03","Nghi lễ"),("2027-11-24","Nghi lễ"),
]


def seed_holidays():
    with get_db() as conn:
        if conn.execute("SELECT COUNT(*) FROM holidays").fetchone()[0] == 0:
            conn.executemany(
                "INSERT OR IGNORE INTO holidays (ngay, loai) VALUES (?,?)",
                HOLIDAY_SEED
            )


def get_holidays(search=""):
    where, params = "WHERE 1=1", []
    if search:
        where += " AND (ngay LIKE ? OR loai LIKE ?)"
        params.extend([f"%{search}%", f"%{search}%"])
    with get_db() as conn:
        return conn.execute(
            f"SELECT * FROM holidays {where} ORDER BY ngay ASC", params
        ).fetchall()


def get_holiday(id):
    with get_db() as conn:
        return conn.execute("SELECT * FROM holidays WHERE id=?", (id,)).fetchone()


def add_holiday(ngay, loai):
    with get_db() as conn:
        cur = conn.execute("INSERT INTO holidays (ngay, loai) VALUES (?,?)", (ngay, loai))
        return cur.lastrowid


def update_holiday(id, ngay, loai):
    with get_db() as conn:
        conn.execute("UPDATE holidays SET ngay=?, loai=? WHERE id=?", (ngay, loai, id))


def delete_holiday(id):
    with get_db() as conn:
        conn.execute("DELETE FROM holidays WHERE id=?", (id,))


BANK_SEED = [
    ("100000","Sumitomo","NH SUMITOMO MITSUI BANKING CORP","SMBC"),
    ("100001","Citibank","Ngân hàng Citibank Việt Nam","Citibank"),
    ("157979","Scb Bank","Ngân hàng TMCP Sài Gòn","SCB"),
    ("166888","VietABank","Ngân hàng TMCP Việt Á","VAB"),
    ("180906","VibBank","Ngân hàng TMCP Quốc Tế Việt Nam","VIB"),
    ("191919","AbBank","Ngân hàng TMCP An Bình","ABB"),
    ("422589","CimbBank","Ngân hàng TNHH Một Thành Viên CIMB Việt Nam","CIMB"),
    ("452999","EximBank","Ngân hàng TMCP Xuất Nhập khẩu Việt Nam","EIB"),
    ("458761","Hsbc","Ngân hàng HSBC Việt Nam","HSBC"),
    ("546034","Cake","Ngân hàng số CAKE by VPBank","CAKE"),
    ("546035","UBank","Ngân hàng số Ubank by VPBank","Ubank"),
    ("686868","VietComBank","Ngân hàng TMCP Ngoại thương Việt Nam","VCB"),
    ("801011","Nhb","Ngân hàng NONGHYUP","NHB"),
    ("818188","NcbBank","Ngân hàng TMCP Quốc Dân","NCB"),
    ("888899","TechComBank","Ngân hàng TMCP Kỹ Thương Việt Nam","TCB"),
    ("888999","Ivb","Ngân hàng TNHH Indovina","IVB"),
    ("963399","UMEE","Ngân hàng số UMEE by KienLongBank","UMEE"),
    ("970400","SaiGonBank","Ngân hàng TMCP Sài Gòn Công Thương","SGICB"),
    ("970403","SacomBank","Ngân hàng TMCP Sài Gòn Thương Tín","STB"),
    ("970405","AgriBank","Ngân hàng Nông nghiệp và Phát triển Nông thôn Việt Nam","AGRIBANK"),
    ("970406","DongABank","Ngân hàng TMCP Đông Á","DAB"),
    ("970407","TechComBank","Ngân hàng TMCP Kỹ Thương Việt Nam","TCB"),
    ("970408","GpBank","Ngân hàng Thương mại TNHH Một Thành Viên Dầu Khí Toàn Cầu","GPB"),
    ("970409","BacABank","Ngân hàng TMCP Bắc Á","BAB"),
    ("970410","Scvn","Ngân hàng TNHH Một Thành Viên Standard Chartered","SCVN"),
    ("970412","PvComBank","Ngân hàng TMCP Đại Chúng Việt Nam","PVCB"),
    ("970414","OceanBank","Ngân hàng TNHH Một Thành Viên Đại Dương","OJB"),
    ("970415","VietTinBank","Ngân hàng TMCP Công Thương Việt Nam","CTG"),
    ("970416","AcbBank","Ngân hàng TMCP Á Châu","ACB"),
    ("970418","Bidv","Ngân hàng Đầu tư và Phát triển Việt Nam","BIDV"),
    ("970419","NcbBank","Ngân hàng TMCP Quốc Dân","NCB"),
    ("970421","VrbBank","Ngân hàng liên doanh Việt Nga","VRB"),
    ("970422","MbBank","Ngân hàng TMCP Quân Đội","MB"),
    ("970423","TPBank","Ngân hàng TMCP Tiên Phong","TPB"),
    ("970424","Shbvn","Ngân hàng TNHH Một Thành Viên Shinhan Việt Nam","SHBVN"),
    ("970425","AbBank","Ngân hàng TMCP An Bình","ABB"),
    ("970426","Msb","Ngân hàng TMCP Hàng Hải Việt Nam","MSB"),
    ("970427","VietABank","Ngân hàng TMCP Việt Á","VAB"),
    ("970428","NamABank","Ngân hàng TMCP Nam Á","NAB"),
    ("970429","Scb Bank","Ngân hàng TMCP Sài Gòn","SCB"),
    ("970430","PgBank","Ngân hàng TMCP Thịnh vượng và Phát triển","PGB"),
    ("970431","EximBank","Ngân hàng TMCP Xuất Nhập khẩu Việt Nam","EIB"),
    ("970432","VPBank","Ngân hàng TMCP Việt Nam Thịnh Vượng","VPB"),
    ("970433","VietBank","Ngân hàng TMCP Việt Nam Thương Tín","VIETBANK"),
    ("970434","Ivb","Ngân hàng TNHH Indovina","IVB"),
    ("970436","VietComBank","Ngân hàng TMCP Ngoại thương Việt Nam","VCB"),
    ("970437","HDBank","Ngân hàng TMCP Phát triển TP.HCM","HDB"),
    ("970438","BaoVietBank","Ngân hàng TMCP Bảo Việt","BVB"),
    ("970439","PbvnBank","Ngân hàng TNHH Một Thành Viên Public Việt Nam","PBVN"),
    ("970440","SeaBank","Ngân hàng TMCP Đông Nam Á","SEAB"),
    ("970441","VibBank","Ngân hàng TMCP Quốc Tế Việt Nam","VIB"),
    ("970442","HlbvnBank","Ngân hàng TNHH Một Thành Viên Hong Leong Việt Nam","HLBVN"),
    ("970443","ShbBank","Ngân hàng TMCP Sài Gòn - Hà Nội","SHB"),
    ("970444","CBBANK","TM TNHH MTV Xây Dựng Việt Nam","CBBANK"),
    ("970446","CcfBank","Ngân hàng Hợp Tác Xã Việt Nam","CCF"),
    ("970448","OcBank","Ngân hàng TMCP Phương Đông","OCB"),
    ("970449","LienVietPost","Ngân hàng TMCP Bưu Điện Liên Việt","LPB"),
    ("970452","KienlongBank","NH TMCP Kiên Long (KienLong Bank)","KLB"),
    ("970454","VietCapitalBank","Ngân hàng TMCP Bản Việt","VCCB"),
    ("970455","Ibk-Hn","Ngân hàng Công nghiệp Hàn Quốc - Chi nhánh Hà Nội","IBK-HN"),
    ("970456","Ibk-Hcm","Ngân hàng Industrial Bank of Korea - Chi nhánh Hồ Chí Minh","IBK-HCM"),
    ("970457","WvnBank","Ngân hàng Woori Bank Việt Nam","WVN"),
    ("970458","UobBank","Ngân hàng UOB Việt Nam","UOB"),
    ("970468","SeaBank","Ngân hàng TMCP Đông Nam Á","SEAB"),
    ("970488","Bidv","Ngân hàng Đầu tư và Phát triển Việt Nam","BIDV"),
    ("981957","VPBank","Ngân hàng TMCP Việt Nam Thịnh Vượng","VPB"),
]


def seed_banks():
    with get_db() as conn:
        conn.execute("DELETE FROM banks")
        conn.executemany(
            "INSERT INTO banks (bank_code, medium_name, bank_name, short_name) VALUES (?,?,?,?)",
            BANK_SEED
        )


# ── Order sequence ──

def get_next_order_seq():
    with get_db() as conn:
        row = conn.execute("SELECT value FROM settings WHERE key='next_order_seq'").fetchone()
        return int(row[0]) if row else 10000


def bump_order_seq():
    with get_db() as conn:
        conn.execute(
            "UPDATE settings SET value=CAST(CAST(value AS INTEGER)+1 AS TEXT) WHERE key='next_order_seq'"
        )


def get_bank_by_code(bank_code):
    with get_db() as conn:
        return conn.execute(
            "SELECT * FROM banks WHERE bank_code = ?", (bank_code,)
        ).fetchone()

def get_banks(search=""):
    where  = "WHERE 1=1"
    params = []
    if search:
        where += " AND (bank_code LIKE ? OR bank_name LIKE ? OR short_name LIKE ?)"
        params.extend([f"%{search}%", f"%{search}%", f"%{search}%"])
    with get_db() as conn:
        return conn.execute(
            f"SELECT * FROM banks {where} ORDER BY bank_code ASC", params
        ).fetchall()


# ── Policies (Chính sách sản phẩm) ──

def get_policies():
    with get_db() as conn:
        return conn.execute("SELECT * FROM policies ORDER BY created_at DESC").fetchall()

def get_policy(id):
    with get_db() as conn:
        return conn.execute("SELECT * FROM policies WHERE id=?", (id,)).fetchone()

def add_policy(ten_chinh_sach, pct_the_chap):
    with get_db() as conn:
        cur = conn.execute(
            "INSERT INTO policies (ten_chinh_sach, pct_the_chap) VALUES (?,?)",
            (ten_chinh_sach, pct_the_chap)
        )
        return cur.lastrowid

def update_policy(id, ten_chinh_sach, pct_the_chap):
    with get_db() as conn:
        conn.execute(
            "UPDATE policies SET ten_chinh_sach=?, pct_the_chap=? WHERE id=?",
            (ten_chinh_sach, pct_the_chap, id)
        )

def update_policy_status(id, status):
    with get_db() as conn:
        conn.execute("UPDATE policies SET status=? WHERE id=?", (status, id))

def delete_policy(id):
    with get_db() as conn:
        conn.execute("DELETE FROM policy_tenors WHERE policy_id=?", (id,))
        conn.execute("DELETE FROM policies WHERE id=?", (id,))

def get_policy_tenors(policy_id):
    with get_db() as conn:
        return conn.execute(
            "SELECT * FROM policy_tenors WHERE policy_id=? ORDER BY ky_han ASC",
            (policy_id,)
        ).fetchall()

def add_policy_tenor(policy_id, ky_han, loi_tuc):
    with get_db() as conn:
        cur = conn.execute(
            "INSERT INTO policy_tenors (policy_id, ky_han, loi_tuc) VALUES (?,?,?)",
            (policy_id, ky_han, loi_tuc)
        )
        return cur.lastrowid

def update_policy_tenor(id, ky_han, loi_tuc):
    with get_db() as conn:
        conn.execute(
            "UPDATE policy_tenors SET ky_han=?, loi_tuc=? WHERE id=?",
            (ky_han, loi_tuc, id)
        )

def delete_policy_tenor(id):
    with get_db() as conn:
        conn.execute("DELETE FROM policy_tenors WHERE id=?", (id,))


# ── Bond Lots (Lô Trái Phiếu) ──

def get_bond_lots(search="", page=1, per_page=20):
    where = "WHERE 1=1"
    params = []
    if search:
        where += " AND (ma_vk LIKE ? OR ma_vsdc LIKE ? OR to_chuc_ph LIKE ?)"
        params.extend([f"%{search}%", f"%{search}%", f"%{search}%"])
    offset = (page - 1) * per_page
    with get_db() as conn:
        total = conn.execute(f"SELECT COUNT(*) FROM bond_lots {where}", params).fetchone()[0]
        rows  = conn.execute(
            f"SELECT * FROM bond_lots {where} ORDER BY created_at DESC LIMIT ? OFFSET ?",
            params + [per_page, offset]
        ).fetchall()
    return rows, total


def get_bond_lot(id):
    with get_db() as conn:
        return conn.execute("SELECT * FROM bond_lots WHERE id = ?", (id,)).fetchone()


def add_bond_lot(ma_vk, ma_vsdc, to_chuc_ph, ngay_ph, ngay_dh, don_gia=0):
    with get_db() as conn:
        cur = conn.execute(
            "INSERT INTO bond_lots (ma_vk, ma_vsdc, to_chuc_ph, ngay_ph, ngay_dh, don_gia) VALUES (?,?,?,?,?,?)",
            (ma_vk, ma_vsdc, to_chuc_ph, ngay_ph, ngay_dh, don_gia)
        )
        return cur.lastrowid


def update_bond_lot(id, ma_vk, ma_vsdc, to_chuc_ph, ngay_ph, ngay_dh, don_gia=0):
    with get_db() as conn:
        conn.execute(
            "UPDATE bond_lots SET ma_vk=?, ma_vsdc=?, to_chuc_ph=?, ngay_ph=?, ngay_dh=?, don_gia=? WHERE id=?",
            (ma_vk, ma_vsdc, to_chuc_ph, ngay_ph, ngay_dh, don_gia, id)
        )


def delete_bond_lot(id):
    with get_db() as conn:
        conn.execute("DELETE FROM bond_lots WHERE id=?", (id,))


# ── Batches ──

def add_batch(id, excel_filename, contract_count):
    with get_db() as conn:
        conn.execute(
            "INSERT OR IGNORE INTO batches (id, excel_filename, contract_count) VALUES (?,?,?)",
            (id, excel_filename, contract_count)
        )
