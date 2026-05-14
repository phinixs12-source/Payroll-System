import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), 'payroll.db')


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    conn = get_db()
    cur = conn.cursor()

    # 직원 테이블
    cur.execute("""
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name_code TEXT NOT NULL,
            department TEXT,
            contract_type TEXT DEFAULT '정규직',
            join_date TEXT,
            resign_date TEXT,
            severance_type TEXT DEFAULT 'DB',
            scheduled_hours REAL DEFAULT 8.0,
            overtime_hours REAL DEFAULT 0.0,
            base_salary INTEGER DEFAULT 0,
            is_exception INTEGER DEFAULT 0,
            match_key TEXT,
            deduction_count INTEGER DEFAULT 1,
            created_at TEXT DEFAULT (datetime('now', 'localtime')),
            updated_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    """)

    # 수당 항목 테이블
    cur.execute("""
        CREATE TABLE IF NOT EXISTS allowance_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            is_ordinary_wage INTEGER DEFAULT 0,
            severance_db INTEGER DEFAULT 0,
            severance_dc INTEGER DEFAULT 0,
            payment_condition TEXT DEFAULT 'fixed',
            condition_value INTEGER DEFAULT 0,
            apply_target TEXT DEFAULT 'all',
            is_active INTEGER DEFAULT 1,
            sort_order INTEGER DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    """)

    # 직원별 수당 금액 테이블
    cur.execute("""
        CREATE TABLE IF NOT EXISTS employee_allowances (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL,
            allowance_item_id INTEGER NOT NULL,
            amount INTEGER DEFAULT 0,
            FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE,
            FOREIGN KEY (allowance_item_id) REFERENCES allowance_items(id) ON DELETE CASCADE,
            UNIQUE(employee_id, allowance_item_id)
        )
    """)

    # 시스템 설정 테이블
    cur.execute("""
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT,
            description TEXT,
            updated_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    """)

    # 고지서 포맷 설정 테이블
    cur.execute("""
        CREATE TABLE IF NOT EXISTS notice_formats (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            agency_name TEXT NOT NULL,
            name_col TEXT,
            match_key_col TEXT,
            amount_col TEXT,
            start_row INTEGER DEFAULT 2,
            notes TEXT,
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    """)

    # 기본 수당 항목 삽입 (없을 때만)
    cur.execute("SELECT COUNT(*) FROM allowance_items")
    if cur.fetchone()[0] == 0:
        default_allowances = [
            ('핸드폰보조금', 0, 0, 0, 'min_days', 15, 'all', 1, 1),
            ('직무수당',    1, 1, 1, 'min_days', 15, 'all', 1, 2),
            ('기타수당',    0, 0, 0, 'fixed',    0,  'all', 1, 3),
        ]
        cur.executemany("""
            INSERT INTO allowance_items
            (name, is_ordinary_wage, severance_db, severance_dc,
             payment_condition, condition_value, apply_target, is_active, sort_order)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, default_allowances)

    # 기본 설정값 삽입
    default_settings = [
        ('company_name', '회사명', '회사 이름'),
        ('abnormal_threshold', '20', '이상값 경고 기준 (전월 대비 %)'),
    ]
    for key, value, desc in default_settings:
        cur.execute("""
            INSERT OR IGNORE INTO settings (key, value, description) VALUES (?, ?, ?)
        """, (key, value, desc))

    conn.commit()
    conn.close()


# ─── 직원 ──────────────────────────────────────────────
def get_all_employees():
    conn = get_db()
    rows = conn.execute("""
        SELECT * FROM employees ORDER BY department, name_code
    """).fetchall()
    conn.close()
    return rows


def get_employee(emp_id):
    conn = get_db()
    row = conn.execute("SELECT * FROM employees WHERE id=?", (emp_id,)).fetchone()
    conn.close()
    return row


def create_employee(data):
    conn = get_db()
    conn.execute("""
        INSERT INTO employees
        (name_code, department, contract_type, join_date, resign_date,
         severance_type, scheduled_hours, overtime_hours, base_salary,
         is_exception, match_key, deduction_count)
        VALUES (:name_code, :department, :contract_type, :join_date, :resign_date,
                :severance_type, :scheduled_hours, :overtime_hours, :base_salary,
                :is_exception, :match_key, :deduction_count)
    """, data)
    conn.commit()
    emp_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.close()
    return emp_id


def update_employee(emp_id, data):
    conn = get_db()
    conn.execute("""
        UPDATE employees SET
            name_code=:name_code, department=:department,
            contract_type=:contract_type, join_date=:join_date,
            resign_date=:resign_date, severance_type=:severance_type,
            scheduled_hours=:scheduled_hours, overtime_hours=:overtime_hours,
            base_salary=:base_salary, is_exception=:is_exception,
            match_key=:match_key, deduction_count=:deduction_count,
            updated_at=datetime('now','localtime')
        WHERE id=:id
    """, {**data, 'id': emp_id})
    conn.commit()
    conn.close()


def delete_employee(emp_id):
    conn = get_db()
    conn.execute("DELETE FROM employees WHERE id=?", (emp_id,))
    conn.commit()
    conn.close()


# ─── 수당 항목 ──────────────────────────────────────────
def get_all_allowances():
    conn = get_db()
    rows = conn.execute("""
        SELECT * FROM allowance_items ORDER BY sort_order, id
    """).fetchall()
    conn.close()
    return rows


def get_allowance(item_id):
    conn = get_db()
    row = conn.execute("SELECT * FROM allowance_items WHERE id=?", (item_id,)).fetchone()
    conn.close()
    return row


def create_allowance(data):
    conn = get_db()
    conn.execute("""
        INSERT INTO allowance_items
        (name, is_ordinary_wage, severance_db, severance_dc,
         payment_condition, condition_value, apply_target, is_active)
        VALUES (:name, :is_ordinary_wage, :severance_db, :severance_dc,
                :payment_condition, :condition_value, :apply_target, :is_active)
    """, data)
    conn.commit()
    conn.close()


def update_allowance(item_id, data):
    conn = get_db()
    conn.execute("""
        UPDATE allowance_items SET
            name=:name, is_ordinary_wage=:is_ordinary_wage,
            severance_db=:severance_db, severance_dc=:severance_dc,
            payment_condition=:payment_condition, condition_value=:condition_value,
            apply_target=:apply_target, is_active=:is_active
        WHERE id=:id
    """, {**data, 'id': item_id})
    conn.commit()
    conn.close()


def delete_allowance(item_id):
    conn = get_db()
    conn.execute("DELETE FROM allowance_items WHERE id=?", (item_id,))
    conn.commit()
    conn.close()


def toggle_allowance(item_id):
    conn = get_db()
    conn.execute("""
        UPDATE allowance_items SET is_active = CASE WHEN is_active=1 THEN 0 ELSE 1 END
        WHERE id=?
    """, (item_id,))
    conn.commit()
    conn.close()


# ─── 직원별 수당 금액 ────────────────────────────────────
def get_employee_allowances(emp_id):
    conn = get_db()
    rows = conn.execute("""
        SELECT ea.*, ai.name, ai.payment_condition, ai.condition_value
        FROM employee_allowances ea
        JOIN allowance_items ai ON ea.allowance_item_id = ai.id
        WHERE ea.employee_id = ?
        ORDER BY ai.sort_order
    """, (emp_id,)).fetchall()
    conn.close()
    return rows


def upsert_employee_allowance(emp_id, item_id, amount):
    conn = get_db()
    conn.execute("""
        INSERT INTO employee_allowances (employee_id, allowance_item_id, amount)
        VALUES (?, ?, ?)
        ON CONFLICT(employee_id, allowance_item_id) DO UPDATE SET amount=excluded.amount
    """, (emp_id, item_id, amount))
    conn.commit()
    conn.close()


# ─── 설정 ────────────────────────────────────────────────
def get_setting(key):
    conn = get_db()
    row = conn.execute("SELECT value FROM settings WHERE key=?", (key,)).fetchone()
    conn.close()
    return row['value'] if row else None


def set_setting(key, value):
    conn = get_db()
    conn.execute("""
        INSERT INTO settings (key, value) VALUES (?, ?)
        ON CONFLICT(key) DO UPDATE SET value=excluded.value,
        updated_at=datetime('now','localtime')
    """, (key, value))
    conn.commit()
    conn.close()


def get_all_settings():
    conn = get_db()
    rows = conn.execute("SELECT * FROM settings").fetchall()
    conn.close()
    return {r['key']: r['value'] for r in rows}
