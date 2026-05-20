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
            name_code TEXT NOT NULL UNIQUE,
            department TEXT,
            contract_type TEXT DEFAULT '정규직',
            join_date TEXT,
            resign_date TEXT,
            severance_type TEXT DEFAULT 'DB',
            scheduled_hours REAL DEFAULT 8.0,
            overtime_hours REAL DEFAULT 0.0,
            annual_salary INTEGER DEFAULT 0,
            base_salary INTEGER DEFAULT 0,
            ordinary_wage INTEGER DEFAULT 0,
            match_key TEXT,
            research_tax_exempt TEXT DEFAULT 'N',
            child_deduction TEXT DEFAULT 'N',
            child_count INTEGER DEFAULT 0,
            is_exception TEXT DEFAULT 'N',
            employee_type TEXT DEFAULT '직원',
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

    # 공제 항목 테이블 (수당 항목과 동일 구조)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS deduction_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            is_insurance INTEGER DEFAULT 0,
            payment_condition TEXT DEFAULT 'fixed',
            condition_value INTEGER DEFAULT 0,
            apply_target TEXT DEFAULT 'all',
            is_active INTEGER DEFAULT 1,
            sort_order INTEGER DEFAULT 0,
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        )
    """)

    # 직원별 공제 금액 테이블 (수당 금액과 동일 구조)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS employee_deductions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL,
            deduction_item_id INTEGER NOT NULL,
            amount INTEGER DEFAULT 0,
            FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE,
            FOREIGN KEY (deduction_item_id) REFERENCES deduction_items(id) ON DELETE CASCADE,
            UNIQUE(employee_id, deduction_item_id)
        )
    """)

    # 기본 공제 항목 삽입
    cur.execute("SELECT COUNT(*) FROM deduction_items")
    if cur.fetchone()[0] == 0:
        default_deductions = [
            ('국민연금',         1, 'fixed', 0, 'all', 1, 1),
            ('건강보험',         1, 'fixed', 0, 'all', 1, 2),
            ('장기요양보험',     1, 'fixed', 0, 'all', 1, 3),
            ('고용보험',         1, 'fixed', 0, 'all', 1, 4),
            ('주택자금대출상환', 0, 'fixed', 0, 'all', 1, 5),
            ('기타공제',         0, 'fixed', 0, 'all', 1, 6),
        ]
        cur.executemany("""
            INSERT INTO deduction_items
            (name, is_insurance, payment_condition, condition_value,
             apply_target, is_active, sort_order)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, default_deductions)

    # 급여 귀속기간 테이블
    cur.execute("""
        CREATE TABLE IF NOT EXISTS payroll_periods (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            payment_date TEXT,
            status TEXT DEFAULT 'step1',
            created_at TEXT DEFAULT (datetime('now', 'localtime')),
            UNIQUE(year, month)
        )
    """)

    # 급여대장 항목 테이블 (직원 1명 = 1행)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS payroll_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            period_id INTEGER NOT NULL,
            employee_id INTEGER NOT NULL,
            scheduled_days INTEGER DEFAULT 0,
            work_days INTEGER DEFAULT 0,
            is_new_hire INTEGER DEFAULT 0,
            is_resigned INTEGER DEFAULT 0,
            base_salary INTEGER DEFAULT 0,
            overtime_pay INTEGER DEFAULT 0,
            total_allowance INTEGER DEFAULT 0,
            gross_pay INTEGER DEFAULT 0,
            total_deduction INTEGER DEFAULT 0,
            net_pay INTEGER DEFAULT 0,
            notes TEXT,
            FOREIGN KEY (period_id) REFERENCES payroll_periods(id) ON DELETE CASCADE,
            FOREIGN KEY (employee_id) REFERENCES employees(id),
            UNIQUE(period_id, employee_id)
        )
    """)

    # 수당 지급 내역 테이블
    cur.execute("""
        CREATE TABLE IF NOT EXISTS payroll_allowance_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entry_id INTEGER NOT NULL,
            allowance_item_id INTEGER NOT NULL,
            amount INTEGER DEFAULT 0,
            is_paid INTEGER DEFAULT 1,
            FOREIGN KEY (entry_id) REFERENCES payroll_entries(id) ON DELETE CASCADE,
            FOREIGN KEY (allowance_item_id) REFERENCES allowance_items(id),
            UNIQUE(entry_id, allowance_item_id)
        )
    """)

    # 공제 내역 테이블
    cur.execute("""
        CREATE TABLE IF NOT EXISTS payroll_deduction_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            entry_id INTEGER NOT NULL,
            deduction_name TEXT NOT NULL,
            amount INTEGER DEFAULT 0,
            FOREIGN KEY (entry_id) REFERENCES payroll_entries(id) ON DELETE CASCADE
        )
    """)

    # 기존 DB 컬럼 마이그레이션 (없으면 추가)
    for col, definition in [
        ('annual_salary', 'INTEGER DEFAULT 0'),
        ('ordinary_wage', 'INTEGER DEFAULT 0'),
    ]:
        try:
            cur.execute(f"ALTER TABLE employees ADD COLUMN {col} {definition}")
        except Exception:
            pass

    # payroll_allowance_entries 메모 컬럼 마이그레이션
    try:
        cur.execute("ALTER TABLE payroll_allowance_entries ADD COLUMN notes TEXT")
    except Exception:
        pass

    # 기본 수당 항목 삽입
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

    # 기본 설정값
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
    rows = conn.execute(
        "SELECT * FROM employees ORDER BY department, name_code"
    ).fetchall()
    conn.close()
    return rows


def get_employee(emp_id):
    conn = get_db()
    row = conn.execute("SELECT * FROM employees WHERE id=?", (emp_id,)).fetchone()
    conn.close()
    return row


def get_employee_by_code(name_code):
    conn = get_db()
    row = conn.execute("SELECT * FROM employees WHERE name_code=?", (name_code,)).fetchone()
    conn.close()
    return row


def create_employee(data):
    conn = get_db()
    conn.execute("""
        INSERT INTO employees
        (name_code, department, contract_type, join_date, resign_date,
         severance_type, scheduled_hours, overtime_hours,
         annual_salary, base_salary, ordinary_wage,
         match_key, research_tax_exempt, child_deduction, child_count,
         is_exception, employee_type)
        VALUES (:name_code, :department, :contract_type, :join_date, :resign_date,
                :severance_type, :scheduled_hours, :overtime_hours,
                :annual_salary, :base_salary, :ordinary_wage,
                :match_key, :research_tax_exempt, :child_deduction, :child_count,
                :is_exception, :employee_type)
    """, data)
    conn.commit()
    emp_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.close()
    return emp_id


def upsert_employee(data):
    """코드 기준으로 있으면 업데이트, 없으면 삽입"""
    conn = get_db()
    existing = conn.execute(
        "SELECT id FROM employees WHERE name_code=?", (data['name_code'],)
    ).fetchone()
    if existing:
        conn.execute("""
            UPDATE employees SET
                department=:department, contract_type=:contract_type,
                join_date=:join_date, resign_date=:resign_date,
                severance_type=:severance_type, scheduled_hours=:scheduled_hours,
                overtime_hours=:overtime_hours,
                annual_salary=:annual_salary, base_salary=:base_salary,
                ordinary_wage=:ordinary_wage,
                match_key=:match_key, research_tax_exempt=:research_tax_exempt,
                child_deduction=:child_deduction, child_count=:child_count,
                is_exception=:is_exception, employee_type=:employee_type,
                updated_at=datetime('now','localtime')
            WHERE name_code=:name_code
        """, data)
        emp_id = existing['id']
    else:
        conn.execute("""
            INSERT INTO employees
            (name_code, department, contract_type, join_date, resign_date,
             severance_type, scheduled_hours, overtime_hours,
             annual_salary, base_salary, ordinary_wage,
             match_key, research_tax_exempt, child_deduction, child_count,
             is_exception, employee_type)
            VALUES (:name_code, :department, :contract_type, :join_date, :resign_date,
                    :severance_type, :scheduled_hours, :overtime_hours,
                    :annual_salary, :base_salary, :ordinary_wage,
                    :match_key, :research_tax_exempt, :child_deduction, :child_count,
                    :is_exception, :employee_type)
        """, data)
        emp_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.commit()
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
            annual_salary=:annual_salary, base_salary=:base_salary,
            ordinary_wage=:ordinary_wage, match_key=:match_key,
            research_tax_exempt=:research_tax_exempt,
            child_deduction=:child_deduction, child_count=:child_count,
            is_exception=:is_exception, employee_type=:employee_type,
            updated_at=datetime('now','localtime')
        WHERE id=:id
    """, {**data, 'id': emp_id})
    conn.commit()
    conn.close()


def update_employee_salary(name_code, annual_salary):
    """Step2 연봉 업로드용 — annual_salary만 갱신, base_salary는 상세등록 시 계산"""
    conn = get_db()
    conn.execute(
        "UPDATE employees SET annual_salary=?, updated_at=datetime('now','localtime') WHERE name_code=?",
        (annual_salary, name_code)
    )
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
    rows = conn.execute(
        "SELECT * FROM allowance_items ORDER BY sort_order, id"
    ).fetchall()
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
        UPDATE allowance_items
        SET is_active = CASE WHEN is_active=1 THEN 0 ELSE 1 END
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


# ─── 공제 항목 ──────────────────────────────────────────────

def get_all_deductions():
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM deduction_items ORDER BY sort_order, id"
    ).fetchall()
    conn.close()
    return rows


def get_deduction(item_id):
    conn = get_db()
    row = conn.execute("SELECT * FROM deduction_items WHERE id=?", (item_id,)).fetchone()
    conn.close()
    return row


def create_deduction(data):
    conn = get_db()
    conn.execute("""
        INSERT INTO deduction_items
        (name, is_insurance, payment_condition, condition_value, apply_target, is_active)
        VALUES (:name, :is_insurance, :payment_condition, :condition_value, :apply_target, :is_active)
    """, data)
    conn.commit()
    conn.close()


def update_deduction(item_id, data):
    conn = get_db()
    conn.execute("""
        UPDATE deduction_items SET
            name=:name, is_insurance=:is_insurance,
            payment_condition=:payment_condition, condition_value=:condition_value,
            apply_target=:apply_target, is_active=:is_active
        WHERE id=:id
    """, {**data, 'id': item_id})
    conn.commit()
    conn.close()


def delete_deduction(item_id):
    conn = get_db()
    conn.execute("DELETE FROM deduction_items WHERE id=?", (item_id,))
    conn.commit()
    conn.close()


def toggle_deduction(item_id):
    conn = get_db()
    conn.execute("""
        UPDATE deduction_items
        SET is_active = CASE WHEN is_active=1 THEN 0 ELSE 1 END
        WHERE id=?
    """, (item_id,))
    conn.commit()
    conn.close()


# ─── 직원별 공제 금액 ─────────────────────────────────────────

def get_employee_deductions(emp_id):
    conn = get_db()
    rows = conn.execute("""
        SELECT ed.*, di.name, di.is_insurance, di.payment_condition
        FROM employee_deductions ed
        JOIN deduction_items di ON ed.deduction_item_id = di.id
        WHERE ed.employee_id = ?
        ORDER BY di.sort_order
    """, (emp_id,)).fetchall()
    conn.close()
    return rows


def upsert_employee_deduction(emp_id, item_id, amount):
    conn = get_db()
    conn.execute("""
        INSERT INTO employee_deductions (employee_id, deduction_item_id, amount)
        VALUES (?, ?, ?)
        ON CONFLICT(employee_id, deduction_item_id) DO UPDATE SET amount=excluded.amount
    """, (emp_id, item_id, amount))
    conn.commit()
    conn.close()


# ─── 급여 귀속기간 ──────────────────────────────────────────

def get_all_payroll_periods():
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM payroll_periods ORDER BY year DESC, month DESC"
    ).fetchall()
    conn.close()
    return rows


def get_payroll_period(period_id):
    conn = get_db()
    row = conn.execute("SELECT * FROM payroll_periods WHERE id=?", (period_id,)).fetchone()
    conn.close()
    return row


def get_payroll_period_by_ym(year, month):
    conn = get_db()
    row = conn.execute(
        "SELECT * FROM payroll_periods WHERE year=? AND month=?", (year, month)
    ).fetchone()
    conn.close()
    return row


def create_payroll_period(year, month, payment_date):
    conn = get_db()
    conn.execute(
        "INSERT INTO payroll_periods (year, month, payment_date) VALUES (?, ?, ?)",
        (year, month, payment_date)
    )
    conn.commit()
    period_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.close()
    return period_id


def update_payroll_period_status(period_id, status):
    conn = get_db()
    conn.execute("UPDATE payroll_periods SET status=? WHERE id=?", (status, period_id))
    conn.commit()
    conn.close()


def delete_payroll_period(period_id):
    conn = get_db()
    conn.execute("DELETE FROM payroll_periods WHERE id=?", (period_id,))
    conn.commit()
    conn.close()


# ─── 급여대장 항목 ──────────────────────────────────────────

def get_payroll_entries(period_id):
    conn = get_db()
    rows = conn.execute("""
        SELECT pe.*, e.name_code, e.department, e.join_date, e.resign_date,
               e.annual_salary, e.ordinary_wage, e.employee_type,
               e.scheduled_hours, e.overtime_hours
        FROM payroll_entries pe
        JOIN employees e ON pe.employee_id = e.id
        WHERE pe.period_id = ?
        ORDER BY e.department, e.name_code
    """, (period_id,)).fetchall()
    conn.close()
    return rows


def get_payroll_entry(entry_id):
    conn = get_db()
    row = conn.execute("""
        SELECT pe.*, e.name_code, e.department, e.join_date, e.resign_date,
               e.annual_salary, e.base_salary AS emp_base, e.ordinary_wage
        FROM payroll_entries pe
        JOIN employees e ON pe.employee_id = e.id
        WHERE pe.id = ?
    """, (entry_id,)).fetchone()
    conn.close()
    return row


def create_payroll_entry(data):
    conn = get_db()
    conn.execute("""
        INSERT OR IGNORE INTO payroll_entries
        (period_id, employee_id, scheduled_days, work_days,
         is_new_hire, is_resigned, base_salary, overtime_pay, gross_pay, net_pay)
        VALUES (:period_id, :employee_id, :scheduled_days, :work_days,
                :is_new_hire, :is_resigned, :base_salary, :overtime_pay, :gross_pay, :net_pay)
    """, data)
    conn.commit()
    conn.close()


def update_payroll_entry(entry_id, data):
    conn = get_db()
    conn.execute("""
        UPDATE payroll_entries SET
            work_days=:work_days, is_new_hire=:is_new_hire, is_resigned=:is_resigned,
            base_salary=:base_salary, overtime_pay=:overtime_pay,
            gross_pay=:gross_pay, notes=:notes
        WHERE id=:id
    """, {**data, 'id': entry_id})
    conn.commit()
    conn.close()


def recalculate_entry_totals(entry_id):
    """수당 합산 후 총지급/실지급 재계산"""
    conn = get_db()
    allowance_total = conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM payroll_allowance_entries WHERE entry_id=? AND is_paid=1",
        (entry_id,)
    ).fetchone()[0]
    deduction_total = conn.execute(
        "SELECT COALESCE(SUM(amount),0) FROM payroll_deduction_entries WHERE entry_id=?",
        (entry_id,)
    ).fetchone()[0]
    base = conn.execute(
        "SELECT base_salary, overtime_pay FROM payroll_entries WHERE id=?", (entry_id,)
    ).fetchone()
    gross = base['base_salary'] + base['overtime_pay'] + allowance_total
    net   = gross - deduction_total
    conn.execute("""
        UPDATE payroll_entries
        SET total_allowance=?, gross_pay=?, total_deduction=?, net_pay=?
        WHERE id=?
    """, (allowance_total, gross, deduction_total, net, entry_id))
    conn.commit()
    conn.close()


# ─── 수당 지급 내역 ──────────────────────────────────────────

def get_payroll_allowances_by_period(period_id):
    """{ entry_id: { allowance_item_id: {amount, is_paid, notes} } }"""
    conn = get_db()
    rows = conn.execute("""
        SELECT pae.entry_id, pae.allowance_item_id, pae.amount, pae.is_paid, pae.notes
        FROM payroll_allowance_entries pae
        JOIN payroll_entries pe ON pae.entry_id = pe.id
        WHERE pe.period_id = ?
    """, (period_id,)).fetchall()
    conn.close()
    result = {}
    for r in rows:
        result.setdefault(r['entry_id'], {})[r['allowance_item_id']] = {
            'amount': r['amount'], 'is_paid': r['is_paid'], 'notes': r['notes'] or ''
        }
    return result


def upsert_payroll_allowance(entry_id, allowance_item_id, amount, is_paid=1, notes=None):
    conn = get_db()
    conn.execute("""
        INSERT INTO payroll_allowance_entries (entry_id, allowance_item_id, amount, is_paid, notes)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(entry_id, allowance_item_id)
        DO UPDATE SET amount=excluded.amount, is_paid=excluded.is_paid,
                      notes=CASE WHEN excluded.notes IS NOT NULL THEN excluded.notes ELSE notes END
    """, (entry_id, allowance_item_id, amount, is_paid, notes))
    conn.commit()
    conn.close()


def get_employee_allowances_for_item(emp_ids, allowance_item_id):
    """특정 수당에 대한 직원별 기본금액 {employee_id: amount} — 일괄 조회"""
    if not emp_ids:
        return {}
    conn = get_db()
    placeholders = ','.join('?' * len(emp_ids))
    rows = conn.execute(
        f"SELECT employee_id, amount FROM employee_allowances "
        f"WHERE employee_id IN ({placeholders}) AND allowance_item_id=?",
        (*emp_ids, allowance_item_id)
    ).fetchall()
    conn.close()
    return {r['employee_id']: r['amount'] for r in rows}


# ─── 공제 내역 ──────────────────────────────────────────────

def get_payroll_deductions_by_period(period_id):
    conn = get_db()
    rows = conn.execute("""
        SELECT pde.*
        FROM payroll_deduction_entries pde
        JOIN payroll_entries pe ON pde.entry_id = pe.id
        WHERE pe.period_id = ?
        ORDER BY pde.deduction_name
    """, (period_id,)).fetchall()
    conn.close()
    return rows


def upsert_payroll_deduction(entry_id, deduction_name, amount):
    conn = get_db()
    existing = conn.execute(
        "SELECT id FROM payroll_deduction_entries WHERE entry_id=? AND deduction_name=?",
        (entry_id, deduction_name)
    ).fetchone()
    if existing:
        conn.execute("UPDATE payroll_deduction_entries SET amount=? WHERE id=?",
                     (amount, existing['id']))
    else:
        conn.execute(
            "INSERT INTO payroll_deduction_entries (entry_id, deduction_name, amount) VALUES (?,?,?)",
            (entry_id, deduction_name, amount)
        )
    conn.commit()
    conn.close()


def delete_payroll_deductions_by_entry(entry_id):
    conn = get_db()
    conn.execute("DELETE FROM payroll_deduction_entries WHERE entry_id=?", (entry_id,))
    conn.commit()
    conn.close()


def delete_payroll_deduction_by_name(entry_id, deduction_name):
    conn = get_db()
    conn.execute(
        "DELETE FROM payroll_deduction_entries WHERE entry_id=? AND deduction_name=?",
        (entry_id, deduction_name)
    )
    conn.commit()
    conn.close()
