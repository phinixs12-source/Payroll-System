from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
import database as db
from utils import generate_code_from_korean, generate_code_from_english
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import os
import calendar as _cal
from datetime import date as _date

app = Flask(__name__)
app.secret_key = 'payroll-secret-key-2025'

with app.app_context():
    db.init_db()


# ─── 포맷 필터 ───────────────────────────────────────────
def fmt_number(value):
    try:
        return f"{int(value):,}"
    except (ValueError, TypeError):
        return value

app.jinja_env.filters['fmt_number'] = fmt_number


# ─── 엑셀 스타일 헬퍼 ────────────────────────────────────
def _header_style(ws, row, cols, fill_color='1a4f8a'):
    fill = PatternFill('solid', fgColor=fill_color)
    font = Font(bold=True, color='FFFFFF', size=10)
    align = Alignment(horizontal='center', vertical='center')
    thin = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col in cols:
        cell = ws.cell(row=row, column=col)
        cell.fill = fill
        cell.font = font
        cell.alignment = align
        cell.border = border


def _sample_style(ws, row, cols):
    fill = PatternFill('solid', fgColor='F0F4F8')
    font = Font(color='718096', size=10, italic=True)
    thin = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col in cols:
        cell = ws.cell(row=row, column=col)
        cell.fill = fill
        cell.font = font
        cell.border = border


# ════════════════════════════════════════════════════════
#  대시보드
# ════════════════════════════════════════════════════════
@app.route('/')
def index():
    employees = db.get_all_employees()
    allowances = db.get_all_allowances()
    active_count = sum(1 for e in employees if not e['resign_date'])
    resign_count = sum(1 for e in employees if e['resign_date'])
    return render_template('index.html',
                           total_emp=len(employees),
                           active_count=active_count,
                           resign_count=resign_count,
                           allowance_count=len(allowances))


# ════════════════════════════════════════════════════════
#  인원 등록 — Step 1: 한글 이름 암호화
# ════════════════════════════════════════════════════════
@app.route('/employees/register')
def employee_register():
    return render_template('employees/register.html')


@app.route('/employees/register/template/step1')
def download_step1_template():
    """Step1 양식 다운로드 (이름 입력용)"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '이름목록'
    headers = ['이름(한글)']
    ws.append(headers)
    _header_style(ws, 1, range(1, 2))
    ws.append(['홍길동'])
    ws.append(['홍길동a'])
    ws.append(['남궁민호'])
    _sample_style(ws, 2, [1])
    _sample_style(ws, 3, [1])
    _sample_style(ws, 4, [1])
    ws.column_dimensions['A'].width = 20
    ws['A1'].comment = None
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name='01_이름목록_양식.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/employees/register/step1', methods=['POST'])
def process_step1():
    """한글 이름 엑셀 업로드 → 암호화 → 결과 다운로드"""
    file = request.files.get('file')
    if not file or not file.filename.endswith(('.xlsx', '.xls')):
        flash('엑셀 파일(.xlsx)을 업로드해주세요.', 'danger')
        return redirect(url_for('employee_register'))

    try:
        wb_in = openpyxl.load_workbook(file)
        ws_in = wb_in.active
        rows = list(ws_in.iter_rows(min_row=2, values_only=True))
        names = [str(r[0]).strip() for r in rows if r[0] and str(r[0]).strip()]

        if not names:
            flash('이름 데이터가 없습니다.', 'warning')
            return redirect(url_for('employee_register'))

        # 암호화
        results = []
        seen_codes = {}
        for name in names:
            code = generate_code_from_korean(name)
            if code in seen_codes:
                flash(f'중복 코드 발생: {name} → {code} (기존: {seen_codes[code]})', 'warning')
            seen_codes[code] = name
            results.append((name, code))

        # 결과 엑셀 생성
        wb_out = openpyxl.Workbook()
        ws_out = wb_out.active
        ws_out.title = '암호화결과'
        ws_out.append(['원본이름(한글)', '암호화코드'])
        _header_style(ws_out, 1, [1, 2])
        for orig, code in results:
            ws_out.append([orig, code])
        ws_out.column_dimensions['A'].width = 20
        ws_out.column_dimensions['B'].width = 20

        buf = io.BytesIO()
        wb_out.save(buf)
        buf.seek(0)
        flash(f'{len(results)}명 암호화 완료. 파일을 다운로드하세요.', 'success')
        return send_file(buf, as_attachment=True,
                         download_name='02_암호화결과.xlsx',
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        flash(f'파일 처리 중 오류: {str(e)}', 'danger')
        return redirect(url_for('employee_register'))


# ════════════════════════════════════════════════════════
#  인원 등록 — Step 2: 연봉 등록
# ════════════════════════════════════════════════════════
@app.route('/employees/register/template/step2')
def download_step2_template():
    """Step2 양식 다운로드 (코드+연봉)"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '연봉등록'
    ws.append(['암호화코드', '연봉(원)'])
    _header_style(ws, 1, [1, 2])
    ws.append(['h4g3d411', 36000000])
    _sample_style(ws, 2, [1, 2])
    ws.column_dimensions['A'].width = 22
    ws.column_dimensions['B'].width = 18
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name='03_연봉등록_양식.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/employees/register/step2', methods=['POST'])
def process_step2():
    """암호화코드 + 연봉 업로드 → DB 등록"""
    file = request.files.get('file')
    if not file or not file.filename.endswith(('.xlsx', '.xls')):
        flash('엑셀 파일(.xlsx)을 업로드해주세요.', 'danger')
        return redirect(url_for('employee_register'))

    try:
        wb = openpyxl.load_workbook(file)
        ws = wb.active
        rows = list(ws.iter_rows(min_row=2, values_only=True))

        ok, skipped, errors = 0, 0, []
        for i, row in enumerate(rows, start=2):
            if not row[0]:
                continue
            code = str(row[0]).strip()
            try:
                salary = int(str(row[1]).replace(',', '').replace(' ', ''))
            except (ValueError, TypeError):
                errors.append(f'{i}행: 연봉 값 오류 ({row[1]})')
                continue

            existing = db.get_employee_by_code(code)
            if existing:
                db.update_employee_salary(code, salary)
                ok += 1
            else:
                # 신규 등록 (기본값으로, 연봉만 입력)
                db.create_employee({
                    'name_code':           code,
                    'department':          '',
                    'contract_type':       '정규직',
                    'join_date':           None,
                    'resign_date':         None,
                    'severance_type':      'DB',
                    'scheduled_hours':     209.0,
                    'overtime_hours':      0.0,
                    'annual_salary':       salary,
                    'base_salary':         0,
                    'ordinary_wage':       0,
                    'match_key':           '',
                    'research_tax_exempt': 'N',
                    'child_deduction':     'N',
                    'child_count':         0,
                    'is_exception':        'N',
                    'employee_type':       '직원',
                })
                ok += 1

        msg = f'{ok}명 연봉 등록/업데이트 완료.'
        if errors:
            msg += f' 오류 {len(errors)}건: ' + ' / '.join(errors[:3])
            flash(msg, 'warning')
        else:
            flash(msg, 'success')

    except Exception as e:
        flash(f'파일 처리 중 오류: {str(e)}', 'danger')

    return redirect(url_for('employee_register'))


# ════════════════════════════════════════════════════════
#  인원 관리 — 목록 / 개별 수정 / 삭제
# ════════════════════════════════════════════════════════
@app.route('/employees')
def employee_list():
    employees = db.get_all_employees()
    return render_template('employees/list.html', employees=employees)


@app.route('/employees/new', methods=['GET', 'POST'])
def employee_new():
    allowances = db.get_all_allowances()
    if request.method == 'POST':
        data = _parse_employee_form(request.form)
        if not data['name_code']:
            flash('이름 코드를 입력해주세요.', 'danger')
            return render_template('employees/form.html', employee=data,
                                   allowances=allowances, mode='new')
        emp_id = db.create_employee(data)
        for a in allowances:
            amt_str = request.form.get(f'allowance_{a["id"]}', '0').replace(',', '')
            db.upsert_employee_allowance(emp_id, a['id'], int(amt_str or 0))
        flash(f'인원이 등록되었습니다. [{data["name_code"]}]', 'success')
        return redirect(url_for('employee_list'))
    return render_template('employees/form.html', employee=None,
                           allowances=allowances, mode='new')


@app.route('/employees/<int:emp_id>/edit', methods=['GET', 'POST'])
def employee_edit(emp_id):
    employee = db.get_employee(emp_id)
    if not employee:
        flash('존재하지 않는 인원입니다.', 'danger')
        return redirect(url_for('employee_list'))
    allowances = db.get_all_allowances()
    emp_allowances = {ea['allowance_item_id']: ea['amount']
                      for ea in db.get_employee_allowances(emp_id)}
    if request.method == 'POST':
        data = _parse_employee_form(request.form)
        db.update_employee(emp_id, data)
        for a in allowances:
            amt_str = request.form.get(f'allowance_{a["id"]}', '0').replace(',', '')
            db.upsert_employee_allowance(emp_id, a['id'], int(amt_str or 0))
        # 통상임금 자동계산 후 저장
        _recalc_ordinary_wage(emp_id, allowances, emp_allowances)
        flash('인원 정보가 수정되었습니다.', 'success')
        return redirect(url_for('employee_list'))
    # 통상임금·연장근로수당 계산 (GET용)
    ordinary_wage_calc = _calc_ordinary_wage(employee, allowances, emp_allowances)
    monthly_pay = round(employee['annual_salary'] / 12) if employee['annual_salary'] else 0
    overtime_pay_calc = monthly_pay - (employee['base_salary'] or 0)
    return render_template('employees/form.html', employee=employee,
                           allowances=allowances, emp_allowances=emp_allowances,
                           ordinary_wage_calc=ordinary_wage_calc,
                           overtime_pay_calc=max(0, overtime_pay_calc),
                           mode='edit')


@app.route('/employees/<int:emp_id>/delete', methods=['POST'])
def employee_delete(emp_id):
    emp = db.get_employee(emp_id)
    if emp:
        db.delete_employee(emp_id)
        flash(f'인원이 삭제되었습니다. [{emp["name_code"]}]', 'success')
    return redirect(url_for('employee_list'))


@app.route('/employees/bulk-delete', methods=['POST'])
def employee_bulk_delete():
    ids = request.form.getlist('delete_ids')
    count = 0
    for eid in ids:
        try:
            emp = db.get_employee(int(eid))
            if emp:
                db.delete_employee(int(eid))
                count += 1
        except Exception:
            pass
    flash(f'{count}명이 삭제되었습니다.', 'success' if count else 'warning')
    return redirect(url_for('employee_list'))


def _calc_ordinary_wage(employee, allowances, emp_allowances):
    """통상임금 = 기본급 + is_ordinary_wage=1 수당 합계"""
    total = employee['base_salary'] or 0
    for a in allowances:
        if a['is_ordinary_wage'] and a['id'] in emp_allowances:
            total += emp_allowances[a['id']]
    return total


def _recalc_ordinary_wage(emp_id, allowances, emp_allowances):
    """편집 완료 후 통상임금을 DB에 저장"""
    emp = db.get_employee(emp_id)
    if not emp:
        return
    # 최신 수당 다시 읽기
    fresh_ea = {ea['allowance_item_id']: ea['amount']
                for ea in db.get_employee_allowances(emp_id)}
    ow = emp['base_salary'] or 0
    for a in allowances:
        if a['is_ordinary_wage'] and a['id'] in fresh_ea:
            ow += fresh_ea[a['id']]
    import sqlite3 as _s3
    conn = db.get_db()
    conn.execute("UPDATE employees SET ordinary_wage=? WHERE id=?", (ow, emp_id))
    conn.commit()
    conn.close()


def _parse_employee_form(form):
    child_ded = 'Y' if form.get('child_deduction') == 'Y' else 'N'
    annual_salary    = int(form.get('annual_salary', '0').replace(',', '') or 0)
    scheduled_hours  = float(form.get('scheduled_hours', 209) or 209)
    overtime_hours   = float(form.get('overtime_hours', 0) or 0)
    ordinary_wage    = int(form.get('ordinary_wage', '0').replace(',', '') or 0)
    base_salary_raw  = int(form.get('base_salary', '0').replace(',', '') or 0)

    # 기본급 자동계산: (연봉/12) / (소정+연장) × 소정
    total_hours = scheduled_hours + overtime_hours
    if annual_salary > 0 and total_hours > 0:
        base_salary = int((annual_salary / 12) / total_hours * scheduled_hours)
    else:
        base_salary = base_salary_raw

    return {
        'name_code':            form.get('name_code', '').strip(),
        'department':           form.get('department', '').strip(),
        'contract_type':        form.get('contract_type', '정규직'),
        'join_date':            form.get('join_date', '') or None,
        'resign_date':          form.get('resign_date', '') or None,
        'severance_type':       form.get('severance_type', 'DB'),
        'scheduled_hours':      scheduled_hours,
        'overtime_hours':       overtime_hours,
        'annual_salary':        annual_salary,
        'base_salary':          base_salary,
        'ordinary_wage':        ordinary_wage,
        'match_key':            form.get('match_key', '').strip(),
        'research_tax_exempt':  form.get('research_tax_exempt', 'N'),
        'child_deduction':      child_ded,
        'child_count':          int(form.get('child_count', 0) or 0) if child_ded == 'Y' else 0,
        'is_exception':         form.get('is_exception', 'N'),
        'employee_type':        form.get('employee_type', '직원'),
    }


# ════════════════════════════════════════════════════════
#  인원 관리 — 일괄 업로드
# ════════════════════════════════════════════════════════
@app.route('/employees/bulk-upload')
def employee_bulk_upload():
    return render_template('employees/bulk_upload.html')


@app.route('/employees/bulk-upload/template')
def download_bulk_template():
    """인원관리 일괄업로드 템플릿 (기등록 인원 pre-populate)"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '인원정보'
    headers = [
        '암호화코드', '부서/팀', '계약유형', '입사일(YYYY-MM-DD)',
        '퇴사일(없으면공백)', '퇴직금유형', '소정근로시간(월)', '연장근로시간(월)',
        '매칭키(주민번호앞7자리)', '연구보조비비과세(Y/N)',
        '자녀공제대상(Y/N)', '자녀수(0-3)', '예외자여부(Y/N)', '직원임원구분'
    ]
    ws.append(headers)
    _header_style(ws, 1, range(1, len(headers) + 1))

    employees = db.get_all_employees()
    thin = Side(style='thin', color='CCCCCC')
    data_border = Border(left=thin, right=thin, top=thin, bottom=thin)

    if employees:
        for row_idx, emp in enumerate(employees, start=2):
            # 날짜: YYYY-MM-DD 앞 10자리만
            join_d  = (str(emp['join_date'])[:10]  if emp['join_date']  else '')
            resign_d = (str(emp['resign_date'])[:10] if emp['resign_date'] else '')
            row_data = [
                emp['name_code'],
                emp['department']        or '',
                emp['contract_type']     or '정규직',
                join_d,
                resign_d,
                emp['severance_type']    or 'DB',
                emp['scheduled_hours']   if emp['scheduled_hours'] is not None else 209,
                emp['overtime_hours']    if emp['overtime_hours']  is not None else 0,
                emp['match_key']         or '',
                emp['research_tax_exempt'] or 'N',
                emp['child_deduction']   or 'N',
                emp['child_count']       or 0,
                emp['is_exception']      or 'N',
                emp['employee_type']     or '직원',
            ]
            ws.append(row_data)
            for col in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=col).border = data_border
    else:
        # 등록된 인원 없을 때 예시 행
        sample = [
            'h4g3d411', '영업팀', '정규직', '2020-01-15',
            '', 'DB', '209', '0',
            '8501012', 'N', 'N', '0', 'N', '직원'
        ]
        ws.append(sample)
        _sample_style(ws, 2, range(1, len(headers) + 1))

    col_widths = [18, 14, 12, 20, 16, 12, 14, 14, 22, 20, 16, 12, 14, 14]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name='04_인원정보_일괄업로드_양식.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/employees/bulk-upload/process', methods=['POST'])
def process_bulk_upload():
    import re as _re
    file = request.files.get('file')
    if not file or not file.filename.endswith(('.xlsx', '.xls')):
        flash('엑셀 파일(.xlsx)을 업로드해주세요.', 'danger')
        return redirect(url_for('employee_bulk_upload'))

    VALID_CONTRACT      = {'정규직', '계약직', '파트타임'}
    VALID_SEVERANCE     = {'DB', 'DC', '1년미만'}
    VALID_EMP_TYPE      = {'직원', '임원'}
    DATE_RE             = _re.compile(r'^\d{4}-\d{2}-\d{2}$')

    def _val(v, default=''):
        return str(v).strip() if v is not None else default

    def _yn(v):
        return 'Y' if str(v).strip().upper() == 'Y' else 'N'

    def _date(v):
        """값이 있으면 앞 10자리만, 없으면 None. 형식 오류 시 (None, 오류메시지) 반환"""
        s = _val(v)
        if not s:
            return None, None
        s = s[:10]
        if not DATE_RE.match(s):
            return None, f'날짜 형식 오류: {v} → YYYY-MM-DD 필요'
        return s, None

    try:
        wb = openpyxl.load_workbook(file)
        ws = wb.active
        rows = list(ws.iter_rows(min_row=2, values_only=True))

        ok, error_rows = 0, []

        for i, row in enumerate(rows, start=2):
            if not row[0]:
                continue

            code = _val(row[0])
            row_errors = []

            # ── 계약유형
            contract_type = _val(row[2], '정규직')
            if contract_type not in VALID_CONTRACT:
                row_errors.append(f'계약유형 오류 "{contract_type}" → 정규직/계약직/파트타임')

            # ── 퇴직금유형
            severance_type = _val(row[5], 'DB')
            if severance_type not in VALID_SEVERANCE:
                row_errors.append(f'퇴직금유형 오류 "{severance_type}" → DB/DC/1년미만')

            # ── 직원임원구분
            employee_type = _val(row[13], '직원')
            if employee_type not in VALID_EMP_TYPE:
                row_errors.append(f'직원임원구분 오류 "{employee_type}" → 직원/임원')

            # ── 날짜
            join_date,   e1 = _date(row[3])
            resign_date, e2 = _date(row[4])
            if e1: row_errors.append(f'입사일 {e1}')
            if e2: row_errors.append(f'퇴사일 {e2}')

            # ── 근로시간
            try:
                scheduled_hours = float(_val(row[6], '209') or 209)
                if scheduled_hours <= 0:
                    row_errors.append('소정근로시간은 0보다 커야 합니다')
            except ValueError:
                row_errors.append(f'소정근로시간 숫자 오류: {row[6]}')
                scheduled_hours = 209

            try:
                overtime_hours = float(_val(row[7], '0') or 0)
            except ValueError:
                row_errors.append(f'연장근로시간 숫자 오류: {row[7]}')
                overtime_hours = 0

            # ── 자녀공제
            child_ded = _yn(row[10])
            try:
                child_count = int(_val(row[11], '0') or 0)
                if not (0 <= child_count <= 3):
                    row_errors.append(f'자녀수 오류 {child_count} → 0~3')
                    child_count = 0
            except ValueError:
                row_errors.append(f'자녀수 숫자 오류: {row[11]}')
                child_count = 0
            if child_ded == 'Y' and child_count == 0:
                row_errors.append('자녀공제 대상(Y)인데 자녀수가 0')

            # ── 오류 있으면 저장하지 않고 기록
            if row_errors:
                error_rows.append({'row': i, 'code': code, 'errors': row_errors})
                continue

            # ── 기존 인원 salary 보존
            existing = db.get_employee_by_code(code)
            try:
                data = {
                    'name_code':           code,
                    'department':          _val(row[1]),
                    'contract_type':       contract_type,
                    'join_date':           join_date,
                    'resign_date':         resign_date,
                    'severance_type':      severance_type,
                    'scheduled_hours':     scheduled_hours,
                    'overtime_hours':      overtime_hours,
                    'annual_salary':       existing['annual_salary'] if existing else 0,
                    'base_salary':         existing['base_salary']   if existing else 0,
                    'ordinary_wage':       existing['ordinary_wage'] if existing else 0,
                    'match_key':           _val(row[8]),
                    'research_tax_exempt': _yn(row[9]),
                    'child_deduction':     child_ded,
                    'child_count':         child_count if child_ded == 'Y' else 0,
                    'is_exception':        _yn(row[12]),
                    'employee_type':       employee_type,
                }
                db.upsert_employee(data)
                ok += 1
            except Exception as e:
                error_rows.append({'row': i, 'code': code, 'errors': [str(e)]})

        if error_rows:
            flash(f'{ok}명 등록/업데이트 완료. {len(error_rows)}행에서 오류가 발생했습니다.', 'warning')
            return render_template('employees/bulk_upload.html',
                                   error_rows=error_rows, ok_count=ok)

        flash(f'{ok}명 등록/업데이트 완료.', 'success')
        return redirect(url_for('employee_list'))

    except Exception as e:
        flash(f'파일 처리 중 오류: {str(e)}', 'danger')
        return redirect(url_for('employee_bulk_upload'))


# ════════════════════════════════════════════════════════
#  수당 항목 관리
# ════════════════════════════════════════════════════════
@app.route('/allowances')
def allowance_list():
    allowances = db.get_all_allowances()
    return render_template('allowances/list.html', allowances=allowances)


@app.route('/allowances/new', methods=['GET', 'POST'])
def allowance_new():
    if request.method == 'POST':
        data = _parse_allowance_form(request.form)
        if not data['name']:
            flash('수당명을 입력해주세요.', 'danger')
            return render_template('allowances/form.html', item=data, mode='new')
        db.create_allowance(data)
        flash(f'수당 항목이 추가되었습니다. [{data["name"]}]', 'success')
        return redirect(url_for('allowance_list'))
    return render_template('allowances/form.html', item=None, mode='new')


@app.route('/allowances/<int:item_id>/edit', methods=['GET', 'POST'])
def allowance_edit(item_id):
    item = db.get_allowance(item_id)
    if not item:
        flash('존재하지 않는 수당 항목입니다.', 'danger')
        return redirect(url_for('allowance_list'))
    if request.method == 'POST':
        data = _parse_allowance_form(request.form)
        db.update_allowance(item_id, data)
        flash(f'수당 항목이 수정되었습니다. [{data["name"]}]', 'success')
        return redirect(url_for('allowance_list'))
    return render_template('allowances/form.html', item=item, mode='edit')


@app.route('/allowances/<int:item_id>/delete', methods=['POST'])
def allowance_delete(item_id):
    item = db.get_allowance(item_id)
    if item:
        db.delete_allowance(item_id)
        flash(f'수당 항목이 삭제되었습니다. [{item["name"]}]', 'success')
    return redirect(url_for('allowance_list'))


@app.route('/allowances/<int:item_id>/toggle', methods=['POST'])
def allowance_toggle(item_id):
    db.toggle_allowance(item_id)
    return redirect(url_for('allowance_list'))


def _parse_allowance_form(form):
    return {
        'name':              form.get('name', '').strip(),
        'is_ordinary_wage':  1 if form.get('is_ordinary_wage') else 0,
        'severance_db':      1 if form.get('severance_db') else 0,
        'severance_dc':      1 if form.get('severance_dc') else 0,
        'payment_condition': form.get('payment_condition', 'fixed'),
        'condition_value':   int(form.get('condition_value', 0) or 0),
        'apply_target':      form.get('apply_target', 'all'),
        'is_active':         1 if form.get('is_active') else 0,
    }


# ════════════════════════════════════════════════════════
#  급여작업
# ════════════════════════════════════════════════════════

def _get_active_employees_for_period(year, month):
    """해당 귀속월에 포함될 직원 목록 (예외자 제외)"""
    total_days = _cal.monthrange(year, month)[1]
    month_start = f'{year:04d}-{month:02d}-01'
    month_end   = f'{year:04d}-{month:02d}-{total_days:02d}'
    result = []
    for emp in db.get_all_employees():
        if emp['is_exception'] == 'Y':
            continue
        resign_d = (emp['resign_date'] or '')[:10]
        join_d   = (emp['join_date']   or '')[:10]
        if resign_d and resign_d < month_start:
            continue   # 이미 전월 이전 퇴사
        if join_d and join_d > month_end:
            continue   # 아직 입사 전
        result.append(emp)
    return result


def _build_payroll_entries(period_id, year, month, employees):
    """직원 목록으로 급여대장 기본 항목 생성"""
    total_days = _cal.monthrange(year, month)[1]
    ym = f'{year:04d}-{month:02d}'
    for emp in employees:
        monthly  = round((emp['annual_salary'] or 0) / 12)
        base     = emp['base_salary'] or 0
        overtime = max(0, monthly - base)

        join_d   = (emp['join_date']   or '')[:10]
        resign_d = (emp['resign_date'] or '')[:10]
        is_new   = 1 if join_d[:7]   == ym else 0
        is_res   = 1 if resign_d[:7] == ym else 0

        if is_new:
            work_days = total_days - int(join_d[8:10]) + 1
        elif is_res:
            work_days = int(resign_d[8:10])
        else:
            work_days = total_days

        if is_new or is_res:
            ratio    = work_days / total_days
            base     = round(base * ratio)
            overtime = round(overtime * ratio)

        gross = base + overtime
        db.create_payroll_entry({
            'period_id':     period_id,
            'employee_id':   emp['id'],
            'scheduled_days': total_days,
            'work_days':     work_days,
            'is_new_hire':   is_new,
            'is_resigned':   is_res,
            'base_salary':   base,
            'overtime_pay':  overtime,
            'gross_pay':     gross,
            'net_pay':       gross,
        })


@app.route('/payroll')
def payroll():
    periods = db.get_all_payroll_periods()
    return render_template('payroll/index.html', periods=periods)


@app.route('/payroll/new', methods=['GET', 'POST'])
def payroll_new():
    today = _date.today()
    if request.method == 'POST':
        year  = int(request.form['year'])
        month = int(request.form['month'])
        payment_date = request.form.get('payment_date', '')
        mode  = request.form.get('mode', 'new')

        if db.get_payroll_period_by_ym(year, month):
            flash(f'{year}년 {month}월 귀속 급여대장이 이미 존재합니다.', 'warning')
            return redirect(url_for('payroll'))

        period_id = db.create_payroll_period(year, month, payment_date)

        if mode == 'copy':
            prev_y = year if month > 1 else year - 1
            prev_m = month - 1 if month > 1 else 12
            prev   = db.get_payroll_period_by_ym(prev_y, prev_m)
            if prev:
                prev_entries = db.get_payroll_entries(prev['id'])
                emp_ids = [e['employee_id'] for e in prev_entries]
                employees = [db.get_employee(eid) for eid in emp_ids]
                employees = [e for e in employees if e and
                             e['is_exception'] != 'Y' and
                             ((e['resign_date'] or '') == '' or
                              e['resign_date'][:7] >= f'{year:04d}-{month:02d}')]
            else:
                flash('이전 귀속기간이 없어 신규로 생성합니다.', 'info')
                employees = _get_active_employees_for_period(year, month)
        else:
            employees = _get_active_employees_for_period(year, month)

        _build_payroll_entries(period_id, year, month, employees)
        db.update_payroll_period_status(period_id, 'step1')
        flash(f'{year}년 {month}월 급여대장 생성 완료 (총 {len(employees)}명)', 'success')
        return redirect(url_for('payroll_detail', period_id=period_id))

    prev_y = today.year if today.month > 1 else today.year - 1
    prev_m = today.month - 1 if today.month > 1 else 12
    has_prev = bool(db.get_payroll_period_by_ym(prev_y, prev_m))
    return render_template('payroll/new.html', today=today, has_prev=has_prev)


@app.route('/payroll/<int:period_id>')
def payroll_detail(period_id):
    period = db.get_payroll_period(period_id)
    if not period:
        flash('급여대장을 찾을 수 없습니다.', 'danger')
        return redirect(url_for('payroll'))
    entries    = db.get_payroll_entries(period_id)
    allowances = db.get_all_allowances()
    paid_map   = db.get_payroll_allowances_by_period(period_id)
    ded_rows   = db.get_payroll_deductions_by_period(period_id)
    ded_names  = list(dict.fromkeys([d['deduction_name'] for d in ded_rows]))
    ded_map    = {}
    for d in ded_rows:
        ded_map.setdefault(d['entry_id'], {})[d['deduction_name']] = d['amount']
    return render_template('payroll/detail.html',
                           period=period, entries=entries,
                           allowances=allowances, paid_map=paid_map,
                           ded_names=ded_names, ded_map=ded_map)


@app.route('/payroll/<int:period_id>/step2', methods=['GET', 'POST'])
def payroll_step2(period_id):
    period = db.get_payroll_period(period_id)
    if not period:
        return redirect(url_for('payroll'))
    entries     = db.get_payroll_entries(period_id)
    total_days  = _cal.monthrange(period['year'], period['month'])[1]
    targets     = [e for e in entries if e['is_new_hire'] or e['is_resigned']]

    if request.method == 'POST':
        # 건너뛰기 (입퇴사자 없음)
        if request.form.get('_skip'):
            db.update_payroll_period_status(period_id, 'step2')
            return redirect(url_for('payroll_step3', period_id=period_id))

        for e in targets:
            wd = int(request.form.get(f'work_days_{e["id"]}', e['work_days']))
            wd = max(1, min(wd, total_days))
            emp  = db.get_employee(e['employee_id'])
            monthly  = round((emp['annual_salary'] or 0) / 12)
            base_full = emp['base_salary'] or 0
            ot_full   = max(0, monthly - base_full)
            ratio = wd / total_days
            base_p = round(base_full * ratio)
            ot_p   = round(ot_full   * ratio)
            db.update_payroll_entry(e['id'], {
                'work_days':   wd,
                'is_new_hire': e['is_new_hire'],
                'is_resigned': e['is_resigned'],
                'base_salary': base_p,
                'overtime_pay': ot_p,
                'gross_pay':   base_p + ot_p,
                'notes':       request.form.get(f'notes_{e["id"]}', ''),
            })
        db.update_payroll_period_status(period_id, 'step2')
        flash('입퇴사자 일할계산이 반영되었습니다.', 'success')
        return redirect(url_for('payroll_detail', period_id=period_id))

    return render_template('payroll/step2.html',
                           period=period, targets=targets, total_days=total_days)


@app.route('/payroll/<int:period_id>/step3', methods=['GET', 'POST'])
def payroll_step3(period_id):
    period     = db.get_payroll_period(period_id)
    if not period:
        return redirect(url_for('payroll'))
    entries    = db.get_payroll_entries(period_id)
    allowances = [a for a in db.get_all_allowances() if a['is_active']]
    total_days = _cal.monthrange(period['year'], period['month'])[1]

    # 전월 수당 데이터
    prev_y = period['year'] if period['month'] > 1 else period['year'] - 1
    prev_m = period['month'] - 1 if period['month'] > 1 else 12
    prev   = db.get_payroll_period_by_ym(prev_y, prev_m)
    prev_paid_map = db.get_payroll_allowances_by_period(prev['id']) if prev else {}

    if request.method == 'POST':
        for entry in entries:
            for a in allowances:
                amt_str = request.form.get(f'amt_{entry["id"]}_{a["id"]}', '0').replace(',', '')
                amount  = int(amt_str or 0)
                is_paid = 1 if request.form.get(f'paid_{entry["id"]}_{a["id"]}') else 0
                db.upsert_payroll_allowance(entry['id'], a['id'], amount, is_paid)
        for entry in entries:
            db.recalculate_entry_totals(entry['id'])
        db.update_payroll_period_status(period_id, 'step3')
        flash('수당 처리가 완료되었습니다.', 'success')
        return redirect(url_for('payroll_detail', period_id=period_id))

    # GET — 기본값 준비 (등록 수당 → 없으면 employee_allowances 기본값)
    current_paid_map = db.get_payroll_allowances_by_period(period_id)
    load_prev = request.args.get('load_prev') == '1'

    amounts = {}  # amounts[entry_id][allowance_id] = {amount, is_paid, flag}
    for entry in entries:
        emp_ea = {ea['allowance_item_id']: ea['amount']
                  for ea in db.get_employee_allowances(entry['employee_id'])}
        amounts[entry['id']] = {}
        for a in allowances:
            # 우선순위: 현재저장 > 전월불러오기 > 직원기본값
            if load_prev and prev_paid_map:
                # 전월 데이터에서 같은 직원 entry 찾기
                prev_entry_data = _find_prev_entry_allowance(
                    prev_paid_map, prev['id'] if prev else None,
                    entry['employee_id'], a['id']
                ) if prev else None
                if prev_entry_data:
                    val = prev_entry_data
                else:
                    val = {'amount': emp_ea.get(a['id'], 0), 'is_paid': 1}
            elif entry['id'] in current_paid_map and a['id'] in current_paid_map[entry['id']]:
                val = current_paid_map[entry['id']][a['id']]
            else:
                val = {'amount': emp_ea.get(a['id'], 0), 'is_paid': 1}

            # 15일 조건 자동 플래그
            flag = False
            if (entry['is_new_hire'] or entry['is_resigned']) and \
               a['payment_condition'] == 'min_days' and \
               entry['work_days'] < a['condition_value']:
                val = dict(val)
                val['is_paid'] = 0
                flag = True
            val['flag'] = flag
            amounts[entry['id']][a['id']] = val

    return render_template('payroll/step3.html',
                           period=period, entries=entries,
                           allowances=allowances, amounts=amounts,
                           total_days=total_days,
                           has_prev=bool(prev),
                           load_prev=load_prev)


def _find_prev_entry_allowance(prev_paid_map, prev_period_id, employee_id, allowance_id):
    """전월 급여대장에서 같은 직원의 수당 금액 찾기"""
    if not prev_period_id:
        return None
    conn = db.get_db()
    row = conn.execute(
        "SELECT id FROM payroll_entries WHERE period_id=? AND employee_id=?",
        (prev_period_id, employee_id)
    ).fetchone()
    conn.close()
    if not row:
        return None
    entry_id = row['id']
    return prev_paid_map.get(entry_id, {}).get(allowance_id)


@app.route('/payroll/<int:period_id>/step4', methods=['GET', 'POST'])
def payroll_step4(period_id):
    period  = db.get_payroll_period(period_id)
    if not period:
        return redirect(url_for('payroll'))
    entries = db.get_payroll_entries(period_id)

    ded_rows  = db.get_payroll_deductions_by_period(period_id)
    ded_names = list(dict.fromkeys([d['deduction_name'] for d in ded_rows]))
    if not ded_names:
        ded_names = ['건강보험', '장기요양', '국민연금', '고용보험', '소득세', '지방소득세']
    ded_map = {}
    for d in ded_rows:
        ded_map.setdefault(d['entry_id'], {})[d['deduction_name']] = d['amount']

    if request.method == 'POST':
        # 공제 항목명 목록
        names = [n.strip() for n in request.form.get('ded_names_list', '').split(',') if n.strip()]
        for entry in entries:
            db.delete_payroll_deductions_by_entry(entry['id'])
            for dname in names:
                key = f'ded_{entry["id"]}_{dname}'
                amt = int(request.form.get(key, '0').replace(',', '') or 0)
                if amt > 0:
                    db.upsert_payroll_deduction(entry['id'], dname, amt)
            db.recalculate_entry_totals(entry['id'])
        db.update_payroll_period_status(period_id, 'completed')
        flash('공제 처리 완료. 급여대장이 확정되었습니다.', 'success')
        return redirect(url_for('payroll_detail', period_id=period_id))

    return render_template('payroll/step4.html',
                           period=period, entries=entries,
                           ded_names=ded_names, ded_map=ded_map)


@app.route('/payroll/<int:period_id>/delete', methods=['POST'])
def payroll_period_delete(period_id):
    period = db.get_payroll_period(period_id)
    if period:
        db.delete_payroll_period(period_id)
        flash(f'{period["year"]}년 {period["month"]}월 급여대장이 삭제되었습니다.', 'success')
    return redirect(url_for('payroll'))


@app.route('/labor-cost')
def labor_cost():
    return render_template('labor_cost/index.html')


@app.route('/analysis')
def analysis():
    return render_template('analysis/index.html')


# ════════════════════════════════════════════════════════
#  설정
# ════════════════════════════════════════════════════════
@app.route('/settings', methods=['GET', 'POST'])
def settings():
    if request.method == 'POST':
        db.set_setting('company_name', request.form.get('company_name', ''))
        db.set_setting('abnormal_threshold', request.form.get('abnormal_threshold', '20'))
        flash('설정이 저장되었습니다.', 'success')
        return redirect(url_for('settings'))
    cfg = db.get_all_settings()
    return render_template('settings/index.html', cfg=cfg)


# ════════════════════════════════════════════════════════
if __name__ == '__main__':
    print("=" * 50)
    print("  인건비 관리 시스템 시작")
    print("  브라우저에서 http://localhost:5000 접속")
    print("=" * 50)
    app.run(debug=True, port=5000)
