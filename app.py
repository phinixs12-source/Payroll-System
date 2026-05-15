from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
import database as db
from utils import generate_code_from_korean, generate_code_from_english
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import os

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
                # 신규 등록 (기본값으로)
                db.create_employee({
                    'name_code': code,
                    'department': '',
                    'contract_type': '정규직',
                    'join_date': None,
                    'resign_date': None,
                    'severance_type': 'DB',
                    'scheduled_hours': 8.0,
                    'overtime_hours': 0.0,
                    'base_salary': salary,
                    'match_key': '',
                    'research_tax_exempt': 'N',
                    'child_deduction': 'N',
                    'child_count': 0,
                    'is_exception': 'N',
                    'employee_type': '직원',
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
        flash('인원 정보가 수정되었습니다.', 'success')
        return redirect(url_for('employee_list'))
    return render_template('employees/form.html', employee=employee,
                           allowances=allowances, emp_allowances=emp_allowances,
                           mode='edit')


@app.route('/employees/<int:emp_id>/delete', methods=['POST'])
def employee_delete(emp_id):
    emp = db.get_employee(emp_id)
    if emp:
        db.delete_employee(emp_id)
        flash(f'인원이 삭제되었습니다. [{emp["name_code"]}]', 'success')
    return redirect(url_for('employee_list'))


def _parse_employee_form(form):
    child_ded = 'Y' if form.get('child_deduction') == 'Y' else 'N'
    return {
        'name_code':            form.get('name_code', '').strip(),
        'department':           form.get('department', '').strip(),
        'contract_type':        form.get('contract_type', '정규직'),
        'join_date':            form.get('join_date', '') or None,
        'resign_date':          form.get('resign_date', '') or None,
        'severance_type':       form.get('severance_type', 'DB'),
        'scheduled_hours':      float(form.get('scheduled_hours', 8) or 8),
        'overtime_hours':       float(form.get('overtime_hours', 0) or 0),
        'base_salary':          int(form.get('base_salary', '0').replace(',', '') or 0),
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
    """인원관리 일괄업로드 템플릿"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '인원정보'
    headers = [
        '암호화코드', '부서/팀', '계약유형', '입사일(YYYY-MM-DD)',
        '퇴사일(없으면공백)', '퇴직금유형', '소정근로시간', '연장근로시간',
        '매칭키(주민번호앞7자리)', '연구보조비비과세(Y/N)',
        '자녀공제대상(Y/N)', '자녀수(0-3)', '예외자여부(Y/N)', '직원임원구분'
    ]
    ws.append(headers)
    _header_style(ws, 1, range(1, len(headers) + 1))
    # 예시 행
    sample = [
        'h4g3d411', '영업팀', '정규직', '2020-01-15',
        '', 'DB', '8', '0',
        '8501012', 'N',
        'N', '0', 'N', '직원'
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
    file = request.files.get('file')
    if not file or not file.filename.endswith(('.xlsx', '.xls')):
        flash('엑셀 파일(.xlsx)을 업로드해주세요.', 'danger')
        return redirect(url_for('employee_bulk_upload'))

    try:
        wb = openpyxl.load_workbook(file)
        ws = wb.active
        rows = list(ws.iter_rows(min_row=2, values_only=True))

        ok, errors = 0, []

        def _val(v, default=''):
            return str(v).strip() if v is not None else default

        def _yn(v):
            return 'Y' if str(v).strip().upper() == 'Y' else 'N'

        for i, row in enumerate(rows, start=2):
            if not row[0]:
                continue
            try:
                code = _val(row[0])
                child_ded = _yn(row[10])
                data = {
                    'name_code':           code,
                    'department':          _val(row[1]),
                    'contract_type':       _val(row[2], '정규직'),
                    'join_date':           _val(row[3]) or None,
                    'resign_date':         _val(row[4]) or None,
                    'severance_type':      _val(row[5], 'DB'),
                    'scheduled_hours':     float(_val(row[6], '8') or 8),
                    'overtime_hours':      float(_val(row[7], '0') or 0),
                    'match_key':           _val(row[8]),
                    'research_tax_exempt': _yn(row[9]),
                    'child_deduction':     child_ded,
                    'child_count':         int(_val(row[11], '0') or 0) if child_ded == 'Y' else 0,
                    'is_exception':        _yn(row[12]),
                    'employee_type':       _val(row[13], '직원'),
                    'base_salary':         0,
                }
                db.upsert_employee(data)
                ok += 1
            except Exception as e:
                errors.append(f'{i}행 오류: {str(e)}')

        msg = f'{ok}명 등록/업데이트 완료.'
        if errors:
            msg += f' 오류 {len(errors)}건: ' + ' / '.join(errors[:5])
            flash(msg, 'warning')
        else:
            flash(msg, 'success')

    except Exception as e:
        flash(f'파일 처리 중 오류: {str(e)}', 'danger')

    return redirect(url_for('employee_list'))


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
#  급여작업 / 인건비 / 기타분석 (placeholder)
# ════════════════════════════════════════════════════════
@app.route('/payroll')
def payroll():
    return render_template('payroll/index.html')


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
