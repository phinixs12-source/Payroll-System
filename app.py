from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
import database as db
import re

app = Flask(__name__)
app.secret_key = 'payroll-secret-key-2025'

# ─── 앱 시작 시 DB 초기화 ────────────────────────────────
with app.app_context():
    db.init_db()


# ─── 이름 코드 생성 유틸 ─────────────────────────────────
def generate_name_code(name: str) -> str:
    """
    Hong Gil Dong  →  h4g3d411
    각 단어: (첫글자 소문자) + (단어 철자 수)
    마지막: 전체 철자 수 합산
    """
    words = name.strip().split()
    code = ""
    total = 0
    for word in words:
        clean = re.sub(r'[^a-zA-Z]', '', word)
        if clean:
            code += clean[0].lower() + str(len(clean))
            total += len(clean)
    code += str(total)
    return code


def fmt_number(value):
    """숫자 천단위 콤마 포맷"""
    try:
        return f"{int(value):,}"
    except (ValueError, TypeError):
        return value

app.jinja_env.filters['fmt_number'] = fmt_number


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
#  이름 코드 생성
# ════════════════════════════════════════════════════════
@app.route('/employees/code-gen', methods=['GET', 'POST'])
def code_gen():
    result = None
    input_name = ''
    if request.method == 'POST':
        input_name = request.form.get('name', '').strip()
        if input_name:
            result = generate_name_code(input_name)
        else:
            flash('이름을 입력해주세요.', 'warning')
    return render_template('employees/code_gen.html',
                           result=result, input_name=input_name)


@app.route('/api/code-gen')
def api_code_gen():
    name = request.args.get('name', '')
    if not name:
        return jsonify({'error': '이름을 입력해주세요'}), 400
    return jsonify({'code': generate_name_code(name)})


# ════════════════════════════════════════════════════════
#  인원 관리
# ════════════════════════════════════════════════════════
@app.route('/employees')
def employee_list():
    employees = db.get_all_employees()
    return render_template('employees/list.html', employees=employees)


@app.route('/employees/new', methods=['GET', 'POST'])
def employee_new():
    allowances = db.get_all_allowances()
    if request.method == 'POST':
        data = {
            'name_code':       request.form.get('name_code', '').strip(),
            'department':      request.form.get('department', '').strip(),
            'contract_type':   request.form.get('contract_type', '정규직'),
            'join_date':       request.form.get('join_date', '') or None,
            'resign_date':     request.form.get('resign_date', '') or None,
            'severance_type':  request.form.get('severance_type', 'DB'),
            'scheduled_hours': float(request.form.get('scheduled_hours', 8)),
            'overtime_hours':  float(request.form.get('overtime_hours', 0)),
            'base_salary':     int(request.form.get('base_salary', 0).replace(',', '')),
            'is_exception':    1 if request.form.get('is_exception') else 0,
            'match_key':       request.form.get('match_key', '').strip(),
            'deduction_count': int(request.form.get('deduction_count', 1)),
        }
        if not data['name_code']:
            flash('이름 코드를 입력해주세요.', 'danger')
            return render_template('employees/form.html', employee=data,
                                   allowances=allowances, mode='new')
        emp_id = db.create_employee(data)
        # 수당 금액 저장
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
        data = {
            'name_code':       request.form.get('name_code', '').strip(),
            'department':      request.form.get('department', '').strip(),
            'contract_type':   request.form.get('contract_type', '정규직'),
            'join_date':       request.form.get('join_date', '') or None,
            'resign_date':     request.form.get('resign_date', '') or None,
            'severance_type':  request.form.get('severance_type', 'DB'),
            'scheduled_hours': float(request.form.get('scheduled_hours', 8)),
            'overtime_hours':  float(request.form.get('overtime_hours', 0)),
            'base_salary':     int(request.form.get('base_salary', '0').replace(',', '')),
            'is_exception':    1 if request.form.get('is_exception') else 0,
            'match_key':       request.form.get('match_key', '').strip(),
            'deduction_count': int(request.form.get('deduction_count', 1)),
        }
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
#  급여작업 (placeholder)
# ════════════════════════════════════════════════════════
@app.route('/payroll')
def payroll():
    return render_template('payroll/index.html')


# ════════════════════════════════════════════════════════
#  인건비 (placeholder)
# ════════════════════════════════════════════════════════
@app.route('/labor-cost')
def labor_cost():
    return render_template('labor_cost/index.html')


# ════════════════════════════════════════════════════════
#  기타분석 (placeholder)
# ════════════════════════════════════════════════════════
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
