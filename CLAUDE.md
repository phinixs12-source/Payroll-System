# CLAUDE.md

Behavioral guidelines to reduce common LLM coding mistakes. Merge with project-specific instructions as needed.

**Tradeoff:** These guidelines bias toward caution over speed. For trivial tasks, use judgment.

## 1. Think Before Coding

**Don't assume. Don't hide confusion. Surface tradeoffs.**

Before implementing:
- State your assumptions explicitly. If uncertain, ask.
- If multiple interpretations exist, present them - don't pick silently.
- If a simpler approach exists, say so. Push back when warranted.
- If something is unclear, stop. Name what's confusing. Ask.

## 2. Simplicity First

**Minimum code that solves the problem. Nothing speculative.**

- No features beyond what was asked.
- No abstractions for single-use code.
- No "flexibility" or "configurability" that wasn't requested.
- No error handling for impossible scenarios.
- If you write 200 lines and it could be 50, rewrite it.

Ask yourself: "Would a senior engineer say this is overcomplicated?" If yes, simplify.

## 3. Surgical Changes

**Touch only what you must. Clean up only your own mess.**

When editing existing code:
- Don't "improve" adjacent code, comments, or formatting.
- Don't refactor things that aren't broken.
- Match existing style, even if you'd do it differently.
- If you notice unrelated dead code, mention it - don't delete it.

When your changes create orphans:
- Remove imports/variables/functions that YOUR changes made unused.
- Don't remove pre-existing dead code unless asked.

The test: Every changed line should trace directly to the user's request.

## 4. Goal-Driven Execution

**Define success criteria. Loop until verified.**

Transform tasks into verifiable goals:
- "Add validation" → "Write tests for invalid inputs, then make them pass"
- "Fix the bug" → "Write a test that reproduces it, then make it pass"
- "Refactor X" → "Ensure tests pass before and after"

For multi-step tasks, state a brief plan:
```
1. [Step] → verify: [check]
2. [Step] → verify: [check]
3. [Step] → verify: [check]
```

Strong success criteria let you loop independently. Weak criteria ("make it work") require constant clarification.

---

**These guidelines are working if:** fewer unnecessary changes in diffs, fewer rewrites due to overcomplication, and clarifying questions come before implementation rather than after mistakes.

---

## 5. Project-Specific Rules (Payroll System)

**Stack:** Python Flask · SQLite (database.py) · Jinja2 templates · Bootstrap 5 · openpyxl

### DB 변경 규칙
- `database.py` 스키마(CREATE TABLE) 변경은 명시적 요청 없이 절대 하지 않는다.
- 컬럼 추가 시 반드시 `init_db()` 안 마이그레이션(ALTER TABLE IF NOT EXISTS) 블록에도 추가한다.
- 기존 데이터를 날릴 수 있는 변경(DROP, RENAME)은 먼저 묻는다.

### 코드 스타일 규칙
- 라우트 함수는 `app.py` 하단의 기존 섹션 구조(주석 구분선 `═══`)에 맞춰 추가한다.
- 헬퍼 함수(`_parse_*`, `_calc_*`, `_build_*`)는 관련 라우트 바로 위에 위치한다.
- 템플릿은 `{% extends "base.html" %}` + `{% block content %}` 구조를 유지한다.
- 숫자 포맷은 Jinja 필터 `| fmt_number` 사용, 직접 `{:,}` 쓰지 않는다.

### 도메인 규칙 (한국 급여)
- 직원 식별자는 실명이 아닌 `name_code`(암호화 코드) 사용 — 실명 노출 금지.
- 일할계산 기준: `근무일수 / 해당월 달력일수` (역일 기준).
- 소정/연장근로시간은 **월 기준** (기본값 209h/0h).
- 기본급 공식: `(연봉 ÷ 12) ÷ (소정 + 연장) × 소정`.
- 통상임금 = 기본급 + `is_ordinary_wage=1`인 수당 합계.

### 변경 전 확인이 필요한 것
- `upsert_employee()` 함수 — 연봉/기본급 덮어쓰기 영향 큼
- `payroll_entries` 총액 재계산 — `recalculate_entry_totals()` 반드시 호출
- Step 상태(status) 흐름: `step1 → step2 → step3 → completed` 순서 유지
