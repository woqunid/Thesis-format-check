# Format Checker Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a local Python Tkinter app that checks uploaded `.docx` thesis documents against parseable template formatting rules.

**Architecture:** Keep Word parsing, rule checks, report rendering, and UI in separate modules. Rules operate on a normalized document model so they can be tested without opening the UI.

**Tech Stack:** Python 3.13, `python-docx`, `tkinter`, `pytest`.

---

### Task 1: Core Model And Report

**Files:**
- Create: `format_checker/__init__.py`
- Create: `format_checker/document_model.py`
- Create: `format_checker/report.py`
- Test: `tests/test_report.py`

- [ ] Write failing tests for report formatting.
- [ ] Implement `Violation` and report rendering.
- [ ] Run `pytest tests/test_report.py -q`.

### Task 2: Parseable Rules

**Files:**
- Create: `format_checker/rules.py`
- Test: `tests/test_rules.py`

- [ ] Write failing tests for title, body, keyword, and unsupported `.doc` behavior.
- [ ] Implement deterministic checks using only parsed `python-docx` attributes.
- [ ] Run `pytest tests/test_rules.py -q`.

### Task 3: Tkinter App

**Files:**
- Create: `format_checker/app.py`
- Create: `run.py`

- [ ] Build upload, check, result display, and save-report actions.
- [ ] Keep UI code separate from rule logic.
- [ ] Run import smoke test with `python -m py_compile`.

### Task 4: Verification

**Files:**
- Existing template: `00 学生用-附件1：2026届毕业论文（设计）撰写规范模板(1).docx`

- [ ] Run all tests with a 60 second timeout.
- [ ] Run checker against the template to confirm parser works.
- [ ] Report unsupported checks explicitly.
