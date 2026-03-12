# AGENTS.md

This file gives coding agents a fast, practical default for this repository.

## Project Snapshot
- Name: `education-pdf-merger` (Windows desktop utility)
- Stack: Python + customtkinter + Win32 COM + pywinauto + PDF tooling
- Goal: Convert education documents (Word/Excel/PowerPoint/Ichitaro/images) to PDF and merge with TOC, page numbers, bookmarks, and optional compression.
- Platform: Windows only (Office COM + UI automation dependencies)

## Environment Setup
```powershell
python -m venv venv
venv\Scripts\activate
pip install -r requirements-dev.txt
pre-commit install
```

## Run
```powershell
python run_app.py
```

## Test and Quality
```powershell
pytest
pytest tests/test_pdf_converter.py
pytest -m unit
pytest -m integration
pytest -m "not slow"

ruff check --fix .
ruff format .
mypy --config-file mypy.ini
bandit -r . --skip B101,B601 --exclude tests/
pre-commit run --all-files
```

## Build
```powershell
build.bat
# or
pyinstaller build_installer.spec --clean
```

## High-Impact Files
- `run_app.py`: GUI entrypoint
- `pdf_merge_orchestrator.py`: top-level workflow orchestration
- `document_collector.py`: source file discovery + conversion routing
- `pdf_converter.py`: format conversion logic
- `pdf_processor.py`: PDF merge/TOC/bookmarks/page-number/compress flow
- `config_loader.py`: config loading, defaults, and validation entry
- `path_validator.py`: path checks and safety
- `update_excel_files.py`: Excel update processing
- `gui/`: UI layer
- `tests/`: unit and integration tests

## Coding Rules For Agents
- Keep Python compatibility at `>=3.8`.
- Prefer explicit type hints for new/changed functions.
- Use project custom exceptions instead of raw generic exceptions.
- Preserve Windows path behavior and filesystem safety checks.
- Avoid broad `except:` blocks; include actionable context when raising.
- For behavior changes, add or update tests in `tests/`.

## Testing Strategy
- Use `unit` tests for pure logic and file-path handling.
- Use `integration` tests when touching COM/UI-bound behavior.
- Mark long-running tests with `slow`.

## Common Constraints
- Office/Ichitaro conversion depends on local app installs and UI automation state.
- Some tests/features require Ghostscript path in config.
- Config changes should remain aligned with `config.json.example` and loader/validator logic.

## Agent Workflow Checklist
1. Read touched module(s) and related tests first.
2. Implement minimal scoped change.
3. Run targeted tests, then broader suite if needed.
4. Run lint/type/security checks for modified areas.
5. Summarize behavior change and risk in PR/commit notes.
