# Agent instructions

## Python formatting

GitHub Action calling Super-linter is taking care of formatting. Description below is just to explain what's going on.

Use a 120-character line length for Python. Super-Linter loads Black, isort, and Pylint settings from
`.github/linters/.python-black`, `.github/linters/.isort.cfg`, and `.github/linters/.python-lint`;
these match its Ruff configuration. This keeps formatting and lint checks consistent.
Custom configuration files replace Super-Linter's defaults. The Pylint configuration preserves
Super-Linter v9.0.0's `jobs=0` and `disable=import-error` settings and adds the 120-character limit.
Compare these defaults when upgrading Super-Linter.

To run the same checks locally after installing `black`, `isort`, `pylint`, and `ruff`:

```bash
python -m black --check --config .github/linters/.python-black xlsx_compare.py test_xlsx_compare.py
python -m isort --check-only --settings-path .github/linters/.isort.cfg xlsx_compare.py test_xlsx_compare.py
python -m ruff format --check --line-length 120 xlsx_compare.py test_xlsx_compare.py
python -m pylint --persistent=n --rcfile .github/linters/.python-lint xlsx_compare.py test_xlsx_compare.py
```
