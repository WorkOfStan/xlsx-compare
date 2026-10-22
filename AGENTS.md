# Agent instructions

## Project layout

- Keep the application in the root-level `xlsx_compare.py` module.
- Keep regression tests in `tests/` and sample inputs and reports in `examples/`.
- Declare runtime dependencies in `pyproject.toml` and development tools in `requirements-dev.txt`.
- Use `git mv` for file and directory moves, and update affected code, documentation, and CI paths.

## Setup and tests

Use Python 3.11 or later and an activated virtual environment. Run commands from the repository root:

```bash
python -m pip install -e . -r requirements-dev.txt
python -c "from pathlib import Path; Path('.tmp').mkdir(exist_ok=True)"
python -m pytest -q -p no:cacheprovider --basetemp=.tmp/pytest tests
```

Do not load pytest's cache plugin. Keep test and temporary output under `.tmp/`.
Do not inspect, lint, or recurse into `.tmp`, `.pytest-tmp`, `.pytest_cache`, `.venv`, or generated build artifacts.
If pytest temporary-directory cleanup fails on Windows, remove only `.tmp/pytest` and rerun tests serially.

CI tests Python 3.11–3.15 as required versions and Python 3.16 as an experimental version with prerelease support.
Keep only the Python 3.16 job advisory. Preserve required-version fail-fast behavior and the dependency chain
`pytest` -> `prettier-fix` -> `super-linter` in `.github/workflows/polish-the-code.yml`.

## Change requirements

- Preserve source comments unless removing a TODO resolved by the change; comments may be translated to English.
- Keep `main()` at the bottom of Python files.
- Preserve the CLI, cell coordinates, whitespace normalization, formula comparison modes, and `old -> new` report text.
- Keep report values as literal text, protect both input files, and publish reports only after a successful save.
- Use synthetic workbooks or the committed examples for validation; do not read private workbooks under `delete/`.
- Document changes in English at the end of the relevant changelog block, starting each entry with a present-tense verb.
- Update these instructions and README when commands, paths, or documented behavior change.

## Python formatting

GitHub Action calling Super-linter is taking care of formatting. Description below is just to explain what's going on.

Use a 120-character line length for Python. Super-Linter loads Black, isort, and Pylint settings from
`.github/linters/.python-black`, `.github/linters/.isort.cfg`, and `.github/linters/.python-lint`;
these match its Ruff configuration. This keeps formatting and lint checks consistent.
Custom configuration files replace Super-Linter's defaults. The Pylint configuration preserves
Super-Linter v9.0.0's `jobs=0` and `disable=import-error` settings and adds the 120-character limit.
Compare these defaults when upgrading Super-Linter.

The development requirements install `black`, `isort`, `pylint`, and `ruff`. Run the same checks locally with:

```bash
python -m black --check --config .github/linters/.python-black xlsx_compare.py tests
python -m isort --check-only --settings-path .github/linters/.isort.cfg xlsx_compare.py tests
python -m ruff format --check --line-length 120 xlsx_compare.py tests
python -m pylint --persistent=n --rcfile .github/linters/.python-lint xlsx_compare.py tests/test_xlsx_compare.py
```
