# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### `Added` for new features

### `Changed` for changes in existing functionality

### `Deprecated` for soon-to-be removed features

### `Removed` for now removed features

### `Fixed` for any bugfixes

### `Security` in case of vulnerabilities

## [0.2.2] - TBD

### Changed

- Simplify `.gitignore` to cover Python caches, package build output, virtual environments, local credentials, IDE settings, and temporary files used by this project.
- Move Python formatting guidelines and local check commands from `README.md` to a new `AGENTS.md` and link to them from the contributing section.
- Align README descriptions of difference sheets, log colors, and comparison behavior with the implementation and set the documented and package Python minimum to 3.11 (the oldest supported Python version).

## [0.2.1] - 2026-10-21

feat: compare formulas as default

### Added

- Add a `Difference Sheet` summary column, test dependencies, and regression tests for comparison accuracy and safe report publication.
- Add a `--cached-values` CLI and GUI option to compare stored formula results and reject missing results before publishing a report.

### Changed

- docs(example): compare cells where a different formula have a different/same result
- docs: Document ignored Sparkline formatting and clarify that its warning does not modify input files or prevent cell comparison.

### Fixed

- Preserve literal missing-value strings and compare formula expressions directly from the open workbooks.
- Apply whitespace normalization consistently and allocate unique report sheet names after truncation.
- Validate sheet selections and report missing inputs with resolved paths and working-directory guidance.
- Close workbook and writer resources on failure and replace reports only after a successful temporary-file save.
- Keep bold, centered, thin-bordered COMPARISON headers.
- Preserve the mininterface CLI and combine its regression tests with formula comparison and file-safety coverage.
- Align Black and isort with Ruff's 120-character line length to prevent conflicting CI formatting changes.
- Set Pylint's maximum line length to 120 characters to match Black, isort, and Ruff.
- Preserve Super-Linter v9.0.0's Pylint defaults for parallel jobs and unavailable dependency imports.

### Security

- Reject output paths that alias either input and write formula-like report values as literal text.

## [0.2.0] - 2026-10-21

feat: installation from the GitHub package, CLI migrated to mininterface

### Added

- installation from the GitHub package
- Run `test_xlsx_compare.py` with pytest in the Polish the code workflow.

### Changed

- CLI migrated to mininterface
- Limit mininterface to version 1.2.2 or earlier for Windows compatibility.
- Allow `--sheets` to accept a space-separated list of sheet names to compare (default: all shared and unique sheets).

### Fixed

- Restore the optional third positional argument for the output file, defaulting to `comparison_output.xlsx`.
- Correct README sheet-selection syntax for the CLI and document GUI list input.
- ci: Explicitly package only `xlsx_compare` so the top-level test module does not prevent installation in CI.
- ci: Install `mininterface[basic]` so fresh installations include the CLI parser dependencies required by the application and tests.
- ci: Create `.tmp` before running pytest in the Polish the code workflow so temporary fixtures work in fresh checkouts.

## [0.1.1] - 2025-06-10

feat: Comma-separated list of sheet names to compare

### Added

- argument `--sheets` allows for a comma-separated list of sheet names to compare (default: all shared and unique sheets)

### Changed

- GitHub Action polish-the-code.yml chains prettier-fix and only then super-linter

### Fixed

- minor pylint identified issues
  - xlsx_compare.py:24:5: W0511: todo potential to speed up the comparison (fixme)
  - xlsx_compare.py:81:9: W0511: todo Add a function to save the differences to an Excel file (fixme)
  - xlsx_compare.py:1:0: C0114: Missing module docstring (missing-module-docstring)
  - xlsx_compare.py:12:0: C0115: Missing class docstring (missing-class-docstring)
  - xlsx_compare.py:12:0: R0903: Too few public methods (0/2) (too-few-public-methods)
  - xlsx_compare.py:23:0: C0116: Missing function or method docstring (missing-function-docstring)
  - xlsx_compare.py:23:36: W0621: Redefining name 'df1' from outer scope (line 142) (redefined-outer-name)
  - xlsx_compare.py:23:41: W0621: Redefining name 'df2' from outer scope (line 143) (redefined-outer-name)

## [0.1.0] - 2025-01-03

### Added

`XLSX Compare` is a Python script that compares two Excel files, sheet by sheet, and identifies differences. This tool is particularly useful for quickly comparing large Excel files and generating an organized output of the differences. CLI output includes a count of differences in red, making it easy to identify and manage the differences.

[Unreleased]: https://github.com/WorkOfStan/xlsx-compare/compare/v0.2.1...HEAD
[0.2.1]: https://github.com/WorkOfStan/xlsx-compare/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/WorkOfStan/xlsx-compare/compare/v0.1.1...v0.2.0
[0.1.1]: https://github.com/WorkOfStan/xlsx-compare/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/WorkOfStan/xlsx-compare/releases/tag/v0.1.0
