# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### `Added` for new features

- argument `--sheets` allows for Comma-separated list of sheet names to compare (default: all shared and unique sheets)

### `Changed` for changes in existing functionality

- GitHub Action polish-the-code.yml chains prettier-fix and only then super-linter

### `Deprecated` for soon-to-be removed features

### `Removed` for now removed features

### `Fixed` for any bugfixes

- minor pylint identified issues
  - xlsx_compare.py:24:5: W0511: todo potential to speed up the comparison (fixme)
  - xlsx_compare.py:81:9: W0511: todo Add a function to save the differences to an Excel file (fixme)
  - xlsx_compare.py:1:0: C0114: Missing module docstring (missing-module-docstring)
  - xlsx_compare.py:12:0: C0115: Missing class docstring (missing-class-docstring)
  - xlsx_compare.py:12:0: R0903: Too few public methods (0/2) (too-few-public-methods)
  - xlsx_compare.py:23:0: C0116: Missing function or method docstring (missing-function-docstring)
  - xlsx_compare.py:23:36: W0621: Redefining name 'df1' from outer scope (line 142) (redefined-outer-name)
  - xlsx_compare.py:23:41: W0621: Redefining name 'df2' from outer scope (line 143) (redefined-outer-name)

### `Security` in case of vulnerabilities

## [0.1.0] - 2025-01-03

### Added

`XLSX Compare` is a Python script that compares two Excel files, sheet by sheet, and identifies differences. This tool is particularly useful for quickly comparing large Excel files and generating an organized output of the differences. CLI output includes a count of differences in red, making it easy to identify and manage the differences.

[Unreleased]: https://github.com/WorkOfStan/xlsx-compare/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/WorkOfStan/xlsx-compare/releases/tag/v0.1.0
