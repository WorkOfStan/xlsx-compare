# xlsx_compare

`XLSX Compare` is a Python script that compares two Excel files, sheet by sheet, and identifies differences. This tool is particularly useful for quickly comparing large Excel files and generating an organized output of the differences.

---

## Features

1. **Sheet Comparison**:

   - Identifies sheets that exist only in one file.
   - Compares the contents of sheets present in both files.

2. **Handles Differences**:

   - Highlights cell-by-cell differences for sheets present in both files.
   - Writes differences into a separate sheet named `df-<sheet_name>`.
   - Truncates long report sheet names and adds numbered suffixes when names collide; the summary's `Difference Sheet` column identifies each report sheet.
   - CLI output contains difference count (in red).

3. **Generates a Summary Sheet**:

   - Creates a `COMPARISON` sheet summarizing:
     - Sheets that exist only in `file1` or `file2`.
     - Sheets with no differences.
     - Sheets with differences.

4. **Organized Output**:

   - Sheets that exist in only one file are noted in the `COMPARISON` sheet.
   - Sheets with differences are written with only the changed cells in a new sheet.

5. **Performance Optimization**:

   - Uses `read_only=True` mode with `openpyxl` for processing large files efficiently.
   - Handles different sheet sizes by padding smaller sheets to match dimensions.

6. **User-Friendly CLI**:

   - Accepts file paths as parameters.
   - Optionally, allows specifying the output filename.

7. **Option to select sheets to compare**:

   - Select sheet names through the CLI or GUI (default: all shared and unique sheets).

---

## Installation

- Python 3.7 or later.
- Install the package:

```bash
pip install git+https://github.com/WorkOfStan/xlsx-compare.git
```

The package limits `mininterface` to version 1.2.2 or earlier for Windows compatibility.
It also installs the `basic` extra required for CLI argument parsing.

For running the regression tests, also install `pytest`:

```bash
python -m pip install -r requirements-dev.txt
python -c "from pathlib import Path; Path('.tmp').mkdir(exist_ok=True)"
python -m pytest -q -p no:cacheprovider --basetemp=.tmp/pytest
```

---

## Usage

### Command-Line Interface

Run `xlsx_compare`, the command installed by the Python package, in a terminal such as Bash or PowerShell.
The usage synopsis below uses `<...>` for required arguments and `[...]` for optional arguments; omit these markers when running the command.

```text
xlsx_compare <file1.xlsx> <file2.xlsx> [output.xlsx] [--sheets Sheet1 Sheet2 ...] [--cached-values]
```

- `<file1.xlsx>`: Path to the first Excel file.
- `<file2.xlsx>`: Path to the second Excel file.
- `[output.xlsx]` (optional): Name of the output file. Defaults to `comparison_output.xlsx`.
- `[--sheets Sheet1 Sheet2 ...]` (optional): Sheet names separated by spaces. Quote names containing spaces. Omit this option to compare all sheets.
- `[--cached-values]` (optional): Compare cached formula results instead of formula expressions. Disabled by default.

```bash
xlsx_compare a.xlsx b.xlsx c1.xlsx --sheets Sheet1 "Sheet 2"
```

`--sheets Sheet1,Sheet2` is treated as one sheet name, not two.

To compare cached results:

```bash
xlsx_compare a.xlsx b.xlsx report.xlsx --cached-values
```

### Graphical User Interface

In the sheet-selection field, enter a list with square brackets and quoted names separated by commas:

```json
["Sheet1", "Sheet 2"]
```

Use `[]` to compare all sheets. Plain text such as `Sheet1,Sheet2` or `Sheet1 Sheet2` is not the correct GUI format.

Enable the `cached_values` checkbox to compare cached formula results. Leave it unchecked to compare formula expressions.

Relative paths are resolved from the terminal's current working directory. Use absolute input paths or change to the directory containing your workbooks before running the script. Missing-file errors display the resolved path and working directory.

The output must refer to a different file from both inputs, including aliases and hard links. An existing output report is replaced only after the new report is saved successfully. Expected errors return a nonzero exit status without publishing a partial report.

Sheet selections must not contain empty names or names missing from both inputs. Duplicate names are compared once, in the order supplied. A sheet present in only one input is still included in the summary.

### Comparison behavior

- Compare formula expressions, such as `=SUM(A1:A3)`, by default. With `--cached-values`, compare their stored results using `openpyxl`'s `data_only=True`. The script does not calculate formulas or verify whether cached results are current; recalculate and save both inputs in Excel before comparing cached values.
- If a formula on a selected sheet shared by both inputs has no cached result, report its file, sheet, and cell and exit with status `1` before creating a report. An existing report remains unchanged. Zero, `False`, and cached empty text are valid results. Missing results on unselected or one-sided sheets do not block comparison.
- Preserve literal text such as `NA`, `NULL`, and `N/A`; these are distinct from blank cells.
- Ignore leading and trailing whitespace consistently. Treat missing cells as empty strings and normalize cell values to strings.
- Compare values at their original coordinates, including the first row. Formatting changes are outside the comparison.
- Write report values as literal text, including values beginning with `=`, so Excel does not execute them as formulas.

---

### Example

```bash
xlsx_compare example/a.xlsx example/b.xlsx example/comparison.xlsx
xlsx_compare example/a.xlsx example/b.xlsx example/comparison-cached-values.xlsx --cached-values
# see the example folder, how it looks
```

## Example Output

### COMPARISON Sheet

| **Sheet Name** | **Status**        | **Difference Sheet** |
| -------------- | ----------------- | -------------------- |
| `File 1`       | file1.xlsx        |                      |
| `File 2`       | file2.xlsx        |                      |
|                |                   |                      |
| `Sheet1`       | No differences    |                      |
| `Sheet2`       | Only in file1     |                      |
| `Sheet3`       | Only in file2     |                      |
| `Sheet4`       | Differences found | `df-Sheet4`          |

### Sheet with Differences (e.g., `df-Sheet4`)

| **Cell A1**        | **Cell A2** |
| ------------------ | ----------- |
| `Value1 -> Value2` | `...`       |

---

## Debugging

- **Row Count Issue**:
  - The script reads worksheet rows directly, so the first row is compared at its original coordinates.
- **Performance**:
  - The script uses `read_only=True` for large files to reduce memory usage.
- **Shapes Mismatch**:
  - Pads the smaller DataFrame with empty values to match the dimensions of the larger DataFrame.

---

## Color-Coded Logs

The script provides color-coded logs for better readability:

- **Cyan**: Processing progress.
- **Yellow**: Sheets only in one file. (TODO)
- **Green**: Sheets being compared. (TODO)
- **Blue**: Sheets with no differences. (TODO)
- **Red**: Sheets with differences.

---

## Known Issues

1. **DataFrames with Different Shapes**:

   - Handled by padding smaller DataFrames with empty values.

2. **Hidden Characters in Data**:
   - Automatically strips leading/trailing whitespace.

3. **Sparklines**:
   - Sparkline formatting (in-cell charts) is ignored because `openpyxl` does not support Sparkline Group extensions.
   - The warning `Sparkline Group extension is not supported and will be removed` does not prevent cell-value or formula comparison. Input files remain unchanged.

---

## Future Enhancements

1. Improve comparison speed by leveraging optimized pandas operations. E.g. developing this idea:

```python
def compare_dataframes_cell_by_cell(df_lft, df_rgt, sheet_handle: str) -> pd.DataFrame:
    """
    Compares two dataframes cell-by-cell and returns a dataframe with differences.
    """
    # TODO potential to speed up the comparison
    # if df_lft.shape == df_rgt.shape:
    #    print(f"{Colors.YELLOW}Comparing same shape{Colors.RESET}")
    #    comparison = df_lft.compare(df_rgt, keep_shape=True, keep_equal=False)
    #    return comparison.dropna(how='all')
    # BUT THE RETURNED VALUE SHOULD BE PROCESSED LIKE THIS:
    #    if differences.empty:
    #        # If there's no difference, just add an info
    #        pd.DataFrame({"Info": ["No differences found"]})
    #            .to_excel(output_writer, sheet_name=f"eq-{sheet_handle[:28]}", index=False)
    #    else:
    #        # Save the differences
    #        differences.to_excel(output_writer, sheet_name=f"df-{sheet_handle[:28]}", index=True)

    # Normalize data for consistent comparison
```

---

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

### Python formatting

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

---

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

---

Let me know if you'd like any further customization! 😊
