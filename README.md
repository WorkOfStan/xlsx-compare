# xlsx_compare

`XLSX Compare` is a Python script that compares two Excel files, sheet by sheet, and identifies differences. This tool is particularly useful for quickly comparing large Excel files and generating an organized output of the differences.

---

## Features

1. **Sheet Comparison**:

   - Identifies sheets that exist only in one file.
   - Compares the contents of sheets present in both files.

2. **Handles Differences**:

   - Records cell-by-cell differences as plain text (`old -> new`) without cell highlighting.
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
   - Compares normalized cell arrays in bulk and builds difference text only for changed cells.

6. **User-Friendly CLI**:

   - Accepts file paths as parameters.
   - Optionally, allows specifying the output filename.

7. **Option to select sheets to compare**:

   - Select sheet names through the CLI or GUI (default: all shared and unique sheets).

---

## Installation

- Python 3.11 or later.
- Install the package:

```bash
pip install git+https://github.com/WorkOfStan/xlsx-compare.git
```

The package limits `mininterface` to version 1.2.2 or earlier for Windows compatibility.
It also installs the `basic` extra required for CLI argument parsing.

For editable installation and running tests, see [Development and contributing](#development-and-contributing).

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
- Handle different sheet sizes by padding both DataFrames with empty values to the maximum row and column counts across the two sheets.
- Compare values at their original coordinates, including the first row. Formatting changes are outside the comparison.
- Write report values as literal text, including values beginning with `=`, so Excel does not execute them as formulas.

---

### Example

```bash
xlsx_compare examples/a.xlsx examples/b.xlsx examples/comparison.xlsx
xlsx_compare examples/a.xlsx examples/b.xlsx examples/comparison-cached-values.xlsx --cached-values
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

If cell `A1` changes from `Value1` to `Value2`, cell `A1` in `df-Sheet4` contains the literal text `Value1 -> Value2`.

Difference sheets preserve the original cell coordinates, add no header row or index column, and leave unchanged cells blank. Changed cells use plain text without cell highlighting.

---

## Development and contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Use an activated virtual environment with Python 3.11 or later. From the repository root, install the project in editable mode and its development tools, then run the regression tests:

```bash
python -m pip install -e . -r requirements-dev.txt
python -c "from pathlib import Path; Path('.tmp').mkdir(exist_ok=True)"
python -m pytest -q -p no:cacheprovider --basetemp=.tmp/pytest tests
```

See [AGENTS.md](AGENTS.md) for repository change requirements and detailed formatting and lint commands.

GitHub Actions runs pytest across Python 3.11–3.15 as required checks and Python 3.16 as an experimental check with prerelease support. Python 3.16 failures, including unavailable interpreters or dependencies, are advisory. The workflow runs pytest first, then Prettier, then Super-Linter; failures on required Python versions stop the later stages.

---

## Debugging

- **Row Count Issue**:
  - The script reads worksheet rows directly, so the first row is compared at its original coordinates.
- **Performance**:
  - The script uses `read_only=True` for large files to reduce memory usage.

---

## Color-Coded Logs

The script currently uses these log colors:

- **Cyan**: Processing progress.
- **Red**: Difference counts and sheets with differences.
- **Yellow `[ONLY]`**: Sheets only in one file.
- **Green `[COMPARE]`**: Shared sheets being read and compared.
- **Blue `[SAME]`**: Shared sheets with no differences.

Only the `[ONLY]`, `[COMPARE]`, and `[SAME]` markers are colored; their following message text uses the terminal's default foreground. Marker shades and contrast depend on the terminal theme, so yellow may be faint on a white background while the message remains readable.

---

## Known Limitations

**Sparklines**:

- Sparkline formatting (in-cell charts) is ignored because `openpyxl` does not support Sparkline Group extensions.
- The warning `Sparkline Group extension is not supported and will be removed` does not prevent cell-value or formula comparison. Input files remain unchanged.

---

## Comparison performance

The comparison normalizes whitespace and missing values, pads both sheets to matching dimensions, and compares their cell arrays in bulk. A boolean mask identifies changed cells, so `old -> new` text is built only for those positions. This avoids per-cell pandas indexing while preserving the report's original cell coordinates and blank unchanged cells.

Sheets are processed sequentially in one process. Each shared sheet is loaded into DataFrames, so memory use still grows with sheet size; reading and writing XLSX files also contribute to total runtime.

---

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.
