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

   - Comma-separated list of sheet names to compare (default: all shared and unique sheets)

---

## Installation

### Prerequisites

- Python 3.7 or later.
- Install the required dependencies:

```bash
pip install pandas openpyxl
```

---

## Usage

### Command-Line Interface

```bash
python xlsx_compare.py <file1.xlsx> <file2.xlsx> [output.xlsx] [--sheets Sheet1,Sheet2]
```

- `<file1.xlsx>`: Path to the first Excel file.
- `<file2.xlsx>`: Path to the second Excel file.
- `[output.xlsx]` (optional): Name of the output file. Defaults to `comparison_output.xlsx`.
- `[--sheets Sheet1,Sheet2]` (optional): Switch with comma separated sheet names to compare.

---

### Example

```bash
python xlsx_compare.py example/a.xlsx example/b.xlsx example/comparison.xlsx
# see the example folder, how it looks
```

## Example Output

### COMPARISON Sheet

| **Sheet Name** | **Status**        |
| -------------- | ----------------- |
| `File 1`       | file1.xlsx        |
| `File 2`       | file2.xlsx        |
|                |                   |
| `Sheet1`       | No differences    |
| `Sheet2`       | Only in file1     |
| `Sheet3`       | Only in file2     |
| `Sheet4`       | Differences found |

### Sheet with Differences (e.g., `df-Sheet4`)

| **Cell A1**        | **Cell A2** |
| ------------------ | ----------- |
| `Value1 -> Value2` | `...`       |

---

## Debugging

- **Row Count Issue**:
  - By default, `pandas.read_excel` treats the first row as the header. To compare all rows, the script uses `header=None`.
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

---

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

---

Let me know if you'd like any further customization! 😊
