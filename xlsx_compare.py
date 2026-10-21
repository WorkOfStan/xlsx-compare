"""
Compare two Excel files sheet-by-sheet and cell-by-cell and export the differences.
"""

import os
import sys
from contextlib import ExitStack, contextmanager
from dataclasses import dataclass, field
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import BadZipFile

import pandas as pd
from mininterface import run
from mininterface.cli import Positional
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils.exceptions import IllegalCharacterError, InvalidFileException


@dataclass
class Env:
    """Compare Excel files sheet by sheet."""

    file1: Positional[Path]
    """First Excel file path"""
    file2: Positional[Path]
    """Second Excel file path"""
    output: Positional[Path] = Path("comparison_output.xlsx")
    """Output Excel file path"""

    sheets: list[str] = field(default_factory=list)
    "Sheet names to compare (default: all sheets)"


# ANSI escape sequences for console coloring
class Colors:  # pylint: disable=too-few-public-methods
    """ANSI color escape sequences"""

    RESET = "\033[0m"
    BOLD = "\033[1m"
    GREEN = "\033[32m"
    YELLOW = "\033[33m"
    RED = "\033[31m"
    BLUE = "\033[34m"
    CYAN = "\033[36m"


def compare_dataframes_cell_by_cell(df_lft, df_rgt, sheet_handle: str) -> pd.DataFrame:
    """
    Compares two dataframes cell-by-cell and returns a dataframe with differences.
    """

    # Normalize data for consistent comparison
    df_lft = df_lft.fillna("").astype(str).apply(lambda col: col.str.strip())
    df_rgt = df_rgt.fillna("").astype(str).apply(lambda col: col.str.strip())
    # debug info
    print(f"Shape of File1 ({sheet_handle}): {df_lft.shape}")
    print(f"Shape of File2 ({sheet_handle}): {df_rgt.shape}")

    # Ensure both DataFrames have the same shape by padding
    max_rows = max(df_lft.shape[0], df_rgt.shape[0])
    max_cols = max(df_lft.shape[1], df_rgt.shape[1])
    print(f"Size rows: {max_rows} cols: {max_cols}")

    df_lft = df_lft.reindex(index=range(max_rows), columns=range(max_cols)).fillna("")
    df_rgt = df_rgt.reindex(index=range(max_rows), columns=range(max_cols)).fillna("")

    # Create a DataFrame to store differences
    diff = pd.DataFrame(index=range(max_rows), columns=range(max_cols), dtype=object)
    # print("diff DataFrame created")

    diff_count = 0

    for i in range(max_rows):
        for j in range(max_cols):
            value1 = df_lft.iloc[i, j]
            value2 = df_rgt.iloc[i, j]

            # Debugging
            # print(f"Comparing cell ({i}, {j}): '{repr(value1)}' vs. '{repr(value2)}'")

            if value1 != value2:
                # print(f"Difference: ({i}, {j}): '{repr(value1)}' vs. '{repr(value2)}'")
                diff.iloc[i, j] = f"{value1} -> {value2}"
                diff_count = diff_count + 1
    if diff_count > 0:
        print(f"{Colors.RED}{diff_count} difference found{Colors.RESET}")

    return diff


def save_differences_to_excel(writer: pd.ExcelWriter, sheet_name: str, differences_df: pd.DataFrame) -> str:
    """
    Saves the differences dataframe to the given Excel writer under a safe sheet name.
    """
    safe_sheet_name = sheet_name[:28]  # Truncate for valid Excel sheet name
    base_name = f"df-{safe_sheet_name}"
    report_name = base_name
    existing_names = {name.casefold() for name in writer.book.sheetnames}
    suffix_number = 2
    while report_name.casefold() in existing_names:
        suffix = f"-{suffix_number}"
        report_name = f"{base_name[: 31 - len(suffix)]}{suffix}"
        suffix_number += 1
    # index=False, header=False - not to show additional first row and column with indexes
    write_text_dataframe(writer, report_name, differences_df, header=False)
    return report_name


def write_text_dataframe(writer, sheet_name, dataframe, header=True):
    """Write report cells as literal text, never executable Excel formulas."""
    worksheet = writer.book.create_sheet(sheet_name)
    rows = dataframe.itertuples(index=False, name=None)
    if header:
        worksheet.append(list(dataframe.columns))
        thin_side = Side(style="thin")
        header_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="top")
            cell.border = header_border
    for row in rows:
        worksheet.append([None if pd.isna(value) else str(value) for value in row])
    for row in worksheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                cell.data_type = "s"


def select_sheets(requested, file1_sheets, file2_sheets):
    """Validate selections before creating any output."""
    available = set(file1_sheets) | set(file2_sheets)
    if not requested:
        return sorted(available)
    selected = [name.strip() for name in requested]
    if any(not name for name in selected):
        raise ValueError("--sheets must not contain empty sheet names.")
    # Sheet is only in args.sheet
    missing = [name for name in selected if name not in available]
    if missing:
        raise ValueError(f"Sheets absent from both inputs: {', '.join(missing)}")
    return list(dict.fromkeys(selected))


def validate_paths(args):
    """Resolve paths and prevent output from replacing either input."""
    inputs = [Path(args.file1).resolve(), Path(args.file2).resolve()]
    output = Path(args.output).resolve()
    for source in inputs:
        if not source.is_file():
            raise FileNotFoundError(
                f"Input file not found: '{source}'. Relative paths use the terminal's "
                f"working directory ('{Path.cwd()}')."
            )
        if output == source or (output.exists() and output.samefile(source)):
            raise ValueError(f"Output must differ from both input files: '{output}'.")
    return inputs, output


def read_sheet(worksheet):
    """Preserve text and formula expressions without missing-value inference."""
    return pd.DataFrame(worksheet.iter_rows(values_only=True), dtype=object)


def remove_temporary_file(path):
    """Remove temporary output if it has not been published."""
    if path.exists():
        path.unlink()


@contextmanager
def open_workbook(path):
    """Close both the workbook and its input stream, including on load failures."""
    # Load Excel workbooks
    # Using the read_only=True parameter makes openpyxl load the workbook in optimized read-only
    # mode. This is particularly useful for large files because it reduces memory usage and skips
    # features like formatting.
    with path.open("rb") as input_stream:
        workbook = load_workbook(input_stream, read_only=True, data_only=False)
        try:
            yield workbook
        finally:
            workbook.close()


def compare_sheet(sheet, workbooks, output_writer):
    """Write differences for one sheet and return its summary row."""
    wb1, wb2 = workbooks
    print(f"{Colors.CYAN}Processing sheet: {sheet}{Colors.RESET}")  # Cyan for processing progress
    if sheet not in wb1.sheetnames:
        # Sheet only in file2
        print("... only in file2")
        return [sheet, "Only in file2", ""]
    if sheet not in wb2.sheetnames:
        # Sheet only in file1
        print("... only in file1")
        return [sheet, "Only in file1", ""]

    # Sheet in both files
    # dtype=str : all data is treated as strings and any non-standard formats are handled
    # header=None : the first row is also compared
    df1 = read_sheet(wb1[sheet])
    df2 = read_sheet(wb2[sheet])
    # debug
    # print(f"Data from File1 ({sheet}):")
    # print(df1)
    # print(f"Data from File2 ({sheet}):")
    # print(df2)

    # Compare content cell by cell
    differences = compare_dataframes_cell_by_cell(df1, df2, sheet)
    if differences.isnull().all().all():
        print("... no difference")
        return [sheet, "No differences", ""]
    print(f"{Colors.RED}... some difference{Colors.RESET}")
    # Save differences to a separate sheet
    report_name = save_differences_to_excel(output_writer, sheet, differences)
    return [sheet, "Differences found", report_name]


def write_report(args, workbooks, selected_sheets, output_writer):
    """Populate difference sheets and the formatted comparison summary."""
    # Keep a visible sheet so writer cleanup cannot mask a comparison error.
    output_writer.book.create_sheet("COMPARISON")
    print("... Excel writer prepared")  # to show processing progress

    # Create a summary list for the COMPARISON sheet
    comparison_summary = [
        ["File 1", args.file1, ""],
        ["File 2", args.file2, ""],
        [],  # Add an empty row for separation
    ]

    for sheet in selected_sheets:
        comparison_summary.append(compare_sheet(sheet, workbooks, output_writer))

    # Write the summary to the COMPARISON sheet
    comparison_df = pd.DataFrame(
        comparison_summary,
        columns=[
            "Sheet Name",
            "Status",
            "Difference Sheet",
        ],  # Header row for the comparison summary
    )
    del output_writer.book["COMPARISON"]
    write_text_dataframe(output_writer, "COMPARISON", comparison_df)


def run_comparison(args, resources):
    """Compare workbooks and publish the report only after a successful save."""
    inputs, output = validate_paths(args)
    # output_path = args.output

    # Processing start
    print(f"File1: {args.file1}")
    print(f"File2: {args.file2}")

    wb1 = resources.enter_context(open_workbook(inputs[0]))
    print("... workbook1 loaded, now loading workbook2")  # to show processing progress
    wb2 = resources.enter_context(open_workbook(inputs[1]))

    # Compare sheets
    selected_sheets = select_sheets(args.sheets, wb1.sheetnames, wb2.sheetnames)

    # Prepare an Excel writer
    print("... Preparing Excel writer")  # to show processing progress
    with NamedTemporaryFile(dir=output.parent, prefix=".xlsx-compare-", suffix=".xlsx", delete=False) as temporary_file:
        temporary_path = Path(temporary_file.name)
    resources.callback(remove_temporary_file, temporary_path)
    with temporary_path.open("w+b") as output_stream, pd.ExcelWriter(output_stream, engine="openpyxl") as output_writer:
        write_report(args, (wb1, wb2), selected_sheets, output_writer)

    # Save the output file
    os.replace(temporary_path, output)
    print(f"Comparison completed. Output saved as '{args.output}'.")


def main():
    """Main function to compare two Excel files."""
    args = run(Env).env
    try:
        with ExitStack() as resources:
            run_comparison(args, resources)
    except (
        OSError,
        ValueError,
        BadZipFile,
        InvalidFileException,
        IllegalCharacterError,
        SyntaxError,
        KeyError,
    ) as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
