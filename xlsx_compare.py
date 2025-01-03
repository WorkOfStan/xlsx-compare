# first install dependencies:
# pip install pandas openpyxl
# v0.1.0

import sys

import pandas as pd
from openpyxl import load_workbook


# Define ANSI escape sequences for colors
class Colors:
    RESET = "\033[0m"
    BOLD = "\033[1m"
    GREEN = "\033[32m"
    YELLOW = "\033[33m"
    RED = "\033[31m"
    BLUE = "\033[34m"
    CYAN = "\033[36m"


# Function to compare two dataframes cell by cell
def compare_dataframes_cell_by_cell(df1, df2):
    # TODO potential to speed up the comparison
    # if df1.shape == df2.shape:
    #    print(f"{Colors.YELLOW}Comparing same shape{Colors.RESET}")
    #    comparison = df1.compare(df2, keep_shape=True, keep_equal=False)
    #    return comparison.dropna(how='all')
    # BUT THE RETURNED VALUE SHOULD BE PROCESSED LIKE THIS:
    #    if differences.empty:
    #        # If there's no difference, just add an info
    #        pd.DataFrame({"Info": ["No differences found"]})
    #            .to_excel(output_writer, sheet_name=f"eq-{sheet[:28]}", index=False)
    #    else:
    #        # Save the differences
    #        differences.to_excel(output_writer, sheet_name=f"df-{sheet[:28]}", index=True)

    # Normalize data for consistent comparison
    df1 = (
        df1.fillna("")
        .astype(str)
        .apply(lambda col: col.str.strip() if col.dtype == "object" else col)
    )
    df2 = (
        df2.fillna("")
        .astype(str)
        .apply(lambda col: col.str.strip() if col.dtype == "object" else col)
    )
    # debug info
    print(f"Shape of File1 ({sheet}): {df1.shape}")
    print(f"Shape of File2 ({sheet}): {df2.shape}")

    # Ensure both DataFrames have the same shape by padding
    max_rows = max(df1.shape[0], df2.shape[0])
    max_cols = max(df1.shape[1], df2.shape[1])
    print(f"Size rows: {max_rows} cols: {max_cols}")

    df1 = df1.reindex(index=range(max_rows), columns=range(max_cols)).fillna("")
    df2 = df2.reindex(index=range(max_rows), columns=range(max_cols)).fillna("")

    # Create a DataFrame to store differences
    diff = pd.DataFrame(index=range(max_rows), columns=range(max_cols))
    # print("diff DataFrame created")

    diff_count = 0

    for i in range(max_rows):
        for j in range(max_cols):
            value1 = df1.iloc[i, j]
            value2 = df2.iloc[i, j]

            # Debugging
            # print(f"Comparing cell ({i}, {j}): '{repr(value1)}' vs. '{repr(value2)}'")

            if value1 != value2:
                # print(f"Difference: ({i}, {j}): '{repr(value1)}' vs. '{repr(value2)}'")
                diff.iloc[i, j] = f"{value1} -> {value2}"
                diff_count = diff_count + 1
    if diff_count > 0:
        print(f"{Colors.RED}{diff_count} difference{Colors.RESET}")
        # TODO Add a function to save the differences to an Excel file

    return diff


# Ensure the script is called with the correct number of arguments
if len(sys.argv) < 3:
    print("Usage: python xlsx_compare.py <file1.xlsx> <file2.xlsx> [output.xlsx]")
    sys.exit(1)

# Get file paths from command-line arguments
file1_path = sys.argv[1]
file2_path = sys.argv[2]
output_path = sys.argv[3] if len(sys.argv) > 3 else "comparison_output.xlsx"

# Processing start
print(f"File1: {file1_path}")
print(f"File2: {file2_path}")

# Load Excel workbooks
# Using the read_only=True parameter makes openpyxl load the workbook in optimized read-only
# mode. This is particularly useful for large files because it reduces memory usage and skips
# features like formatting.
wb1 = load_workbook(file1_path, read_only=True)
print("... workbook1 loaded, now loading workbook2")  # to show processing progress
wb2 = load_workbook(file2_path, read_only=True)

file1_sheets = wb1.sheetnames
file2_sheets = wb2.sheetnames

# Prepare an Excel writer
print("... Preparing Excel writer")  # to show processing progress
output_writer = pd.ExcelWriter(output_path, engine="openpyxl")
print("... Excel writer prepared")  # to show processing progress

# Create a summary list for the COMPARISON sheet
comparison_summary = [
    ["File 1", file1_path],
    ["File 2", file2_path],
    [],  # Add an empty row for separation
]

# Compare sheets
all_sheets = set(file1_sheets) | set(file2_sheets)

for sheet in all_sheets:
    print(
        f"{Colors.CYAN}Processing sheet: {sheet}{Colors.RESET}"
    )  # Cyan for processing progress
    if sheet not in file1_sheets:
        # Sheet only in file2
        print("... only in file2")
        comparison_summary.append([sheet, "Only in file2"])
    elif sheet not in file2_sheets:
        # Sheet only in file1
        print("... only in file1")
        comparison_summary.append([sheet, "Only in file1"])
    else:
        # Sheet in both files
        # dtype=str : all data is treated as strings and any non-standard formats are handled
        # header=None : the first row is also compared
        df1 = pd.read_excel(file1_path, sheet_name=sheet, dtype=str, header=None)
        df2 = pd.read_excel(file2_path, sheet_name=sheet, dtype=str, header=None)
        # debug
        # print(f"Data from File1 ({sheet}):")
        # print(df1)
        # print(f"Data from File2 ({sheet}):")
        # print(df2)

        # Compare content cell by cell
        differences = compare_dataframes_cell_by_cell(df1, df2)
        if differences.isnull().all().all():
            comparison_summary.append([sheet, "No differences"])
            print("... no difference")
        else:
            print(f"{Colors.RED}... some difference{Colors.RESET}")
            # Save differences to a separate sheet
            safe_sheet_name = sheet[:28]  # Truncate for valid Excel sheet name
            # index=False, header=False - not to show additional first row and column with indexes
            differences.to_excel(
                output_writer,
                sheet_name=f"df-{safe_sheet_name}",
                index=False,
                header=False,
            )
            comparison_summary.append([sheet, "Differences found"])

# Write the summary to the COMPARISON sheet
comparison_df = pd.DataFrame(
    comparison_summary,
    columns=["Sheet Name", "Status"],  # Header row for the comparison summary
)
comparison_df.to_excel(output_writer, sheet_name="COMPARISON", index=False)

# Save the output file
output_writer.close()
print(f"Comparison completed. Output saved as '{output_path}'.")
