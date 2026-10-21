"""Regression tests for positional output parsing and report generation."""

from contextlib import ExitStack
from functools import partial
from pathlib import Path
import sys

from mininterface import Mininterface, run
from openpyxl import Workbook, load_workbook
import pytest

import xlsx_compare


@pytest.mark.parametrize(
    "arguments, expected_output, expected_sheets",
    [
        (["a.xlsx", "b.xlsx"], "comparison_output.xlsx", []),
        (["a.xlsx", "b.xlsx", "c1.xlsx"], "c1.xlsx", []),
        (
            ["a.xlsx", "b.xlsx", "c1.xlsx", "--sheets", "Sheet1", "Sheet2"],
            "c1.xlsx",
            ["Sheet1", "Sheet2"],
        ),
    ],
)
def test_positional_output(arguments, expected_output, expected_sheets):
    """Parse real CLI arguments without opening an interactive interface."""
    env = run(
        xlsx_compare.Env,
        interface=Mininterface,
        config_file=False,
        ask_for_missing=False,
        args=arguments,
    ).env

    assert env.file1 == Path("a.xlsx")
    assert env.file2 == Path("b.xlsx")
    assert env.output == Path(expected_output)
    assert env.sheets == expected_sheets


@pytest.mark.parametrize("output_name", [None, "c1.xlsx"])
def test_main_writes_requested_report(tmp_path, monkeypatch, output_name):
    """Write a real report using the default or an explicit output path."""
    for filename, value in (("a.xlsx", "old"), ("b.xlsx", "new")):
        workbook = Workbook()
        workbook.active.title = "Sheet1"
        workbook.active["A1"] = value
        workbook.save(tmp_path / filename)
        workbook.close()

    monkeypatch.chdir(tmp_path)
    arguments = ["xlsx_compare.py", "a.xlsx", "b.xlsx"]
    if output_name is not None:
        arguments.append(output_name)
    monkeypatch.setattr(sys, "argv", arguments)
    monkeypatch.setattr(
        xlsx_compare,
        "run",
        partial(run, interface=Mininterface, config_file=False, ask_for_missing=False),
    )

    with ExitStack() as resources:
        def open_input(*args, **kwargs):
            workbook = load_workbook(*args, **kwargs)
            resources.callback(workbook.close)
            return workbook

        monkeypatch.setattr(xlsx_compare, "load_workbook", open_input)
        xlsx_compare.main()

    output_path = tmp_path / (output_name or "comparison_output.xlsx")
    assert output_path.is_file()
    report = load_workbook(output_path)
    try:
        assert report.sheetnames == ["df-Sheet1", "COMPARISON"]
        assert report["df-Sheet1"]["A1"].value == "old -> new"
        assert report["COMPARISON"]["A5"].value == "Sheet1"
        assert report["COMPARISON"]["B5"].value == "Differences found"
    finally:
        report.close()
