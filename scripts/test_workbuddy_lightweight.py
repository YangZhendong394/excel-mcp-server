#!/usr/bin/env python3
"""End-to-end tests for the packaged lightweight WorkBuddy Skill."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path


def run(script: Path, root: Path, *arguments: str, expect_success: bool = True) -> dict:
    command = [
        sys.executable,
        str(script),
        "--allow-dir",
        str(root),
        *arguments,
    ]
    result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8")
    try:
        payload = json.loads(result.stdout)
    except json.JSONDecodeError as exc:
        raise AssertionError(
            f"Command did not return JSON.\nstdout={result.stdout}\nstderr={result.stderr}"
        ) from exc
    if expect_success and result.returncode != 0:
        raise AssertionError(f"Command failed: {command}\n{payload}\n{result.stderr}")
    if not expect_success and result.returncode == 0:
        raise AssertionError(f"Command unexpectedly succeeded: {command}\n{payload}")
    return payload


def test_common_workbook_features(skill: Path, root: Path) -> None:
    """Ensure a normal edit-copy operation preserves commonly used workbook features."""
    vendor = skill / "vendor"
    if not vendor.is_dir():
        raise AssertionError(f"Bundled vendor folder is missing: {vendor}")
    sys.path.insert(0, str(vendor))

    from openpyxl import Workbook, load_workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.comments import Comment
    from openpyxl.formatting.rule import ColorScaleRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table, TableStyleInfo

    source = root / "rich_source.xlsx"
    output = root / "rich_edited.xlsx"

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "销售数据"
    sheet.append(["产品", "数量", "单价", "金额"])
    sheet.append(["苹果", 3, 2.5, 7.5])
    sheet.append(["香蕉", 4, 1.8, 7.2])
    sheet.append(["橙子", 5, 3.2, 16.0])

    # Representative cell formatting and dimensions.
    thin_blue = Side(style="thin", color="FF1F4E78")
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="FF1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin_blue)
    for row in range(2, 5):
        sheet.cell(row, 3).number_format = '"$"#,##0.00'
        sheet.cell(row, 4).number_format = '"$"#,##0.00'
    sheet.column_dimensions["A"].width = 24
    sheet.row_dimensions[1].height = 28

    # Common worksheet structures and objects.
    sheet.freeze_panes = "B2"
    sheet.auto_filter.ref = "A1:D4"
    sheet.merge_cells("G1:H1")
    sheet["G1"] = "合并标题"
    sheet["F2"] = "=SUM(D2:D4)"
    sheet["I2"] = "外部链接"
    sheet["I2"].hyperlink = "https://example.com"
    sheet["I3"] = "带批注单元格"
    sheet["I3"].comment = Comment("验收批注内容", "WorkBuddy Test")

    validation = DataValidation(type="whole", operator="between", formula1=1, formula2=100)
    sheet.add_data_validation(validation)
    validation.add("J2:J4")

    sheet.conditional_formatting.add(
        "D2:D4",
        ColorScaleRule(
            start_type="min",
            start_color="FFF8696B",
            mid_type="percentile",
            mid_value=50,
            mid_color="FFFFEB84",
            end_type="max",
            end_color="FF63BE7B",
        ),
    )

    table = Table(displayName="SalesTable", ref="A1:D4")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    sheet.add_table(table)

    chart = BarChart()
    chart.title = "销售金额"
    chart.add_data(Reference(sheet, min_col=4, min_row=1, max_row=4), titles_from_data=True)
    chart.set_categories(Reference(sheet, min_col=1, min_row=2, max_row=4))
    sheet.add_chart(chart, "K2")

    hidden = workbook.create_sheet("内部字典")
    hidden["A1"] = "保留隐藏工作表"
    hidden.sheet_state = "hidden"
    workbook.save(source)
    workbook.close()

    # Change an unrelated cell through the exact packaged CLI path.
    written = run(
        skill / "scripts" / "excel_workbook.py",
        root,
        "write",
        "--input",
        source.name,
        "--output",
        output.name,
        "--sheet",
        "销售数据",
        "--start-cell",
        "M1",
        "--rows-json",
        json.dumps([["本次验收写入"]], ensure_ascii=False),
    )
    assert written["ok"] and written["source_preserved"] and written["cells_written"] == 1
    assert source.is_file() and output.is_file()

    check = load_workbook(output, data_only=False)
    try:
        result_sheet = check["销售数据"]
        validations = list(result_sheet.data_validations.dataValidation)
        checks = {
            "cell_style": (
                result_sheet["A1"].font.bold
                and result_sheet["A1"].fill.fill_type == "solid"
                and result_sheet["A1"].alignment.horizontal == "center"
                and result_sheet["A1"].border.bottom.style == "thin"
            ),
            "number_format": result_sheet["C2"].number_format == '"$"#,##0.00',
            "merged_cells": "G1:H1" in {str(item) for item in result_sheet.merged_cells.ranges},
            "freeze_panes": result_sheet.freeze_panes == "B2",
            "worksheet_auto_filter": result_sheet.auto_filter.ref == "A1:D4",
            "excel_table": "SalesTable" in result_sheet.tables,
            "formula_preserved": result_sheet["F2"].value == "=SUM(D2:D4)",
            "hyperlink_preserved": (
                result_sheet["I2"].hyperlink is not None
                and result_sheet["I2"].hyperlink.target == "https://example.com"
            ),
            "comment_preserved": (
                result_sheet["I3"].comment is not None
                and result_sheet["I3"].comment.text == "验收批注内容"
            ),
            "data_validation": any("J2:J4" in str(item.sqref) for item in validations),
            "conditional_formatting": len(result_sheet.conditional_formatting) == 1,
            "chart": len(result_sheet._charts) == 1,
            "row_column_dimensions": (
                result_sheet.column_dimensions["A"].width == 24
                and result_sheet.row_dimensions[1].height == 28
            ),
            "hidden_sheet": (
                "内部字典" in check.sheetnames
                and check["内部字典"].sheet_state == "hidden"
            ),
            "new_cell_written": result_sheet["M1"].value == "本次验收写入",
        }
    finally:
        check.close()

    failed = [name for name, passed in checks.items() if not passed]
    if failed:
        raise AssertionError(
            "Workbook feature preservation failed: "
            + ", ".join(failed)
            + "\nChecks: "
            + json.dumps(checks, ensure_ascii=False, indent=2)
        )
    print("WORKBOOK_FEATURE_PRESERVATION=SUCCESS")
    print("WORKBOOK_FEATURE_CHECKS=" + json.dumps(checks, ensure_ascii=False))


def main() -> int:
    if len(sys.argv) != 2:
        raise SystemExit("Usage: test_workbuddy_lightweight.py EXTRACTED_SKILL_FOLDER")
    skill = Path(sys.argv[1]).resolve(strict=True)
    script = skill / "scripts" / "excel_workbook.py"
    if not script.is_file():
        raise SystemExit(f"Missing Skill script: {script}")

    with tempfile.TemporaryDirectory(prefix="workbuddy-excel-lite-") as temp:
        root = Path(temp)
        created = run(
            script,
            root,
            "create",
            "--path",
            "source.xlsx",
            "--sheets",
            "数据,说明",
        )
        assert created["ok"] and len(created["sheets"]) == 2

        inspected = run(script, root, "inspect", "--path", "source.xlsx")
        assert inspected["ok"] and [item["name"] for item in inspected["sheets"]] == ["数据", "说明"]

        written = run(
            script,
            root,
            "write",
            "--input",
            "source.xlsx",
            "--output",
            "edited.xlsx",
            "--sheet",
            "数据",
            "--start-cell",
            "A1",
            "--rows-json",
            json.dumps([["商品", "数量"], ["苹果", 12], ["香蕉", 8]], ensure_ascii=False),
        )
        assert written["ok"] and written["source_preserved"] and written["cells_written"] == 6
        assert (root / "source.xlsx").is_file() and (root / "edited.xlsx").is_file()

        read = run(
            script,
            root,
            "read",
            "--path",
            "edited.xlsx",
            "--sheet",
            "数据",
            "--range",
            "A1:B3",
        )
        assert read["ok"]
        assert read["values"] == [["商品", "数量"], ["苹果", 12], ["香蕉", 8]]

        refused = run(script, root, "create", "--path", "source.xlsx", "--sheets", "No overwrite", expect_success=False)
        assert refused["ok"] is False

        outside = root.parent / "workbuddy-excel-outside-test.xlsx"
        escaped = run(
            script,
            root,
            "create",
            "--path",
            str(outside),
            "--sheets",
            "Blocked",
            expect_success=False,
        )
        assert escaped["ok"] is False
        assert not outside.exists()

        test_common_workbook_features(skill, root)

    print("LIGHTWEIGHT_SKILL_E2E=SUCCESS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
