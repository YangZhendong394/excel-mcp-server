#!/usr/bin/env python3
"""Stability and compatibility tests for the packaged lightweight WorkBuddy Skill."""

from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


def run(script: Path, root: Path, *arguments: str, expect_success: bool = True) -> dict:
    command = [sys.executable, str(script), "--allow-dir", str(root), *arguments]
    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
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


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def test_feature_matrix(skill: Path) -> None:
    matrix_path = skill / "FEATURE-COMPATIBILITY.md"
    assert matrix_path.is_file(), "Feature comparison matrix is missing from the packaged Skill"
    content = matrix_path.read_text(encoding="utf-8")
    default_section = content.split("### 原版额外的 2 个可选 VBA 工具", 1)[0]
    rows = [
        line for line in default_section.splitlines()
        if line.startswith("| ")
        and not line.startswith("|---")
        and not line.startswith("| 原版分组")
    ]
    partial = sum("部分保留" in line for line in rows)
    missing = sum("未保留" in line for line in rows)
    assert len(rows) == 46, f"Expected 46 default MCP tools, found {len(rows)}"
    assert partial == 5, f"Expected 5 partially retained tools, found {partial}"
    assert missing == 41, f"Expected 41 unimplemented tools, found {missing}"

    optional_section = content.split("### 原版额外的 2 个可选 VBA 工具", 1)[1]
    optional_rows = [
        line for line in optional_section.splitlines()
        if line.startswith("| ")
        and not line.startswith("|---")
        and not line.startswith("| 原版可选 MCP 工具")
    ]
    assert len(optional_rows) == 2, f"Expected 2 separately listed optional VBA tools, found {len(optional_rows)}"
    print("MCP_FEATURE_MATRIX=46_TOTAL;5_PARTIAL;41_NOT_IMPLEMENTED;2_OPTIONAL_VBA_SEPARATE")

def test_formula_cache_behavior(skill: Path, root: Path) -> None:
    from openpyxl import Workbook

    source = root / "formula_cache_probe.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "计算测试"
    sheet["A1"] = "数值1"
    sheet["A2"] = 10
    sheet["A3"] = 20
    sheet["B1"] = "公式"
    sheet["B2"] = "=SUM(A2:A3)"
    workbook.save(source)
    workbook.close()

    script = skill / "scripts" / "excel_workbook.py"
    formula = run(
        script, root, "read", "--path", source.name, "--sheet", "计算测试",
        "--range", "B2:B2",
    )
    assert formula["values"] == [["=SUM(A2:A3)"]], formula

    cached = run(
        script, root, "read", "--path", source.name, "--sheet", "计算测试",
        "--range", "B2:B2", "--values-only",
    )
    assert cached["values"] == [[None]], cached
    print("FORMULA_EXPRESSION_READ=SUCCESS")
    print("FORMULA_CACHE_MISSING_REPORTED_AS_NULL=SUCCESS")


def test_real_xlsm_vba_passthrough(skill: Path, root: Path, macro_fixture: Path) -> None:
    from openpyxl import load_workbook

    source = root / "macro_source.xlsm"
    output = root / "macro_roundtrip.xlsm"
    source.write_bytes(macro_fixture.read_bytes())
    original_hash = sha256(source)

    with zipfile.ZipFile(source) as archive:
        names = set(archive.namelist())
        assert "xl/vbaProject.bin" in names, "Real macro fixture has no xl/vbaProject.bin"
        original_vba = archive.read("xl/vbaProject.bin")

    workbook = load_workbook(source, keep_vba=True, data_only=False)
    try:
        sheet_name = workbook.sheetnames[0]
    finally:
        workbook.close()

    result = run(
        skill / "scripts" / "excel_workbook.py",
        root,
        "write",
        "--input", source.name,
        "--output", output.name,
        "--sheet", sheet_name,
        "--start-cell", "ZZ1",
        "--rows-json", json.dumps([["VBA archive preservation probe"]]),
    )
    assert result["ok"] and result["source_preserved"], result
    assert source.exists() and sha256(source) == original_hash, "Original XLSM changed"

    with zipfile.ZipFile(output) as archive:
        assert "xl/vbaProject.bin" in archive.namelist(), "VBA project part disappeared"
        copied_vba = archive.read("xl/vbaProject.bin")
        assert copied_vba == original_vba, "VBA project bytes changed during write-copy"

    reopened = load_workbook(output, keep_vba=True, data_only=False)
    try:
        assert reopened[sheet_name]["ZZ1"].value == "VBA archive preservation probe"
    finally:
        reopened.close()

    print("XLSM_VBA_BINARY_PASSTHROUGH=SUCCESS")
    print("XLSM_SOURCE_UNCHANGED=SUCCESS")


def test_corrupt_workbook_fails_safely(skill: Path, root: Path) -> None:
    bad = root / "corrupt.xlsx"
    bad.write_bytes(b"This is not an Excel ZIP workbook.")
    result = run(
        skill / "scripts" / "excel_workbook.py",
        root,
        "inspect",
        "--path", bad.name,
        expect_success=False,
    )
    assert result.get("ok") is False, result
    assert "valid" in result.get("error", "").lower(), result
    print("CORRUPT_WORKBOOK_SAFE_REJECTION=SUCCESS")


def main() -> int:
    if len(sys.argv) != 3:
        raise SystemExit(
            "Usage: test_workbook_compatibility.py EXTRACTED_SKILL_FOLDER REAL_MACRO_FIXTURE.xlsm"
        )
    skill = Path(sys.argv[1]).resolve(strict=True)
    macro_fixture = Path(sys.argv[2]).resolve(strict=True)
    script = skill / "scripts" / "excel_workbook.py"
    if not script.is_file():
        raise SystemExit(f"Missing Skill script: {script}")
    sys.path.insert(0, str(skill / "vendor"))

    with tempfile.TemporaryDirectory(prefix="workbuddy-excel-compat-") as temp:
        root = Path(temp)
        test_feature_matrix(skill)
        test_formula_cache_behavior(skill, root)
        test_real_xlsm_vba_passthrough(skill, root, macro_fixture)
        test_corrupt_workbook_fails_safely(skill, root)

    print("WORKBOOK_COMPATIBILITY_E2E=SUCCESS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
