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

    print("LIGHTWEIGHT_SKILL_E2E=SUCCESS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
