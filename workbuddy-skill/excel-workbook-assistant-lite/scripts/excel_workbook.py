#!/usr/bin/env python3
"""Offline Excel workbook utilities for the WorkBuddy lightweight Skill.

The package vendors openpyxl and et_xmlfile beside this script. No installation,
network access, MCP process, or modification of the user's source workbook is required.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import sys
import tempfile
import zipfile
from decimal import Decimal
from pathlib import Path
from typing import Any

# WorkBuddy runs on Windows for the current product target. Force UTF-8 JSON output
# so Chinese worksheet names and cell values are never lost on legacy console code pages.
for _stream in (sys.stdout, sys.stderr):
    reconfigure = getattr(_stream, "reconfigure", None)
    if reconfigure is not None:
        reconfigure(encoding="utf-8", errors="replace")

SKILL_ROOT = Path(__file__).resolve().parents[1]
VENDOR_DIR = SKILL_ROOT / "vendor"
if VENDOR_DIR.is_dir():
    sys.path.insert(0, str(VENDOR_DIR))

try:
    import openpyxl
    from openpyxl import Workbook, load_workbook
    from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries
except ImportError as exc:  # pragma: no cover - exercised by packaging and runtime checks
    print(
        json.dumps(
            {
                "ok": False,
                "error": "Bundled Excel libraries are missing or damaged.",
                "detail": str(exc),
            },
            ensure_ascii=False,
        ),
        file=sys.stderr,
    )
    raise SystemExit(2) from exc


MAX_FILE_BYTES = 100 * 1024 * 1024
MAX_UNPACKED_BYTES = 1024 * 1024 * 1024
MAX_COMPRESSION_RATIO = 500
MAX_CELLS_PER_CALL = 50_000
MAX_CELL_TEXT = 32_767
ALLOWED_SUFFIXES = {".xlsx", ".xlsm"}
ILLEGAL_XML_TEXT = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")


class SkillError(Exception):
    """An expected, user-safe validation or workbook error."""


def json_value(value: Any) -> Any:
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    if isinstance(value, dt.timedelta):
        return str(value)
    if isinstance(value, Decimal):
        return str(value)
    if isinstance(value, bytes):
        return value.hex()
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    return str(value)


def emit(payload: dict[str, Any]) -> None:
    print(json.dumps(payload, ensure_ascii=False, default=json_value, indent=2))


def resolve_path(root: Path, raw: str, *, must_exist: bool) -> Path:
    base = root.expanduser().resolve(strict=True)
    candidate = Path(raw).expanduser()
    if not candidate.is_absolute():
        candidate = base / candidate
    resolved = candidate.resolve(strict=must_exist)
    try:
        resolved.relative_to(base)
    except ValueError:
        raise SkillError("The requested path is outside the allowed folder.") from None
    if resolved == base:
        raise SkillError("A file path is required, not the allowed folder itself.")
    if resolved.suffix.lower() not in ALLOWED_SUFFIXES:
        raise SkillError("Only .xlsx and .xlsm workbooks are supported in this version.")
    if must_exist:
        if not resolved.is_file():
            raise SkillError("The requested workbook does not exist or is not a file.")
        if resolved.stat().st_size > MAX_FILE_BYTES:
            raise SkillError(f"The workbook exceeds the {MAX_FILE_BYTES // (1024 * 1024)} MB file limit.")
    else:
        if resolved.exists():
            raise SkillError("The output file already exists. Choose a new filename; no file was overwritten.")
        if not resolved.parent.is_dir():
            raise SkillError("The output folder does not exist.")
    return resolved


def validate_xlsx_container(path: Path) -> None:
    if not zipfile.is_zipfile(path):
        raise SkillError("The file is not a valid unencrypted Excel Open XML workbook.")
    try:
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            names = {item.filename for item in infos}
            if "xl/workbook.xml" not in names:
                raise SkillError("The ZIP file does not contain an Excel workbook.")
            total = 0
            for item in infos:
                total += item.file_size
                if item.file_size and item.compress_size == 0:
                    raise SkillError("The workbook contains a suspicious compressed entry.")
                if item.compress_size and item.file_size / item.compress_size > MAX_COMPRESSION_RATIO:
                    raise SkillError("The workbook contains an unusually high compression ratio.")
                if total > MAX_UNPACKED_BYTES:
                    raise SkillError("The workbook expands beyond the safety limit.")
    except (OSError, zipfile.BadZipFile, RuntimeError) as exc:
        raise SkillError(f"The workbook container could not be checked: {exc}") from None


def load_checked(path: Path, *, data_only: bool = False):
    validate_xlsx_container(path)
    keep_vba = path.suffix.lower() == ".xlsm"
    try:
        return load_workbook(
            path,
            read_only=False,
            data_only=data_only,
            keep_vba=keep_vba,
            keep_links=True,
        )
    except Exception as exc:
        raise SkillError(f"Excel workbook could not be opened: {type(exc).__name__}: {exc}") from None


def save_new(workbook: Workbook, output: Path) -> None:
    """Save atomically without replacing an existing output file."""
    fd, temporary_name = tempfile.mkstemp(
        prefix=f".{output.stem}-",
        suffix=output.suffix,
        dir=str(output.parent),
    )
    os.close(fd)
    temporary = Path(temporary_name)
    try:
        workbook.save(temporary)
        if output.exists():
            raise SkillError("The output file appeared during processing; it was not overwritten.")
        try:
            # Hard-link creation is atomic and fails rather than replacing an existing path.
            os.link(temporary, output)
            temporary.unlink()
        except OSError:
            # Some filesystems do not support hard links. Reserve the destination with O_EXCL.
            if output.exists():
                raise SkillError("The output file already exists; it was not overwritten.") from None
            reserve_fd = None
            try:
                reserve_fd = os.open(
                    str(output),
                    os.O_CREAT | os.O_EXCL | os.O_WRONLY,
                    0o600,
                )
                os.close(reserve_fd)
                reserve_fd = None
                os.replace(temporary, output)
            except FileExistsError:
                raise SkillError("The output file already exists; it was not overwritten.") from None
            finally:
                if reserve_fd is not None:
                    os.close(reserve_fd)
    finally:
        if temporary.exists():
            temporary.unlink()


def command_create(args: argparse.Namespace, root: Path) -> dict[str, Any]:
    output = resolve_path(root, args.path, must_exist=False)
    names = [name.strip() for name in args.sheets.split(",") if name.strip()]
    if not names:
        names = ["Sheet1"]
    if len(names) > 50:
        raise SkillError("A workbook may contain at most 50 sheets in this operation.")
    if len({name.casefold() for name in names}) != len(names):
        raise SkillError("Worksheet names must be unique.")
    for name in names:
        if len(name) > 31 or any(ch in name for ch in "[]:*?/\\"):
            raise SkillError(f"Invalid worksheet name: {name!r}")
    workbook = Workbook()
    workbook.active.title = names[0]
    for name in names[1:]:
        workbook.create_sheet(name)
    try:
        save_new(workbook, output)
    finally:
        workbook.close()
    return {"ok": True, "operation": "create", "path": str(output), "sheets": names}


def command_inspect(args: argparse.Namespace, root: Path) -> dict[str, Any]:
    path = resolve_path(root, args.path, must_exist=True)
    workbook = load_checked(path)
    try:
        sheets = []
        for sheet in workbook.worksheets:
            sheets.append(
                {
                    "name": sheet.title,
                    "state": sheet.sheet_state,
                    "max_row": sheet.max_row,
                    "max_column": sheet.max_column,
                    "dimension": sheet.calculate_dimension(),
                    "merged_ranges": len(sheet.merged_cells.ranges),
                    "tables": list(sheet.tables.keys()),
                    "charts": len(sheet._charts),
                    "freeze_panes": str(sheet.freeze_panes) if sheet.freeze_panes else None,
                }
            )
        return {
            "ok": True,
            "operation": "inspect",
            "path": str(path),
            "file_bytes": path.stat().st_size,
            "openpyxl_version": openpyxl.__version__,
            "sheets": sheets,
        }
    finally:
        workbook.close()


def command_read(args: argparse.Namespace, root: Path) -> dict[str, Any]:
    path = resolve_path(root, args.path, must_exist=True)
    workbook = load_checked(path, data_only=args.values_only)
    try:
        if args.sheet not in workbook.sheetnames:
            raise SkillError(f"Worksheet not found: {args.sheet}")
        try:
            min_col, min_row, max_col, max_row = range_boundaries(args.range)
        except ValueError:
            raise SkillError("Invalid cell range. Use A1 notation such as A1:D20.") from None
        if None in (min_col, min_row, max_col, max_row):
            raise SkillError("A bounded cell range is required, for example A1:D20.")
        cell_count = (max_row - min_row + 1) * (max_col - min_col + 1)
        if cell_count > MAX_CELLS_PER_CALL:
            raise SkillError(f"The requested range exceeds {MAX_CELLS_PER_CALL:,} cells.")
        sheet = workbook[args.sheet]
        rows = [
            [json_value(cell.value) for cell in row]
            for row in sheet.iter_rows(
                min_row=min_row,
                max_row=max_row,
                min_col=min_col,
                max_col=max_col,
            )
        ]
        return {
            "ok": True,
            "operation": "read",
            "path": str(path),
            "sheet": args.sheet,
            "range": args.range,
            "values": rows,
            "values_only": args.values_only,
        }
    finally:
        workbook.close()


def command_write(args: argparse.Namespace, root: Path) -> dict[str, Any]:
    source = resolve_path(root, args.input, must_exist=True)
    output = resolve_path(root, args.output, must_exist=False)
    if source == output:
        raise SkillError("Input and output paths must be different. The source workbook is never overwritten.")
    if source.suffix.lower() != output.suffix.lower():
        raise SkillError("Input and output extensions must match.")
    if args.sheet == "":
        raise SkillError("A worksheet name is required.")
    try:
        rows = json.loads(args.rows_json)
    except json.JSONDecodeError as exc:
        raise SkillError(f"Invalid JSON in --rows-json: {exc}") from None
    if not isinstance(rows, list) or not rows or any(not isinstance(row, list) for row in rows):
        raise SkillError("The rows payload must be a non-empty JSON array of row arrays.")
    cell_count = sum(len(row) for row in rows)
    if cell_count == 0 or cell_count > MAX_CELLS_PER_CALL:
        raise SkillError(f"The write must contain between 1 and {MAX_CELLS_PER_CALL:,} cells.")
    for row in rows:
        for value in row:
            if isinstance(value, (dict, list)):
                raise SkillError("Cell values must be strings, numbers, booleans, or null.")
            if isinstance(value, str):
                if len(value) > MAX_CELL_TEXT:
                    raise SkillError(f"Cell text exceeds Excel's {MAX_CELL_TEXT:,}-character limit.")
                if ILLEGAL_XML_TEXT.search(value):
                    raise SkillError("Cell text contains characters that Excel XML cannot store.")
    try:
        start_row, start_col = coordinate_to_tuple(args.start_cell)
    except Exception:
        raise SkillError("Invalid starting cell. Use a cell such as A1 or D5.") from None

    workbook = load_checked(source)
    try:
        if args.sheet not in workbook.sheetnames:
            raise SkillError(f"Worksheet not found: {args.sheet}")
        sheet = workbook[args.sheet]
        written = 0
        for row_offset, row_values in enumerate(rows):
            for col_offset, value in enumerate(row_values):
                cell = sheet.cell(start_row + row_offset, start_col + col_offset)
                cell.value = value
                if isinstance(value, str) and value.startswith("=") and not args.allow_formulas:
                    cell.data_type = "s"
                written += 1
        save_new(workbook, output)
        return {
            "ok": True,
            "operation": "write-copy",
            "input": str(source),
            "output": str(output),
            "sheet": args.sheet,
            "start_cell": args.start_cell,
            "cells_written": written,
            "formulas_allowed": args.allow_formulas,
            "source_preserved": True,
            "note": (
                "Formulas are stored but are not calculated by this script."
                if args.allow_formulas
                else "Strings beginning with '=' were stored as text."
            ),
        }
    finally:
        workbook.close()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--allow-dir", required=True, type=Path, help="Only folder the script may access.")
    commands = parser.add_subparsers(dest="command", required=True)

    create = commands.add_parser("create", help="Create a new workbook without overwriting files.")
    create.add_argument("--path", required=True)
    create.add_argument("--sheets", default="Sheet1")

    inspect = commands.add_parser("inspect", help="Inspect workbook structure.")
    inspect.add_argument("--path", required=True)

    read = commands.add_parser("read", help="Read a bounded cell range.")
    read.add_argument("--path", required=True)
    read.add_argument("--sheet", required=True)
    read.add_argument("--range", required=True)
    read.add_argument("--values-only", action="store_true")

    write = commands.add_parser("write", help="Write cells to a new copy; the input is never overwritten.")
    write.add_argument("--input", required=True)
    write.add_argument("--output", required=True)
    write.add_argument("--sheet", required=True)
    write.add_argument("--start-cell", required=True)
    write.add_argument("--rows-json", required=True)
    write.add_argument("--allow-formulas", action="store_true")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    try:
        if not args.allow_dir.exists() or not args.allow_dir.is_dir():
            raise SkillError("The allowed folder does not exist or is not a directory.")
        root = args.allow_dir.expanduser().resolve(strict=True)
        if args.command == "create":
            result = command_create(args, root)
        elif args.command == "inspect":
            result = command_inspect(args, root)
        elif args.command == "read":
            result = command_read(args, root)
        elif args.command == "write":
            result = command_write(args, root)
        else:  # pragma: no cover
            raise SkillError("Unknown command.")
        emit(result)
        return 0
    except SkillError as exc:
        emit({"ok": False, "error": str(exc)})
        return 2
    except Exception as exc:  # fail closed while keeping tracebacks out of workbook-facing output
        emit({"ok": False, "error": f"Unexpected {type(exc).__name__}: {exc}"})
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
