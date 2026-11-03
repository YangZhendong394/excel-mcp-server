#!/usr/bin/env python3
"""Create and validate the small uploadable WorkBuddy Skill ZIP."""

from __future__ import annotations

import hashlib
import json
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SKILL_ROOT = ROOT / "workbuddy-skill" / "excel-workbook-assistant"
OUTPUT_DIR = ROOT / "dist"
MAX_BYTES = 3 * 1024 * 1024
REQUIRED = {
    "SKILL.md",
    "README.md",
    "THIRD_PARTY_NOTICES.md",
    "LICENSE",
    "runtime-lock.json",
    "scripts/start_server.py",
    "scripts/excel_mcp_bridge.py",
}


def main() -> int:
    if not SKILL_ROOT.is_dir():
        raise FileNotFoundError(SKILL_ROOT)
    files = sorted(
        path
        for path in SKILL_ROOT.rglob("*")
        if path.is_file() and "__pycache__" not in path.parts and path.suffix != ".pyc"
    )
    relative = {path.relative_to(SKILL_ROOT).as_posix() for path in files}
    missing = sorted(REQUIRED - relative)
    if missing:
        raise RuntimeError(f"The Skill package is missing required files: {missing}")
    if any("runtime-cache" in path.parts for path in files):
        raise RuntimeError("A local runtime cache must never be included in the Skill ZIP.")

    OUTPUT_DIR.mkdir(exist_ok=True)
    archive_path = OUTPUT_DIR / "excel-workbook-assistant-1.0.0.zip"
    checksum_path = OUTPUT_DIR / f"{archive_path.name}.sha256"
    archive_path.unlink(missing_ok=True)
    checksum_path.unlink(missing_ok=True)

    with zipfile.ZipFile(
        archive_path,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=9,
    ) as archive:
        for path in files:
            archive.write(
                path,
                arcname=f"excel-workbook-assistant/{path.relative_to(SKILL_ROOT).as_posix()}",
            )

    size = archive_path.stat().st_size
    if size >= MAX_BYTES:
        raise RuntimeError(f"Skill ZIP is {size} bytes; limit is {MAX_BYTES} bytes.")
    digest = hashlib.sha256(archive_path.read_bytes()).hexdigest()
    checksum_path.write_text(f"{digest}  {archive_path.name}\n", encoding="utf-8")

    with zipfile.ZipFile(archive_path) as archive:
        names = set(archive.namelist())
        expected = {
            f"excel-workbook-assistant/{path.relative_to(SKILL_ROOT).as_posix()}"
            for path in files
        }
        if names != expected:
            raise RuntimeError("ZIP content check failed.")

    summary = {
        "archive": archive_path.name,
        "size_bytes": size,
        "max_bytes": MAX_BYTES,
        "sha256": digest,
        "files": len(files),
    }
    print(json.dumps(summary, indent=2), flush=True)
    print("WORKBUDDY_SKILL_PACKAGE=SUCCESS", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
