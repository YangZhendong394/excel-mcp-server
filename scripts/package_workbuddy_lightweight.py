#!/usr/bin/env python3
"""Build a self-contained, offline WorkBuddy Skill ZIP from pinned pure-Python deps."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SKILL_SOURCE = ROOT / "workbuddy-skill" / "excel-workbook-assistant-lite"
PACKAGE_NAME = "excel-workbook-assistant-lite"
MAX_ZIP_BYTES = 3 * 1024 * 1024
VENDORED_PACKAGES = ("openpyxl", "et_xmlfile")
PINNED_VERSIONS = {"openpyxl": "3.1.5", "et_xmlfile": "2.0.0"}


def copy_vendor(source: Path, destination: Path) -> list[dict[str, str]]:
    copied: list[dict[str, str]] = []
    for package in VENDORED_PACKAGES:
        package_dir = source / package
        if not package_dir.is_dir():
            raise RuntimeError(f"Missing dependency package directory: {package_dir}")
        shutil.copytree(
            package_dir,
            destination / package,
            ignore=shutil.ignore_patterns("__pycache__", "*.pyc", "*.pyo", "*.dist-info"),
        )
        copied.append({"name": package, "version": PINNED_VERSIONS[package]})

    dist_infos = list(source.glob("*.dist-info"))
    normalize = lambda value: re.sub(r"[-_.]+", "-", value).lower()
    license_targets = {
        "openpyxl": destination.parent / "vendor-licenses" / "openpyxl-LICENCE.rst",
        "et_xmlfile": destination.parent / "vendor-licenses" / "et-xmlfile-LICENSE",
    }
    for package, target in license_targets.items():
        matching = [
            candidate
            for dist in dist_infos
            for candidate in dist.rglob("*")
            if candidate.is_file()
            and candidate.name.lower() in {"license", "licence", "license.txt", "licence.txt", "license.rst", "licence.rst"}
            and normalize(package) in normalize(dist.name)
        ]
        if matching:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(matching[0], target)
        else:
            raise RuntimeError(
                f"The installed {package} distribution did not provide a licence file; "
                "refusing to package vendored source without its licence."
            )
    return copied


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--vendor-dir", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()

    vendor_source = args.vendor_dir.resolve(strict=True)
    output = args.output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    stage = output.parent / f"{PACKAGE_NAME}-package"
    if stage.exists():
        shutil.rmtree(stage)
    stage.mkdir(parents=True)

    try:
        for required in ("SKILL.md", "README.md", "THIRD_PARTY_NOTICES.md", "FEATURE-COMPATIBILITY.md"):
            shutil.copy2(SKILL_SOURCE / required, stage / required)
        (stage / "scripts").mkdir(parents=True, exist_ok=True)
        shutil.copy2(SKILL_SOURCE / "scripts" / "excel_workbook.py", stage / "scripts" / "excel_workbook.py")
        shutil.copy2(ROOT / "LICENSE", stage / "LICENSE")
        vendor_dest = stage / "vendor"
        vendor_dest.mkdir()
        versions = copy_vendor(vendor_source, vendor_dest)

        manifest = {
            "skill": PACKAGE_NAME,
            "version": "2.0.0-alpha.1",
            "python": "3.13+",
            "network_required_at_runtime": False,
            "runtime_install_required": False,
            "dependencies": versions,
        }
        (stage / "BUILD-MANIFEST.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )

        if output.exists():
            output.unlink()
        with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
            for path in sorted(stage.rglob("*")):
                if path.is_file():
                    archive.write(
                        path,
                        f"{PACKAGE_NAME}/{path.relative_to(stage).as_posix()}",
                    )

        with zipfile.ZipFile(output) as archive:
            expected = {
                f"{PACKAGE_NAME}/{path.relative_to(stage).as_posix()}"
                for path in stage.rglob("*")
                if path.is_file()
            }
            if set(archive.namelist()) != expected:
                raise RuntimeError("Skill ZIP content or top-level folder structure check failed.")

        size = output.stat().st_size
        if size >= MAX_ZIP_BYTES:
            raise RuntimeError(
                f"Skill ZIP is {size:,} bytes; it must be strictly smaller than {MAX_ZIP_BYTES:,} bytes."
            )
        print(
            json.dumps(
                {
                    "ok": True,
                    "output": str(output),
                    "size_bytes": size,
                    "limit_bytes": MAX_ZIP_BYTES,
                    "sha256": sha256(output),
                    "dependencies": versions,
                    "file_count": sum(1 for path in stage.rglob("*") if path.is_file()),
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0
    finally:
        shutil.rmtree(stage, ignore_errors=True)


if __name__ == "__main__":
    raise SystemExit(main())
