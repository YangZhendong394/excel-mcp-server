#!/usr/bin/env python3
"""Build a platform-specific, hash-verified wheelhouse from runtime-lock.json.

The script downloads only the single best compatible wheel for each runtime
dependency, plus this project's already-built wheel. It never builds a dependency
from an sdist.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import platform
import shutil
import sys
import tempfile
import time
import urllib.error
import urllib.request
import zipfile
from pathlib import Path
from typing import Any

from pip._vendor.packaging.tags import sys_tags
from pip._vendor.packaging.utils import parse_wheel_filename

from start_server import package_is_applicable


ROOT = Path(__file__).resolve().parents[1]
LOCK_PATH = ROOT / "runtime-lock.json"
LICENSE_PATH = ROOT / "LICENSE"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def filename_from_url(url: str) -> str:
    return url.split("?", 1)[0].rstrip("/").rsplit("/", 1)[-1]


def choose_wheel(package: dict[str, Any], supported_tags: dict) -> dict[str, Any]:
    candidates = []
    for artifact in package.get("artifacts", []):
        if artifact.get("kind") != "wheel":
            continue
        filename = filename_from_url(artifact["url"])
        try:
            _name, _version, _build, wheel_tags = parse_wheel_filename(filename)
        except Exception as exc:
            raise ValueError(f"Cannot parse locked wheel name {filename!r}: {exc}") from exc

        matching = [tag for tag in wheel_tags if tag in supported_tags]
        if matching:
            # Lower rank is preferred because sys_tags() yields tags in priority order.
            best_rank = min(supported_tags[tag] for tag in matching)
            candidates.append((best_rank, artifact))

    if not candidates:
        raise RuntimeError(
            f"No prebuilt wheel in runtime-lock.json matches "
            f"{package['name']}=={package['version']} on "
            f"{sys.platform}/{platform.machine()}/Python {platform.python_version()}."
        )

    candidates.sort(key=lambda item: item[0])
    return candidates[0][1]


def download_verified(artifact: dict[str, Any], destination: Path) -> dict[str, Any]:
    filename = filename_from_url(artifact["url"])
    target = destination / filename
    if target.exists():
        raise FileExistsError(f"Refusing to overwrite existing wheel: {target}")

    request = urllib.request.Request(
        artifact["url"],
        headers={"User-Agent": "WorkBuddy-Excel-MCP-Wheelhouse/1.0"},
    )
    last_error: Exception | None = None

    for attempt in range(3):
        try:
            with urllib.request.urlopen(request, timeout=90) as response, target.open("wb") as out:
                shutil.copyfileobj(response, out, length=1024 * 1024)
            break
        except (OSError, urllib.error.URLError) as exc:
            last_error = exc
            target.unlink(missing_ok=True)
            if attempt == 2:
                raise RuntimeError(f"Download failed for {filename}: {exc}") from exc
            time.sleep(2 ** attempt)

    actual_hash = sha256_file(target)
    expected_hash = artifact["sha256"].lower()
    if actual_hash != expected_hash:
        target.unlink(missing_ok=True)
        raise RuntimeError(
            f"SHA-256 mismatch for {filename}: expected {expected_hash}, got {actual_hash}"
        )

    actual_size = target.stat().st_size
    expected_size = artifact.get("size")
    if expected_size is not None and actual_size != expected_size:
        target.unlink(missing_ok=True)
        raise RuntimeError(
            f"Size mismatch for {filename}: expected {expected_size}, got {actual_size}"
        )

    return {
        "filename": filename,
        "sha256": actual_hash,
        "size": actual_size,
        "source_url": artifact["url"],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dist-dir", type=Path, default=ROOT / "dist")
    parser.add_argument("--output-dir", type=Path, default=ROOT / "dist")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="select compatible locked wheels without downloading or writing files",
    )
    args = parser.parse_args()

    if sys.version_info < (3, 11):
        raise RuntimeError("The wheelhouse builder requires Python 3.11 or newer.")
    if not LOCK_PATH.is_file():
        raise FileNotFoundError(f"Runtime lock not found: {LOCK_PATH}")

    lock = json.loads(LOCK_PATH.read_text(encoding="utf-8"))
    tags = {tag: index for index, tag in enumerate(sys_tags())}
    selected_packages = [
        package
        for package in lock["packages"]
        if package_is_applicable(
            package,
            {
                "sys_platform": sys.platform,
                "python_version": f"{sys.version_info.major}.{sys.version_info.minor}",
                "python_full_version": platform.python_version(),
                "platform_python_implementation": platform.python_implementation(),
                "implementation_name": sys.implementation.name,
            },
        )
    ]
    choices = [
        (package, choose_wheel(package, tags))
        for package in selected_packages
    ]

    project_wheels = sorted(args.dist_dir.glob("*.whl"))
    project_name = lock["distribution"]["pypi_name"].replace("-", "_").lower()
    project_version = lock["distribution"]["version"]
    matching_project_wheels = [
        path for path in project_wheels
        if path.name.lower().startswith(f"{project_name}-{project_version}-")
    ]
    if len(matching_project_wheels) != 1:
        raise RuntimeError(
            f"Expected exactly one built project wheel for "
            f"{project_name}=={project_version} in {args.dist_dir}, "
            f"found {len(matching_project_wheels)}."
        )

    if args.dry_run:
        print("WHEELHOUSE_DRY_RUN=OK")
        print(f"platform={sys.platform}/{platform.machine()}")
        print(f"python={platform.python_version()}")
        print(f"dependency_count={len(choices)}")
        print(f"project_wheel={matching_project_wheels[0].name}")
        for package, artifact in choices:
            print(f"{package['name']}=={package['version']} -> {filename_from_url(artifact['url'])}")
        return 0

    output_dir = args.output_dir.resolve()
    if output_dir.exists() and any(output_dir.iterdir()):
        raise FileExistsError(
            f"Output directory is not empty: {output_dir}. Use a fresh directory."
        )
    output_dir.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="excel-mcp-wheelhouse-") as temp_name:
        temp_dir = Path(temp_name)
        manifest_files = []
        for package, artifact in choices:
            record = download_verified(artifact, temp_dir)
            record.update({"package": package["name"], "version": package["version"]})
            manifest_files.append(record)
            print(f"VERIFIED {package['name']}=={package['version']}", flush=True)

        project_wheel = matching_project_wheels[0]
        project_copy = temp_dir / project_wheel.name
        shutil.copy2(project_wheel, project_copy)
        manifest_files.append(
            {
                "filename": project_copy.name,
                "sha256": sha256_file(project_copy),
                "size": project_copy.stat().st_size,
                "package": lock["distribution"]["pypi_name"],
                "version": lock["distribution"]["version"],
                "source": "built-from-this-repository",
            }
        )

        for item in manifest_files:
            shutil.copy2(temp_dir / item["filename"], output_dir / item["filename"])

    wheelhouse_manifest = {
        "schema_version": 1,
        "distribution": lock["distribution"],
        "platform": {
            "system": platform.system(),
            "sys_platform": sys.platform,
            "machine": platform.machine(),
            "python": platform.python_version(),
            "implementation": platform.python_implementation(),
        },
        "upstream": {
            "repository": "https://github.com/haris-musa/excel-mcp-server",
            "license": "MIT",
            "copyright": "Copyright (c) 2025-2026 Haris Musa",
        },
        "file_count": len(manifest_files),
        "files": sorted(manifest_files, key=lambda item: item["filename"].lower()),
    }
    (output_dir / "wheelhouse-manifest.json").write_text(
        json.dumps(wheelhouse_manifest, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    shutil.copy2(LICENSE_PATH, output_dir / "LICENSE")

    system = sys.platform
    machine = platform.machine().lower()
    py_tag = f"cp{sys.version_info.major}{sys.version_info.minor}"
    archive_name = (
        f"wheelhouse-{system}-{machine}-{py_tag}-"
        f"{lock['distribution']['version']}.zip"
    )
    archive_path = output_dir.parent / archive_name
    if archive_path.exists():
        raise FileExistsError(f"Refusing to overwrite existing archive: {archive_path}")

    with zipfile.ZipFile(
        archive_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6
    ) as archive:
        for path in sorted(output_dir.iterdir()):
            if path.is_file():
                archive.write(path, arcname=path.name)

    print(f"WHEELHOUSE_BUILD=SUCCESS")
    print(f"dependency_count={len(choices)}")
    print(f"file_count={len(manifest_files)}")
    print(f"archive={archive_path}")
    print(f"archive_size_bytes={archive_path.stat().st_size}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
