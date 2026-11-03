#!/usr/bin/env python3
"""Build a platform-specific, hash-verified wheelhouse from runtime-lock.json.

Only the best compatible prebuilt wheel for each runtime dependency is
downloaded. Dependencies are never built from source distributions.
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
            raise ValueError(
                f"Cannot parse locked wheel name {filename!r}: {exc}"
            ) from exc

        matching = [tag for tag in wheel_tags if tag in supported_tags]
        if matching:
            # sys_tags() yields tags in preference order.
            best_rank = min(supported_tags[tag] for tag in matching)
            candidates.append((best_rank, artifact))

    if not candidates:
        raise RuntimeError(
            f"No prebuilt locked wheel matches {package['name']}=={package['version']} "
            f"on {sys.platform}/{platform.machine()}/Python {platform.python_version()}."
        )

    candidates.sort(key=lambda item: item[0])
    return candidates[0][1]


def download_verified(
    artifact: dict[str, Any],
    destination: Path,
) -> dict[str, Any]:
    filename = filename_from_url(artifact["url"])
    target = destination / filename
    if target.exists():
        raise FileExistsError(f"Refusing to overwrite existing wheel: {target}")

    request = urllib.request.Request(
        artifact["url"],
        headers={"User-Agent": "WorkBuddy-Excel-MCP-Wheelhouse/1.0"},
    )

    for attempt in range(3):
        try:
            with (
                urllib.request.urlopen(request, timeout=90) as response,
                target.open("wb") as output,
            ):
                shutil.copyfileobj(response, output, length=1024 * 1024)
            break
        except (OSError, urllib.error.URLError) as exc:
            target.unlink(missing_ok=True)
            if attempt == 2:
                raise RuntimeError(
                    f"Download failed for {filename}: {exc}"
                ) from exc
            time.sleep(2**attempt)

    actual_hash = sha256_file(target)
    expected_hash = artifact["sha256"].lower()
    if actual_hash != expected_hash:
        target.unlink(missing_ok=True)
        raise RuntimeError(
            f"SHA-256 mismatch for {filename}: "
            f"expected {expected_hash}, got {actual_hash}"
        )

    actual_size = target.stat().st_size
    expected_size = artifact.get("size")
    if expected_size is not None and actual_size != expected_size:
        target.unlink(missing_ok=True)
        raise RuntimeError(
            f"Size mismatch for {filename}: "
            f"expected {expected_size}, got {actual_size}"
        )

    return {
        "filename": filename,
        "sha256": actual_hash,
        "size": actual_size,
        "source_url": artifact["url"],
    }


def load_lock() -> dict[str, Any]:
    if not LOCK_PATH.is_file():
        raise FileNotFoundError(f"Runtime lock not found: {LOCK_PATH}")
    return json.loads(LOCK_PATH.read_text(encoding="utf-8"))


def select_runtime_wheels(lock: dict[str, Any]) -> list[tuple[dict, dict]]:
    supported_tags = {tag: index for index, tag in enumerate(sys_tags())}
    environment = {
        "sys_platform": sys.platform,
        "python_version": f"{sys.version_info.major}.{sys.version_info.minor}",
        "python_full_version": platform.python_version(),
        "platform_python_implementation": platform.python_implementation(),
        "implementation_name": sys.implementation.name,
    }
    packages = [
        package
        for package in lock["packages"]
        if package_is_applicable(package, environment)
    ]
    return [
        (package, choose_wheel(package, supported_tags))
        for package in packages
    ]


def find_project_wheel(
    dist_dir: Path,
    lock: dict[str, Any],
) -> Path:
    package_name = lock["distribution"]["pypi_name"].replace("-", "_").lower()
    package_version = lock["distribution"]["version"]
    matches = [
        path
        for path in sorted(dist_dir.glob("*.whl"))
        if path.name.lower().startswith(f"{package_name}-{package_version}-")
    ]
    if len(matches) != 1:
        raise RuntimeError(
            f"Expected exactly one built wheel for "
            f"{package_name}=={package_version} in {dist_dir}; found {len(matches)}."
        )
    return matches[0]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dist-dir", type=Path, default=ROOT / "dist")
    parser.add_argument("--output-dir", type=Path, default=ROOT / "dist" / "wheelhouse")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="select compatible locked wheels without downloading or writing files",
    )
    args = parser.parse_args()

    if sys.version_info < (3, 11):
        raise RuntimeError("The wheelhouse builder requires Python 3.11 or newer.")

    lock = load_lock()
    choices = select_runtime_wheels(lock)

    if args.dry_run:
        print("WHEELHOUSE_DRY_RUN=OK")
        print(f"platform={sys.platform}/{platform.machine()}")
        print(f"python={platform.python_version()}")
        print(f"dependency_count={len(choices)}")
        for package, artifact in choices:
            print(
                f"{package['name']}=={package['version']} -> "
                f"{filename_from_url(artifact['url'])}"
            )
        return 0

    project_wheel = find_project_wheel(args.dist_dir, lock)
    output_dir = args.output_dir.resolve()
    if output_dir.exists() and any(output_dir.iterdir()):
        raise FileExistsError(
            f"Output directory is not empty: {output_dir}. Use a fresh directory."
        )
    output_dir.mkdir(parents=True, exist_ok=True)

    manifest_files = []
    with tempfile.TemporaryDirectory(prefix="excel-mcp-wheelhouse-") as temp_name:
        temp_dir = Path(temp_name)
        for package, artifact in choices:
            record = download_verified(artifact, temp_dir)
            record.update(
                {"package": package["name"], "version": package["version"]}
            )
            manifest_files.append(record)
            print(
                f"VERIFIED {package['name']}=={package['version']}",
                flush=True,
            )

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
        "wheel_count": len(manifest_files),
        "files": sorted(
            manifest_files,
            key=lambda item: item["filename"].lower(),
        ),
    }
    manifest_path = output_dir / "wheelhouse-manifest.json"
    manifest_path.write_text(
        json.dumps(wheelhouse_manifest, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    shutil.copy2(LICENSE_PATH, output_dir / "LICENSE")

    system = {
        "win32": "windows",
        "darwin": "macos",
        "linux": "linux",
    }.get(sys.platform, sys.platform)
    raw_machine = platform.machine().lower()
    machine = {
        "amd64": "amd64",
        "x86_64": "amd64",
        "arm64": "arm64",
        "aarch64": "arm64",
    }.get(raw_machine, raw_machine)
    python_tag = f"cp{sys.version_info.major}{sys.version_info.minor}"
    archive_name = (
        f"wheelhouse-{system}-{machine}-{python_tag}-"
        f"{lock['distribution']['version']}.zip"
    )
    archive_path = output_dir.parent / archive_name
    if archive_path.exists():
        raise FileExistsError(f"Refusing to overwrite existing archive: {archive_path}")

    with zipfile.ZipFile(
        archive_path,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=6,
    ) as archive:
        for path in sorted(output_dir.iterdir()):
            if path.is_file():
                archive.write(path, arcname=path.name)

    print("WHEELHOUSE_BUILD=SUCCESS")
    print(f"dependency_count={len(choices)}")
    print(f"wheel_count={len(manifest_files)}")
    print(f"archive={archive_path}")
    print(f"archive_size_bytes={archive_path.stat().st_size}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
