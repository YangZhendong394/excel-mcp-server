#!/usr/bin/env python3
"""Launch the pinned Excel MCP Server from a verified GitHub Release wheelhouse.

The wheelhouse is downloaded once, verified against its release checksum and
runtime-lock.json, then cached locally. The actual uvx install is offline.
All server protocol traffic is inherited unchanged; diagnostics go to stderr.
"""
from __future__ import annotations

import hashlib
import json
import os
import platform
import re
import shutil
import stat
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
import zipfile
from pathlib import Path, PurePosixPath
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
LOCK_PATH = ROOT / "runtime-lock.json"
REPOSITORY = "YangZhendong394/excel-mcp-server"
GITHUB_RELEASES = f"https://github.com/{REPOSITORY}/releases/download"


def version_tuple(value: str) -> tuple[int, ...]:
    parts = re.findall(r"\d+", value)
    return tuple(int(part) for part in parts)


def marker_atom_matches(atom: str, environment: dict[str, str]) -> bool:
    atom = atom.strip()
    while atom.startswith("(") and atom.endswith(")"):
        atom = atom[1:-1].strip()

    match = re.fullmatch(
        r"([A-Za-z_][A-Za-z0-9_]*)\s*(==|!=|<=|>=|<|>)\s*['\"]([^'\"]+)['\"]",
        atom,
    )
    if not match:
        raise ValueError(f"Unsupported dependency marker in runtime-lock.json: {atom}")

    variable, operator, expected = match.groups()
    if variable not in environment:
        raise ValueError(f"Unknown dependency marker variable: {variable}")

    actual = environment[variable]
    if variable in {"python_full_version", "python_version"}:
        left, right = version_tuple(actual), version_tuple(expected)
        left = left + (0,) * (max(len(left), len(right)) - len(left))
        right = right + (0,) * (max(len(left), len(right)) - len(right))
    else:
        left, right = actual.lower(), expected.lower()

    return {
        "==": left == right,
        "!=": left != right,
        ">=": left >= right,
        "<=": left <= right,
        ">": left > right,
        "<": left < right,
    }[operator]


def marker_matches(expression: str, environment: dict[str, str]) -> bool:
    # uv.lock currently emits PEP 508 marker conjunctions for this project.
    # Split each expression as OR-of-AND clauses to support future lock entries.
    for disjunction in expression.split(" or "):
        clauses = [part.strip() for part in disjunction.split(" and ")]
        if all(marker_atom_matches(part, environment) for part in clauses):
            return True
    return False


def package_is_applicable(package: dict, environment: dict[str, str]) -> bool:
    paths = package.get("activation_paths", [[]])
    return any(
        all(marker_matches(marker, environment) for marker in path)
        for path in paths
    )


def current_environment() -> dict[str, str]:
    return {
        "sys_platform": sys.platform,
        "python_version": f"{sys.version_info.major}.{sys.version_info.minor}",
        "python_full_version": platform.python_version(),
        "platform_python_implementation": platform.python_implementation(),
        "implementation_name": sys.implementation.name,
    }


def normalized_name(value: str) -> str:
    return re.sub(r"[-_.]+", "-", value).lower()


def filename_from_url(url: str) -> str:
    return url.split("?", 1)[0].rstrip("/").rsplit("/", 1)[-1]


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def platform_identity() -> tuple[str, str, str]:
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
    return system, machine, python_tag


def release_asset_names(lock: dict[str, Any]) -> tuple[str, str, str]:
    system, machine, python_tag = platform_identity()
    version = lock["distribution"]["version"]
    release_tag = f"runtime-{system}-{machine}-{python_tag}-{version}"
    archive_name = f"wheelhouse-{system}-{machine}-{python_tag}-{version}.zip"
    return release_tag, archive_name, f"{archive_name}.sha256"


def cache_root() -> Path:
    if os.environ.get("LOCALAPPDATA"):
        base = Path(os.environ["LOCALAPPDATA"])
    elif os.environ.get("XDG_CACHE_HOME"):
        base = Path(os.environ["XDG_CACHE_HOME"])
    else:
        base = Path.home() / ".cache"
    return base / "WorkBuddy" / "excel-mcp-server"


def expected_packages(lock: dict[str, Any]) -> dict[str, dict[str, Any]]:
    environment = current_environment()
    result: dict[str, dict[str, Any]] = {}
    for package in lock["packages"]:
        if package_is_applicable(package, environment):
            key = normalized_name(package["name"])
            if key in result:
                raise ValueError(f"Duplicate runtime package in lock: {package['name']}")
            result[key] = package
    return result


def verify_wheelhouse(directory: Path, lock: dict[str, Any]) -> None:
    manifest_path = directory / "wheelhouse-manifest.json"
    if not manifest_path.is_file():
        raise RuntimeError("The runtime backup has no wheelhouse manifest.")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))

    distribution = lock["distribution"]
    if manifest.get("distribution", {}).get("pypi_name") != distribution["pypi_name"]:
        raise RuntimeError("The runtime backup belongs to a different package.")
    if manifest.get("distribution", {}).get("version") != distribution["version"]:
        raise RuntimeError("The runtime backup package version does not match the skill.")

    system, machine, _python_tag = platform_identity()
    manifest_platform = manifest.get("platform", {})
    if manifest_platform.get("sys_platform") != sys.platform:
        raise RuntimeError("The runtime backup was built for a different operating system.")
    manifest_machine = str(manifest_platform.get("machine", "")).lower()
    normalized_manifest_machine = {
        "amd64": "amd64",
        "x86_64": "amd64",
        "arm64": "arm64",
        "aarch64": "arm64",
    }.get(manifest_machine, manifest_machine)
    if normalized_manifest_machine != machine:
        raise RuntimeError("The runtime backup was built for a different processor type.")
    built_python = str(manifest_platform.get("python", ""))
    if not built_python.startswith(f"{sys.version_info.major}.{sys.version_info.minor}."):
        raise RuntimeError("The runtime backup was built for a different Python version.")
    if manifest_platform.get("implementation") != "CPython":
        raise RuntimeError("The runtime backup does not contain CPython wheels.")

    packages = expected_packages(lock)
    listed_files = manifest.get("files")
    if not isinstance(listed_files, list):
        raise RuntimeError("The runtime backup manifest has an invalid file list.")
    if manifest.get("wheel_count") != len(listed_files):
        raise RuntimeError("The runtime backup manifest has an incorrect wheel count.")

    seen_packages: set[str] = set()
    project_wheel_count = 0
    for item in listed_files:
        if not isinstance(item, dict):
            raise RuntimeError("The runtime backup manifest contains an invalid file entry.")
        filename = item.get("filename")
        if not isinstance(filename, str) or Path(filename).name != filename:
            raise RuntimeError("The runtime backup contains an unsafe wheel filename.")
        file_path = directory / filename
        if not file_path.is_file():
            raise RuntimeError(f"A runtime backup file is missing: {filename}")
        actual_hash = sha256_file(file_path)
        if actual_hash != str(item.get("sha256", "")).lower():
            raise RuntimeError(f"A runtime backup file failed its SHA-256 check: {filename}")
        if file_path.stat().st_size != item.get("size"):
            raise RuntimeError(f"A runtime backup file has an unexpected size: {filename}")

        if item.get("source") == "built-from-this-repository":
            project_wheel_count += 1
            if normalized_name(str(item.get("package", ""))) != normalized_name(
                distribution["pypi_name"]
            ):
                raise RuntimeError("The project wheel has an unexpected package name.")
            if item.get("version") != distribution["version"]:
                raise RuntimeError("The project wheel has an unexpected package version.")
            expected_prefix = (
                f"{normalized_name(distribution['pypi_name']).replace('-', '_')}-"
                f"{distribution['version']}-"
            )
            if not filename.lower().startswith(expected_prefix.lower()) or not filename.endswith(
                ".whl"
            ):
                raise RuntimeError("The project wheel filename is invalid.")
            continue

        package_key = normalized_name(str(item.get("package", "")))
        if package_key not in packages:
            raise RuntimeError(f"Unexpected dependency in runtime backup: {item.get('package')}")
        package = packages[package_key]
        if item.get("version") != package["version"]:
            raise RuntimeError(f"Unexpected dependency version for {package['name']}.")
        if package_key in seen_packages:
            raise RuntimeError(f"Duplicate dependency wheel for {package['name']}.")
        locked_artifacts = {
            (filename_from_url(artifact["url"]), artifact["sha256"].lower())
            for artifact in package.get("artifacts", [])
            if artifact.get("kind") == "wheel"
        }
        if (filename, actual_hash) not in locked_artifacts:
            raise RuntimeError(
                f"The dependency wheel is not the locked artifact for {package['name']}."
            )
        seen_packages.add(package_key)

    if seen_packages != set(packages):
        missing = sorted(set(packages) - seen_packages)
        raise RuntimeError(f"The runtime backup is missing dependencies: {', '.join(missing)}")
    if project_wheel_count != 1:
        raise RuntimeError("The runtime backup must contain exactly one project wheel.")

    allowed_files = {
        "wheelhouse-manifest.json",
        "LICENSE",
        *(str(item["filename"]) for item in listed_files),
    }
    actual_files = {path.name for path in directory.iterdir() if path.is_file()}
    if actual_files != allowed_files:
        raise RuntimeError("The runtime backup contains unexpected or missing files.")
    if not (directory / "LICENSE").is_file():
        raise RuntimeError("The runtime backup is missing the upstream MIT license.")


def download_file(url: str, destination: Path) -> None:
    request = urllib.request.Request(
        url,
        headers={"User-Agent": "WorkBuddy-Excel-MCP-Launcher/1.0"},
    )
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            with destination.open("wb") as output:
                shutil.copyfileobj(response, output, length=1024 * 1024)
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            raise RuntimeError(
                "No verified runtime backup has been published for this operating system, "
                "processor type, and Python version yet."
            ) from exc
        raise RuntimeError(f"Could not download the runtime backup (HTTP {exc.code}).") from exc
    except (OSError, urllib.error.URLError) as exc:
        raise RuntimeError(f"Could not download the runtime backup: {exc}") from exc


def extract_archive_safely(archive_path: Path, destination: Path) -> None:
    destination.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(archive_path) as archive:
        for info in archive.infolist():
            if info.is_dir():
                continue
            member = PurePosixPath(info.filename)
            if member.is_absolute() or len(member.parts) != 1 or any(
                part in {"", ".", ".."} for part in member.parts
            ):
                raise RuntimeError(f"Unsafe path in runtime backup: {info.filename}")
            mode = (info.external_attr >> 16) & 0o170000
            if mode == stat.S_IFLNK:
                raise RuntimeError(f"Symbolic links are not allowed in runtime backup: {info.filename}")
            target = destination / member.name
            with archive.open(info) as source, target.open("xb") as output:
                shutil.copyfileobj(source, output, length=1024 * 1024)


def ensure_wheelhouse(lock: dict[str, Any]) -> Path:
    system, machine, python_tag = platform_identity()
    release_tag, archive_name, checksum_name = release_asset_names(lock)
    target = cache_root() / f"{system}-{machine}-{python_tag}-{lock['distribution']['version']}"
    target.parent.mkdir(parents=True, exist_ok=True)

    if target.is_dir():
        try:
            verify_wheelhouse(target, lock)
            return target
        except (OSError, ValueError, KeyError, TypeError, RuntimeError, json.JSONDecodeError):
            shutil.rmtree(target, ignore_errors=True)

    stage = Path(tempfile.mkdtemp(prefix="excel-mcp-wheelhouse-", dir=target.parent))
    try:
        archive_path = stage / archive_name
        checksum_path = stage / checksum_name
        base_url = f"{GITHUB_RELEASES}/{release_tag}"
        download_file(f"{base_url}/{checksum_name}", checksum_path)
        checksum_text = checksum_path.read_text(encoding="utf-8").strip()
        checksum_match = re.fullmatch(
            rf"([0-9a-fA-F]{{64}})\s+\*?{re.escape(archive_name)}",
            checksum_text,
        )
        if not checksum_match:
            raise RuntimeError("The runtime backup checksum file is invalid.")
        expected_archive_hash = checksum_match.group(1).lower()

        download_file(f"{base_url}/{archive_name}", archive_path)
        if sha256_file(archive_path) != expected_archive_hash:
            raise RuntimeError("The runtime backup archive failed its SHA-256 check.")

        extracted = stage / "wheelhouse"
        extract_archive_safely(archive_path, extracted)
        verify_wheelhouse(extracted, lock)
        if target.exists():
            shutil.rmtree(target)
        extracted.replace(target)
    finally:
        shutil.rmtree(stage, ignore_errors=True)

    return target


def build_command(
    lock: dict[str, Any],
    uvx: str,
    wheelhouse: Path,
    server_args: list[str],
) -> tuple[list[str], int]:
    distribution = lock["distribution"]
    packages = expected_packages(lock)
    command = [
        uvx,
        "--offline",
        "--no-index",
        "--find-links",
        str(wheelhouse),
        "--from",
        f"{distribution['pypi_name']}=={distribution['version']}",
        distribution["console_command"],
    ]
    command.extend(server_args or ["stdio"])
    return command, len(packages)


def main() -> int:
    if sys.version_info < (3, 11):
        print("Excel MCP Server requires Python 3.11 or newer.", file=sys.stderr, flush=True)
        return 2

    if not LOCK_PATH.is_file():
        print(f"Runtime lock not found: {LOCK_PATH}", file=sys.stderr, flush=True)
        return 2

    uvx = shutil.which("uvx")
    if not uvx:
        print("uvx was not found in PATH; WorkBuddy's uv installation is required.", file=sys.stderr, flush=True)
        return 2

    lock = json.loads(LOCK_PATH.read_text(encoding="utf-8"))
    args = sys.argv[1:]
    show_command = bool(args and args[0] == "--show-command")
    if show_command:
        args = args[1:]

    system, machine, python_tag = platform_identity()
    release_tag, archive_name, _checksum_name = release_asset_names(lock)
    placeholder_wheelhouse = cache_root() / (
        f"{system}-{machine}-{python_tag}-{lock['distribution']['version']}"
    )
    command, selected = build_command(lock, uvx, placeholder_wheelhouse, args)

    if show_command:
        print(
            json.dumps(
                {
                    "runtime_source": "hash-verified GitHub Release wheelhouse",
                    "release_tag": release_tag,
                    "archive": archive_name,
                    "selected_dependency_count": selected,
                    "download_once_then_cache": True,
                    "command": command,
                },
                indent=2,
            )
        )
        return 0

    print(
        f"Preparing the verified runtime backup for {system}/{machine}/{python_tag}. "
        "The first launch downloads it from GitHub; later launches use the local cache.",
        file=sys.stderr,
        flush=True,
    )
    try:
        wheelhouse = ensure_wheelhouse(lock)
        command, selected = build_command(lock, uvx, wheelhouse, args)
    except (OSError, ValueError, KeyError, TypeError, RuntimeError, zipfile.BadZipFile) as exc:
        print(f"Excel MCP Server could not prepare its verified runtime: {exc}", file=sys.stderr, flush=True)
        return 2

    print(
        f"Starting {lock['distribution']['pypi_name']}=={lock['distribution']['version']} "
        f"with {selected} locked dependencies using offline installation.",
        file=sys.stderr,
        flush=True,
    )
    return subprocess.run(command, check=False).returncode


if __name__ == "__main__":
    raise SystemExit(main())
