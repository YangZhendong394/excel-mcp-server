#!/usr/bin/env python3
"""Generate or verify the exact runtime dependency manifest from uv.lock.

This file deliberately uses only the Python standard library so it can run in CI
without adding another dependency.
"""
from __future__ import annotations

import argparse
import json
import sys
import tomllib
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
LOCK_PATH = ROOT / "uv.lock"
PROJECT_PATH = ROOT / "pyproject.toml"
SOURCE_PATH = ROOT / "SOURCE.json"
OUTPUT_PATH = ROOT / "runtime-lock.json"


def build_manifest() -> dict:
    project = tomllib.loads(PROJECT_PATH.read_text(encoding="utf-8"))
    lock = tomllib.loads(LOCK_PATH.read_text(encoding="utf-8"))
    source = json.loads(SOURCE_PATH.read_text(encoding="utf-8"))

    project_meta = project["project"]
    root_name = project_meta["name"]
    packages = {item["name"]: item for item in lock["package"]}

    if root_name not in packages:
        raise ValueError(f"{root_name!r} is missing from uv.lock")

    activation_paths: dict[str, set[tuple[str, ...]]] = {}

    def visit(
        name: str,
        extras: tuple[str, ...] = (),
        conditions: tuple[str, ...] = (),
        ancestors: tuple[str, ...] = (),
    ) -> None:
        if name in ancestors:
            return
        package = packages.get(name)
        if package is None:
            raise ValueError(f"uv.lock is missing dependency {name!r}")

        activation_paths.setdefault(name, set()).add(tuple(sorted(set(conditions))))
        next_ancestors = (*ancestors, name)

        for dependency in package.get("dependencies", []):
            marker = dependency.get("marker")
            next_conditions = conditions + ((marker,) if marker else ())
            visit(
                dependency["name"],
                tuple(dependency.get("extra", [])),
                next_conditions,
                next_ancestors,
            )

        optional = package.get("optional-dependencies", {})
        for extra in extras:
            for dependency in optional.get(extra, []):
                marker = dependency.get("marker")
                next_conditions = conditions + ((marker,) if marker else ())
                visit(
                    dependency["name"],
                    (),
                    next_conditions,
                    next_ancestors,
                )

    visit(root_name)

    runtime_packages = []
    for name in sorted(activation_paths):
        if name == root_name:
            continue

        package = packages[name]
        artifacts = []
        sdist = package.get("sdist")
        if sdist:
            artifacts.append(
                {
                    "kind": "sdist",
                    "url": sdist["url"],
                    "sha256": sdist["hash"].removeprefix("sha256:"),
                    "size": sdist.get("size"),
                }
            )

        for wheel in package.get("wheels", []):
            artifacts.append(
                {
                    "kind": "wheel",
                    "url": wheel["url"],
                    "sha256": wheel["hash"].removeprefix("sha256:"),
                    "size": wheel.get("size"),
                }
            )

        if not artifacts:
            raise ValueError(f"No locked artifacts recorded for {name}=={package['version']}")

        runtime_packages.append(
            {
                "name": name,
                "version": package["version"],
                "source": package.get("source", {}).get("registry", "editable"),
                "activation_paths": [
                    list(path)
                    for path in sorted(
                        activation_paths[name],
                        key=lambda item: json.dumps(item, separators=(",", ":")),
                    )
                ],
                "artifacts": sorted(artifacts, key=lambda item: item["url"]),
            }
        )

    command_name = next(iter(project_meta["scripts"]))
    project_version = project_meta["version"]
    artifact_count = sum(len(item["artifacts"]) for item in runtime_packages)

    return {
        "schema_version": 1,
        "generated_from": {
            "lockfile": "uv.lock",
            "project": "pyproject.toml",
            "project_version": project_version,
            "purpose": (
                "Exact runtime dependency pins and artifact SHA-256 records "
                "for WorkBuddy's isolated engine launcher."
            ),
        },
        "distribution": {
            "pypi_name": root_name,
            "version": project_version,
            "console_command": command_name,
            "python_requires": project_meta["requires-python"],
        },
        "install_strategy": {
            "uv_options": (
                "Repeated --with name==version options, followed by "
                "--from distribution==version and the console command."
            ),
            "reason": (
                "Use stable uvx package-pin options instead of relying on "
                "the preview --with-requirements interface."
            ),
            "supported_platforms": ["win32", "darwin", "linux"],
        },
        "dependency_count": len(runtime_packages),
        "artifact_count": artifact_count,
        "packages": runtime_packages,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--check", action="store_true", help="fail if runtime-lock.json is stale")
    group.add_argument("--write", action="store_true", help="regenerate runtime-lock.json")
    args = parser.parse_args()

    expected = build_manifest()
    expected_text = json.dumps(expected, indent=2, ensure_ascii=False) + "\n"

    if args.write:
        OUTPUT_PATH.write_text(expected_text, encoding="utf-8")
        print(f"Updated {OUTPUT_PATH.relative_to(ROOT)}")
        return 0

    if not OUTPUT_PATH.exists():
        print("runtime-lock.json is missing; run with --write", file=sys.stderr)
        return 1

    actual = json.loads(OUTPUT_PATH.read_text(encoding="utf-8"))
    if actual != expected:
        print(
            "runtime-lock.json does not match pyproject.toml / uv.lock. "
            "Run: python scripts/generate_runtime_lock.py --write",
            file=sys.stderr,
        )
        return 1

    print(
        f"Runtime lock is current: {expected['dependency_count']} dependencies, "
        f"{expected['artifact_count']} artifacts."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
