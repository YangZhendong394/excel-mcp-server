#!/usr/bin/env python3
"""Launch the pinned Excel MCP Server distribution using the runtime lock.

The WorkBuddy skill can invoke this script as a local stdio MCP server process.
All server protocol traffic is inherited unchanged; diagnostics go to stderr.
"""
from __future__ import annotations

import json
import platform
import re
import shutil
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
LOCK_PATH = ROOT / "runtime-lock.json"


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


def build_command(lock: dict, uvx: str, server_args: list[str]) -> tuple[list[str], list[dict]]:
    environment = {
        "sys_platform": sys.platform,
        "python_version": f"{sys.version_info.major}.{sys.version_info.minor}",
        "python_full_version": platform.python_version(),
        "platform_python_implementation": platform.python_implementation(),
        "implementation_name": sys.implementation.name,
    }

    command = [uvx]
    selected = []
    for package in lock["packages"]:
        if package_is_applicable(package, environment):
            selected.append(package)
            command.extend(["--with", f"{package['name']}=={package['version']}"])

    distribution = lock["distribution"]
    command.extend(
        [
            "--from",
            f"{distribution['pypi_name']}=={distribution['version']}",
            distribution["console_command"],
        ]
    )
    command.extend(server_args or ["stdio"])
    return command, selected


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

    command, selected = build_command(lock, uvx, args)
    if show_command:
        print(json.dumps({"selected_dependency_count": len(selected), "command": command}, indent=2))
        return 0

    print(
        f"Starting {lock['distribution']['pypi_name']}=={lock['distribution']['version']} "
        f"with {len(selected)} locked runtime dependencies. "
        "The first launch may download packages from the configured Python package index.",
        file=sys.stderr,
        flush=True,
    )
    return subprocess.run(command, check=False).returncode


if __name__ == "__main__":
    raise SystemExit(main())
