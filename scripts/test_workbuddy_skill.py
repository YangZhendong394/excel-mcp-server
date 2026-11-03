#!/usr/bin/env python3
"""End-to-end test of the packaged WorkBuddy Skill and local Excel bridge."""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SKILL = ROOT / "workbuddy-skill" / "excel-workbook-assistant"


def run_bridge(
    bridge: Path,
    allowed_dir: Path,
    action: str,
    tool_name: str | None = None,
    arguments: dict[str, Any] | None = None,
) -> dict[str, Any]:
    command = [
        sys.executable,
        str(bridge),
        "--allow-dir",
        str(allowed_dir),
        "--timeout",
        "180",
        action,
    ]
    if tool_name is not None:
        command.append(tool_name)
        command.extend(["--arguments-json", json.dumps(arguments or {}, ensure_ascii=False)])

    env = os.environ.copy()
    env["UV_OFFLINE"] = "1"
    cache_root = allowed_dir / "runtime-cache"
    if os.name == "nt":
        env["LOCALAPPDATA"] = str(cache_root)
    else:
        env["XDG_CACHE_HOME"] = str(cache_root)

    completed = subprocess.run(
        command,
        cwd=bridge.parents[1],
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=240,
        check=False,
    )
    if completed.returncode:
        raise RuntimeError(
            f"Bridge command failed ({completed.returncode}): "
            f"{' '.join(command)}\nSTDOUT:\n{completed.stdout}\nSTDERR:\n{completed.stderr}"
        )
    try:
        value = json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            f"Bridge did not return JSON. STDOUT:\n{completed.stdout}\nSTDERR:\n{completed.stderr}"
        ) from exc
    if not isinstance(value, dict):
        raise RuntimeError("Bridge returned an unexpected response.")
    return value


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--skill-dir", type=Path, default=DEFAULT_SKILL)
    args = parser.parse_args()

    skill_root = args.skill_dir.resolve()
    bridge = skill_root / "scripts" / "excel_mcp_bridge.py"
    required = [
        bridge,
        skill_root / "scripts" / "start_server.py",
        skill_root / "runtime-lock.json",
        skill_root / "SKILL.md",
    ]
    for path in required:
        if not path.is_file():
            raise FileNotFoundError(path)

    with tempfile.TemporaryDirectory(prefix="workbuddy-excel-skill-test-") as temp:
        allowed_dir = Path(temp)
        print("=== WORKBUDDY SKILL END-TO-END TEST ===", flush=True)
        print(f"skill_dir={skill_root}", flush=True)

        listing = run_bridge(bridge, allowed_dir, "list-tools")
        tools = listing.get("tools", [])
        names = {tool.get("name") for tool in tools}
        if len(tools) < 40:
            raise RuntimeError(f"Expected at least 40 tools; received {len(tools)}")
        print(f"skill_tools_list=OK ({len(tools)} tools)", flush=True)

        required_tools = {"create_workbook", "write_range", "read_range"}
        if not required_tools.issubset(names):
            missing = sorted(required_tools - names)
            raise RuntimeError(f"The skill is missing required workbook tools: {missing}")

        workbook = "workbuddy-skill-smoke.xlsx"
        created = run_bridge(
            bridge,
            allowed_dir,
            "call-tool",
            "create_workbook",
            {"path": workbook, "sheets": ["Smoke"]},
        )
        if created.get("result", {}).get("isError"):
            raise RuntimeError(f"create_workbook returned an MCP error: {created}")
        file_path = allowed_dir / workbook
        if not file_path.is_file():
            raise RuntimeError("create_workbook did not create the expected local file.")
        print("skill_create_workbook=OK", flush=True)

        written = run_bridge(
            bridge,
            allowed_dir,
            "call-tool",
            "write_range",
            {
                "path": workbook,
                "sheet": "Smoke",
                "at": "A1",
                "rows": [["item", "amount"], ["alpha", 123]],
            },
        )
        if written.get("result", {}).get("isError"):
            raise RuntimeError(f"write_range returned an MCP error: {written}")
        print("skill_write_range=OK", flush=True)

        readback = run_bridge(
            bridge,
            allowed_dir,
            "call-tool",
            "read_range",
            {"path": workbook, "sheet": "Smoke", "range": "A1:B2"},
        )
        serialized = json.dumps(readback, ensure_ascii=False)
        if "alpha" not in serialized or "123" not in serialized:
            raise RuntimeError(f"Read-back did not contain the written values: {readback}")
        print("skill_read_range=OK", flush=True)
        print("RESULT=SUCCESS", flush=True)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
