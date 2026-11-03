#!/usr/bin/env python3
"""Small stdio MCP client bridge for the WorkBuddy Excel Skill.

Usage:
  python scripts/excel_mcp_bridge.py --allow-dir "C:\\Workbooks" list-tools
  python scripts/excel_mcp_bridge.py --allow-dir "C:\\Workbooks" call-tool read_range --arguments-json '{"workbook_path":"book.xlsx","range":"A1:C5"}'

The bridge starts the local, hash-verified runtime launcher, performs an MCP
handshake, invokes one tool, prints a JSON result, and shuts the server down.
"""
from __future__ import annotations

import argparse
import json
import queue
import subprocess
import sys
import threading
import time
from collections import deque
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
LAUNCHER = ROOT / "scripts" / "start_server.py"


class BridgeError(RuntimeError):
    pass


def read_response(
    messages: queue.Queue[str],
    process: subprocess.Popen[str],
    message_id: int,
    timeout: int,
) -> dict[str, Any]:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            line = messages.get(timeout=0.5)
        except queue.Empty:
            if process.poll() is not None:
                raise BridgeError(
                    f"Excel MCP server stopped before replying (exit code {process.returncode})."
                )
            continue
        try:
            payload = json.loads(line)
        except json.JSONDecodeError:
            # Keep stdout protocol-safe; stray diagnostic lines are ignored.
            continue
        if payload.get("id") == message_id:
            if "error" in payload:
                raise BridgeError(f"MCP error: {payload['error']}")
            return payload
    raise BridgeError(f"Timed out waiting for the Excel MCP server (request {message_id}).")


def load_arguments(args: argparse.Namespace) -> dict[str, Any]:
    if args.arguments_file:
        value = json.loads(args.arguments_file.read_text(encoding="utf-8"))
    elif args.arguments_json:
        value = json.loads(args.arguments_json)
    else:
        value = {}
    if not isinstance(value, dict):
        raise BridgeError("Tool arguments must be a JSON object.")
    return value


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--allow-dir", required=True, type=Path, help="Folder containing the Excel files this task may access.")
    parser.add_argument("--timeout", type=int, default=180, help="Seconds to wait for the server and tool response.")
    subparsers = parser.add_subparsers(dest="action", required=True)
    subparsers.add_parser("list-tools", help="Show the available Excel MCP tools and their input schemas.")
    call_parser = subparsers.add_parser("call-tool", help="Call one Excel MCP tool.")
    call_parser.add_argument("tool_name")
    arg_group = call_parser.add_mutually_exclusive_group()
    arg_group.add_argument("--arguments-json")
    arg_group.add_argument("--arguments-file", type=Path)
    args = parser.parse_args()

    if not LAUNCHER.is_file():
        raise BridgeError(f"Runtime launcher is missing: {LAUNCHER}")
    if not args.allow_dir.exists() or not args.allow_dir.is_dir():
        raise BridgeError(f"The allowed workbook folder does not exist: {args.allow_dir}")
    allow_dir = args.allow_dir.resolve()
    if args.timeout < 10:
        raise BridgeError("--timeout must be at least 10 seconds.")

    command = [
        sys.executable,
        str(LAUNCHER),
        "stdio",
        "--allow-dir",
        str(allow_dir),
    ]
    process = subprocess.Popen(
        command,
        cwd=ROOT,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )

    messages: queue.Queue[str] = queue.Queue()
    stderr_lines: deque[str] = deque(maxlen=120)

    def read_stdout() -> None:
        assert process.stdout is not None
        for line in process.stdout:
            messages.put(line)

    def read_stderr() -> None:
        assert process.stderr is not None
        for line in process.stderr:
            stderr_lines.append(line.rstrip())

    threading.Thread(target=read_stdout, daemon=True).start()
    threading.Thread(target=read_stderr, daemon=True).start()

    def send(payload: dict[str, Any]) -> None:
        assert process.stdin is not None
        process.stdin.write(json.dumps(payload, ensure_ascii=False) + "\n")
        process.stdin.flush()

    try:
        send(
            {
                "jsonrpc": "2.0",
                "id": 1,
                "method": "initialize",
                "params": {
                    "protocolVersion": "2025-03-26",
                    "capabilities": {},
                    "clientInfo": {"name": "workbuddy-excel-skill", "version": "1.0.0"},
                },
            }
        )
        initialized = read_response(messages, process, 1, args.timeout)
        server_info = initialized.get("result", {}).get("serverInfo", {})
        send({"jsonrpc": "2.0", "method": "notifications/initialized"})

        if args.action == "list-tools":
            send(
                {
                    "jsonrpc": "2.0",
                    "id": 2,
                    "method": "tools/list",
                    "params": {},
                }
            )
            response = read_response(messages, process, 2, args.timeout)
            result = response.get("result", {})
            output = {
                "server": server_info,
                "tool_count": len(result.get("tools", [])),
                "tools": result.get("tools", []),
            }
        else:
            arguments = load_arguments(args)
            send(
                {
                    "jsonrpc": "2.0",
                    "id": 2,
                    "method": "tools/call",
                    "params": {"name": args.tool_name, "arguments": arguments},
                }
            )
            response = read_response(messages, process, 2, args.timeout)
            result = response.get("result", {})
            output = {"tool": args.tool_name, "result": result}
            if result.get("isError"):
                print(json.dumps(output, ensure_ascii=False, indent=2), flush=True)
                return 1

        print(json.dumps(output, ensure_ascii=False, indent=2), flush=True)
        return 0

    except (BridgeError, OSError, ValueError, json.JSONDecodeError) as exc:
        print(f"Excel Skill error: {exc}", file=sys.stderr, flush=True)
        if stderr_lines:
            print("\n".join(stderr_lines)[-6000:], file=sys.stderr, flush=True)
        return 1

    finally:
        try:
            process.terminate()
            process.wait(timeout=8)
        except Exception:
            try:
                process.kill()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
