#!/usr/bin/env python3
"""Install the maintained Excel MCP server from a local wheelhouse and test MCP.

The test forces uv offline and disables registry indexes. It uses a fresh uv
cache so a successful run proves the wheelhouse is self-contained.
"""

from __future__ import annotations

import argparse
import json
import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from pathlib import Path
from typing import Any

from start_server import build_command, ensure_wheelhouse


ROOT = Path(__file__).resolve().parents[1]
LOCK_PATH = ROOT / "runtime-lock.json"


def read_response(
    process: subprocess.Popen[str],
    messages: queue.Queue[str],
    message_id: int,
    timeout: int,
) -> dict[str, Any]:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            line = messages.get(timeout=0.5)
        except queue.Empty:
            if process.poll() is not None:
                raise RuntimeError(
                    f"Server exited before response {message_id}; exit code={process.returncode}"
                )
            continue

        try:
            response = json.loads(line)
        except json.JSONDecodeError:
            print(f"non_json_stdout={line.rstrip()}", flush=True)
            continue

        if response.get("id") == message_id:
            return response

    raise TimeoutError(f"Timed out waiting for JSON-RPC response {message_id}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--wheelhouse", type=Path)
    source.add_argument(
        "--github-release",
        action="store_true",
        help="download and verify the published GitHub Release wheelhouse",
    )
    args = parser.parse_args()

    if not LOCK_PATH.is_file():
        raise FileNotFoundError(LOCK_PATH)

    lock = json.loads(LOCK_PATH.read_text(encoding="utf-8"))
    distribution = lock["distribution"]
    uvx = shutil.which("uvx")
    if not uvx:
        raise RuntimeError("uvx was not found in PATH")

    # Windows may briefly keep imported .pyd files locked after the child exits.
    # Ignore cleanup-only errors; the hosted runner removes its temp directory at job end.
    with tempfile.TemporaryDirectory(
        prefix="excel-mcp-offline-uv-cache-", ignore_cleanup_errors=True
    ) as cache:
        env = os.environ.copy()
        env["UV_CACHE_DIR"] = str(Path(cache) / "uv-cache")
        env["UV_OFFLINE"] = "1"

        if args.github_release:
            cache_variable = "LOCALAPPDATA" if os.name == "nt" else "XDG_CACHE_HOME"
            previous_cache = os.environ.get(cache_variable)
            os.environ[cache_variable] = str(Path(cache) / "launcher-cache")
            try:
                wheelhouse = ensure_wheelhouse(lock)
            finally:
                if previous_cache is None:
                    os.environ.pop(cache_variable, None)
                else:
                    os.environ[cache_variable] = previous_cache
            print("github_release_download_and_hash_verification=OK", flush=True)
        else:
            if args.wheelhouse is None or not args.wheelhouse.is_dir():
                raise FileNotFoundError(args.wheelhouse)
            wheelhouse = args.wheelhouse.resolve()

        command, dependency_count = build_command(lock, uvx, wheelhouse, ["stdio"])

        print("=== OFFLINE WHEELHOUSE MCP TEST ===", flush=True)
        print(f"wheelhouse={wheelhouse}", flush=True)
        print(f"locked_dependency_count={dependency_count}", flush=True)
        print("registry_access=DISABLED", flush=True)
        print("uv_cache=FRESH", flush=True)

        process = subprocess.Popen(
            command,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
            env=env,
        )
        messages: queue.Queue[str] = queue.Queue()
        stderr_lines: list[str] = []

        def read_stdout() -> None:
            assert process.stdout is not None
            for line in process.stdout:
                messages.put(line)

        def read_stderr() -> None:
            assert process.stderr is not None
            for line in process.stderr:
                stderr_lines.append(line)

        threading.Thread(target=read_stdout, daemon=True).start()
        threading.Thread(target=read_stderr, daemon=True).start()

        def send(payload: dict[str, Any]) -> None:
            assert process.stdin is not None
            process.stdin.write(json.dumps(payload) + "\n")
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
                        "clientInfo": {
                            "name": "workbuddy-offline-wheelhouse-test",
                            "version": "1.0.0",
                        },
                    },
                }
            )
            initialized = read_response(process, messages, 1, 120)
            if "error" in initialized:
                raise RuntimeError(f"initialize error: {initialized['error']}")
            print("offline_install=OK", flush=True)
            print("mcp_initialize=OK", flush=True)

            send(
                {
                    "jsonrpc": "2.0",
                    "method": "notifications/initialized",
                }
            )
            send(
                {
                    "jsonrpc": "2.0",
                    "id": 2,
                    "method": "tools/list",
                    "params": {},
                }
            )
            result = read_response(process, messages, 2, 45)
            if "error" in result:
                raise RuntimeError(f"tools/list error: {result['error']}")

            tools = result.get("result", {}).get("tools", [])
            print(f"tools_list=OK", flush=True)
            print(f"tool_count={len(tools)}", flush=True)
            if len(tools) < 40:
                raise RuntimeError(f"Expected at least 40 Excel tools; received {len(tools)}")
            print("RESULT=SUCCESS", flush=True)
            return 0

        except Exception as exc:
            print(f"RESULT=FAILED: {type(exc).__name__}: {exc}", flush=True)
            if stderr_lines:
                print("=== SERVER STDERR ===", flush=True)
                print("".join(stderr_lines)[-6000:], flush=True)
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
