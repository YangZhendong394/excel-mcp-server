# WorkBuddy Distribution Changelog

This file records changes specific to the maintained distribution. See the upstream [CHANGELOG.md](https://github.com/haris-musa/excel-mcp-server/blob/main/CHANGELOG.md) for the original engine's full history.

## 2.2.6.post1 — preparation baseline

- Established the independent distribution name `yangzhendong394-excel-mcp-server`.
- Retained the original `excel-mcp-server` console command for compatibility with the tested WorkBuddy adapter.
- Added maintainer and upstream provenance metadata.
- Kept the upstream MIT license and Haris Musa copyright notice.
- Updated the package URLs, server metadata, bundle manifest, and PyPI release destination.
- Disabled automatic MCP Registry publishing for this maintained package.
- Pinned the documented install examples to this package version.
- Added `runtime-lock.json` with exact runtime dependency versions, activation markers, artifact URLs, and SHA-256 hashes (34 packages; 552 artifacts).
- Added `scripts/generate_runtime_lock.py` to regenerate and verify the runtime manifest from `uv.lock`.
- Added `scripts/start_server.py` to invoke the pinned distribution with repeated stable `uvx --with` pins.
- No changes to the upstream Excel engine implementation have been made in this preparation baseline.

**Status:** preparation only; not yet published to PyPI. Runtime dependency pinning, package build verification, clean-install tests, and release credentials remain outstanding.
