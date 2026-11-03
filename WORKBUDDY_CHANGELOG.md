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
- Added `scripts/build_wheelhouse.py` to select platform-compatible wheels and verify each downloaded artifact against its locked SHA-256 and file size.
- Added a manually triggered Windows CPython 3.13 workflow that tests installation in offline/no-index mode and publishes the verified wheelhouse to a GitHub prerelease.
- Guarded the PyPI publisher so runtime-backup prereleases cannot accidentally trigger a PyPI publication.
- Adapted `src/excel_mcp/__init__.py` to recognize the maintained distribution metadata first, with a fallback to the original upstream distribution name. This preserves upstream use while allowing the renamed package to start.
- Added regression tests for maintained metadata, upstream fallback, and missing metadata.

**Status:** preparation only; not yet published to PyPI. Runtime dependency pinning, package build verification, clean-install tests, and release credentials remain outstanding.
