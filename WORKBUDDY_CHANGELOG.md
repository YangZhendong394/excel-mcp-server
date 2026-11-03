# WorkBuddy Distribution Changelog

This file records changes specific to the maintained distribution. See the upstream [CHANGELOG.md](https://github.com/haris-musa/excel-mcp-server/blob/main/CHANGELOG.md) for the original engine's full history.

## 1.0.0 — WorkBuddy Skill package

- Added the regular WorkBuddy Skill folder `workbuddy-skill/excel-workbook-assistant/` with official YAML frontmatter, Chinese/English metadata, usage examples, local-file safety guidance, and upstream notices.
- Added `scripts/excel_mcp_bridge.py`, a small Bash-invoked stdio MCP bridge for listing and calling the Excel tools.
- Added a package builder that verifies the required file list and enforces the 3 MB ZIP limit.
- Published `excel-workbook-assistant-1.0.0.zip` and its SHA-256 file as GitHub Release assets.
- Added an end-to-end CI smoke test that loads the packaged Skill, lists 46 tools, creates a workbook, writes sample cells, and reads the values back.

**Status:** the GitHub-hosted Windows x64 / CPython 3.13 Skill package and runtime backup have passed automated end-to-end tests. PyPI publication and a final installation test inside the user's own WorkBuddy desktop remain outstanding.

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
- Updated `scripts/start_server.py` to download a platform-specific GitHub Release wheelhouse, verify the release checksum and every dependency against `runtime-lock.json`, cache it locally, and launch with uv in offline/no-index mode.
- Added `scripts/build_wheelhouse.py` to select platform-compatible wheels and verify each downloaded artifact against its locked SHA-256 and file size.
- Added an automated Windows x64 / CPython 3.13 workflow that builds and publishes the hash-verified wheelhouse to a GitHub prerelease, then tests the published asset through the launcher.
- Guarded the PyPI publisher so runtime-backup prereleases cannot accidentally trigger a PyPI publication.
- Adapted `src/excel_mcp/__init__.py` to recognize the maintained distribution metadata first, with a fallback to the original upstream distribution name. This preserves upstream use while allowing the renamed package to start.
- Added regression tests for maintained metadata, upstream fallback, and missing metadata.
- Extended `scripts/test_offline_wheelhouse.py` to test either a local wheelhouse or the published GitHub Release, then perform a real MCP initialize and tools/list handshake.

**Status:** not yet published to PyPI. The Windows x64 / CPython 3.13 runtime backup is published as a GitHub prerelease. The release workflow now tests downloading that asset through the WorkBuddy launcher and starting the MCP server offline. Other platform/Python backups and validation inside the user's own WorkBuddy installation remain outstanding.
