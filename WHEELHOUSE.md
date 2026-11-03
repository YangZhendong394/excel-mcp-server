# WorkBuddy runtime wheelhouse

This repository builds a platform-specific backup archive for the maintained Excel MCP Server distribution.

## Current target

The first archived runtime target is:

- Windows x86-64
- CPython 3.13
- Exact package versions from `runtime-lock.json`

The archive contains one compatible prebuilt wheel per runtime dependency, the built wheel for this project, a file manifest with SHA-256 values, and the upstream MIT license notice.

## Reproducibility and verification

`scripts/build_wheelhouse.py --dry-run` lists the wheels selected for the current Python and platform without downloading or writing files.

The full build downloads the exact wheel URLs recorded in `runtime-lock.json`, verifies each SHA-256 and size, and refuses to build if any package has no compatible locked wheel.

GitHub Actions stores a 30-day workflow artifact for branch validation. When a GitHub Release is published, the workflow attaches the wheelhouse ZIP as a permanent release asset.

## Installation status

This archive is a backup source, not yet wired into `scripts/start_server.py`. The next step is to add a downloader that fetches the release archive, verifies its manifest, caches it locally, and launches uv with `--no-index --find-links --offline`.

The archive does not eliminate the need to download the archive itself from GitHub on a new machine. It does eliminate the need for that machine to fetch the package wheels from PyPI once the archive is available.
