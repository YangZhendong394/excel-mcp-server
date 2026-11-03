"""Tests for version metadata across upstream and maintained distributions."""

from importlib.metadata import PackageNotFoundError

import excel_mcp


def test_version_prefers_maintained_distribution(monkeypatch):
    calls = []

    def fake_version(distribution_name: str) -> str:
        calls.append(distribution_name)
        if distribution_name == "yangzhendong394-excel-mcp-server":
            return "2.2.6.post1"
        raise PackageNotFoundError(distribution_name)

    monkeypatch.setattr(excel_mcp, "version", fake_version)

    assert excel_mcp._get_version() == "2.2.6.post1"
    assert calls == ["yangzhendong394-excel-mcp-server"]


def test_version_falls_back_to_upstream_distribution(monkeypatch):
    calls = []

    def fake_version(distribution_name: str) -> str:
        calls.append(distribution_name)
        if distribution_name == "excel-mcp-server":
            return "2.2.6"
        raise PackageNotFoundError(distribution_name)

    monkeypatch.setattr(excel_mcp, "version", fake_version)

    assert excel_mcp._get_version() == "2.2.6"
    assert calls == [
        "yangzhendong394-excel-mcp-server",
        "excel-mcp-server",
    ]


def test_version_uses_unknown_when_no_distribution_metadata_exists(monkeypatch):
    def missing_version(distribution_name: str) -> str:
        raise PackageNotFoundError(distribution_name)

    monkeypatch.setattr(excel_mcp, "version", missing_version)

    assert excel_mcp._get_version() == "0+unknown"
