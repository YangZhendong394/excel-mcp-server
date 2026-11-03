"""MCP server for Excel workbooks."""

from importlib.metadata import PackageNotFoundError, version

_DISTRIBUTION_NAMES = (
    "yangzhendong394-excel-mcp-server",
    "excel-mcp-server",
)


def _get_version() -> str:
    """Read the installed version from the maintained or upstream distribution."""
    for distribution_name in _DISTRIBUTION_NAMES:
        try:
            return version(distribution_name)
        except PackageNotFoundError:
            continue
    return "0+unknown"


__version__ = _get_version()
