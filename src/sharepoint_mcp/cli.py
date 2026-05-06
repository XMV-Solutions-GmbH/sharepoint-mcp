# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Command-line entry point for the sharepoint-mcp server.

Parses arguments, then hands off to the MCP server which runs on stdio.
The CLI surface is intentionally minimal — the protocol is the API.
"""

from __future__ import annotations

import argparse
from collections.abc import Sequence

from sharepoint_mcp import __version__


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="sharepoint-mcp",
        description=(
            "MCP server for SharePoint document libraries with audit-preserving checkout/checkin."
        ),
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    """Parse CLI arguments and start the MCP server on stdio.

    Returns the process exit code.
    """
    parser = _build_parser()
    parser.parse_args(argv)

    from sharepoint_mcp.server import run

    run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
