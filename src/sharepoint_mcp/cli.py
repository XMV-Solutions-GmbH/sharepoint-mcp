# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Command-line entry point.

Three commands:

- `mcp-server-sharepoint login [--profile NAME]` — interactive Device Code
  flow, persists the resulting tokens to the configured TokenStore.
  Run once per profile; subsequent MCP-server starts use the cached
  refresh token silently.
- `mcp-server-sharepoint logout [--profile NAME]` — clears cached credentials.
- `mcp-server-sharepoint` (no command) — starts the MCP server on stdio.
  This is what `.mcp.json` invokes.

The login/logout subcommands deliberately mirror the `gh auth
login` / `gh auth logout` pattern: auth setup is out-of-band, the
running MCP process never blocks for human interaction.
"""

from __future__ import annotations

import argparse
from collections.abc import Sequence

from sharepoint_mcp import __version__


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="mcp-server-sharepoint",
        description=(
            "MCP server for SharePoint document libraries with audit-preserving checkout/checkin."
        ),
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )

    subparsers = parser.add_subparsers(dest="command")

    login_p = subparsers.add_parser(
        "login",
        help="Sign in via Microsoft Device Code flow and cache the refresh token.",
    )
    login_p.add_argument(
        "--profile",
        default="default",
        help="Profile name (namespace for token cache + working dir). Default: 'default'.",
    )

    logout_p = subparsers.add_parser(
        "logout",
        help="Remove cached credentials for a profile.",
    )
    logout_p.add_argument(
        "--profile",
        default="default",
        help="Profile name. Default: 'default'.",
    )

    return parser


def main(argv: Sequence[str] | None = None) -> int:
    """Parse CLI arguments and dispatch to the right subcommand.

    Returns the process exit code.
    """
    parser = _build_parser()
    args = parser.parse_args(argv)

    if args.command == "login":
        from sharepoint_mcp.auth import interactive_login

        interactive_login(profile=args.profile)
        return 0

    if args.command == "logout":
        from sharepoint_mcp.auth.store import get_token_store

        get_token_store().delete(args.profile)
        return 0

    # No subcommand — start the MCP server on stdio.
    from sharepoint_mcp.server import run

    run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
