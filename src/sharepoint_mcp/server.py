# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""MCP server skeleton.

This is the stub the v0.1 backlog grows into. Tools land here as their
respective issues close — auth + read tools first, then write tools
behind the SP_ALLOW_WRITES gate.
"""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

mcp: FastMCP = FastMCP("sharepoint-mcp")


def run() -> None:
    """Start the MCP server on stdio.

    Blocks until stdin closes.
    """
    mcp.run()
