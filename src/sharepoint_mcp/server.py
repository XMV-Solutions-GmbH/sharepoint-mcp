# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""MCP server: registers the `sp_*` tools with FastMCP and runs on stdio.

Each tool is wrapped with explicit `ToolAnnotations` so MCP clients
(notably Claude Code's permission system) can render the right
prompt — read-only tools get a different treatment from destructive
ones. The annotations are part of our security story: if we lie here,
the client can't make sensible safety decisions.
"""

from __future__ import annotations

import os
from typing import Any

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from sharepoint_mcp.tools.list_folder import list_folder as _do_list
from sharepoint_mcp.tools.read import read_file as _do_read
from sharepoint_mcp.tools.search import search as _do_search

PROFILE_ENV = "SP_PROFILE"
DEFAULT_PROFILE = "default"

mcp: FastMCP = FastMCP("sharepoint-mcp")


def _get_profile() -> str:
    """Profile name for this MCP-server-process; from `SP_PROFILE` env var."""
    return os.environ.get(PROFILE_ENV, DEFAULT_PROFILE)


@mcp.tool(
    annotations=ToolAnnotations(
        title="Search SharePoint",
        readOnlyHint=True,
        idempotentHint=True,
        openWorldHint=False,
    ),
    description=(
        "Search the SharePoint document libraries the signed-in user has access to. "
        "Returns matching files with name, path, webUrl, last-modified date, and author. "
        "Read-only — does not modify any SharePoint state. "
        "Filter args: site (URL), folder (path), file_type (extension like 'docx'), "
        "modified_after (ISO date)."
    ),
)
def sp_search(
    query: str,
    site: str | None = None,
    folder: str | None = None,
    file_type: str | None = None,
    modified_after: str | None = None,
    limit: int = 25,
) -> list[dict[str, Any]]:
    return _do_search(
        query,
        site=site,
        folder=folder,
        file_type=file_type,
        modified_after=modified_after,
        limit=limit,
        profile=_get_profile(),
    )


@mcp.tool(
    annotations=ToolAnnotations(
        title="List SharePoint Folder",
        readOnlyHint=True,
        idempotentHint=True,
        openWorldHint=False,
    ),
    description=(
        "List the immediate children of a SharePoint or OneDrive folder. "
        "`url` is the folder's human-readable web URL (e.g. from a previous "
        "sp_search hit's web_url, or the SharePoint web UI). Returns each "
        "child with name, type ('folder' or 'file'), size, last-modified date, "
        "and webUrl. Read-only — does not modify SharePoint state."
    ),
)
def sp_list(url: str, limit: int = 100) -> list[dict[str, Any]]:
    return _do_list(url, limit=limit, profile=_get_profile())


@mcp.tool(
    annotations=ToolAnnotations(
        title="Read SharePoint File",
        readOnlyHint=True,
        idempotentHint=True,
        openWorldHint=False,
    ),
    description=(
        "Download a SharePoint file's content to a local temp file. Returns the "
        "absolute path of the temp file with the original extension preserved. "
        "Read-only — does NOT acquire a checkout/lock; use sp_open for that. "
        "`url` is the file's human-readable web URL (e.g. from sp_search hits)."
    ),
)
def sp_read(url: str) -> str:
    return _do_read(url, profile=_get_profile())


def run() -> None:
    """Start the MCP server on stdio.

    Blocks until stdin closes.
    """
    mcp.run()
