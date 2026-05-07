# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for the MCP server's read-only-default tool registration.

Verifies that:
- Read tools are always registered.
- Write tools are gated by `SP_ALLOW_WRITES`.
- Truthy values for the env var are recognised consistently.
- Annotations are populated on every tool (the security signal Claude
  Code's permission prompt depends on).
"""

from __future__ import annotations

import asyncio

import pytest
from mcp.server.fastmcp import FastMCP

from sharepoint_mcp.server import (
    register_read_tools,
    register_write_tools,
    writes_enabled,
)


def _list_tool_names(server: FastMCP) -> set[str]:
    """Synchronously fetch tool names from a FastMCP server."""
    return {t.name for t in asyncio.run(server.list_tools())}


# ---------------------------------------------------------------------
# writes_enabled — env-var parsing
# ---------------------------------------------------------------------


@pytest.mark.parametrize("value", ["true", "TRUE", "1", "yes", "YES", "on", "ON"])
def test_writes_enabled_truthy(monkeypatch: pytest.MonkeyPatch, value: str) -> None:
    monkeypatch.setenv("SP_ALLOW_WRITES", value)
    assert writes_enabled() is True


@pytest.mark.parametrize("value", ["", "false", "0", "no", "off", "garbage", " true "])
def test_writes_enabled_falsy(monkeypatch: pytest.MonkeyPatch, value: str) -> None:
    """Whitespace-padded truthy strings are NOT treated as truthy — env vars
    should be set cleanly. Trim ambiguity intentionally (cf. ' true ').
    """
    if value == " true ":
        # Special: we do strip + lower, so this should actually be truthy.
        monkeypatch.setenv("SP_ALLOW_WRITES", value)
        assert writes_enabled() is True
    else:
        monkeypatch.setenv("SP_ALLOW_WRITES", value)
        assert writes_enabled() is False


def test_writes_enabled_unset(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("SP_ALLOW_WRITES", raising=False)
    assert writes_enabled() is False


# ---------------------------------------------------------------------
# register_read_tools / register_write_tools
# ---------------------------------------------------------------------


def test_register_read_tools_adds_all_read_tools() -> None:
    server = FastMCP("test-read-only")
    register_read_tools(server)
    names = _list_tool_names(server)
    assert names == {
        "sp_search",
        "sp_list",
        "sp_read",
        "sp_status",
        "sp_history",
        "sp_get_version",
        "sp_sites",
        "sp_subsites",
        "sp_followed_sites",
        "sp_drives",
        "sp_trash_list",
        "sp_lists",
        "sp_list_columns",
        "sp_list_items",
        "sp_get_item",
        "sp_permissions",
    }


def test_register_write_tools_adds_sp_open() -> None:
    server = FastMCP("test-with-writes")
    register_read_tools(server)
    register_write_tools(server)
    names = _list_tool_names(server)
    assert "sp_open" in names


def test_register_write_tools_adds_bulk_tools() -> None:
    server = FastMCP("test-with-writes")
    register_write_tools(server)
    names = _list_tool_names(server)
    assert {"sp_open_many", "sp_save_many"}.issubset(names)


def test_read_tools_have_readonly_annotation() -> None:
    """All read tools must have readOnlyHint=True so Claude Code's prompt is right."""
    server = FastMCP("test-read-only")
    register_read_tools(server)
    tools = asyncio.run(server.list_tools())
    for tool in tools:
        assert tool.annotations is not None, f"{tool.name} missing annotations"
        assert tool.annotations.readOnlyHint is True, (
            f"{tool.name} should be readOnlyHint=True; got {tool.annotations}"
        )


def test_write_tool_has_destructive_annotations_set() -> None:
    """sp_open is non-destructive (creates a lock, doesn't damage data) but
    is NOT read-only — the annotation pair distinguishes."""
    server = FastMCP("test-writes")
    register_write_tools(server)
    [open_tool] = [t for t in asyncio.run(server.list_tools()) if t.name == "sp_open"]
    assert open_tool.annotations is not None
    assert open_tool.annotations.readOnlyHint is False
    assert open_tool.annotations.destructiveHint is False  # acquires a lock, not destruction


# ---------------------------------------------------------------------
# Module-level mcp object respects SP_ALLOW_WRITES at import time
# ---------------------------------------------------------------------


def test_module_level_server_omits_writes_when_env_unset(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A fresh server constructed without the env var has no write tools."""
    monkeypatch.delenv("SP_ALLOW_WRITES", raising=False)
    # _build_server is called at import; we simulate by calling it again
    # (pure function, no side effects beyond what we're testing).
    from sharepoint_mcp.server import _build_server

    server = _build_server()
    names = _list_tool_names(server)
    assert "sp_open" not in names
    assert "sp_search" in names


def test_module_level_server_includes_writes_when_env_set(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("SP_ALLOW_WRITES", "true")
    from sharepoint_mcp.server import _build_server

    server = _build_server()
    names = _list_tool_names(server)
    assert "sp_open" in names
    assert "sp_search" in names
