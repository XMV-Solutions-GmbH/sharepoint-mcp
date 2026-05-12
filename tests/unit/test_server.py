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

from sharepoint_mcp.auth.flow import SharepointConsentNotConfiguredError
from sharepoint_mcp.server import (
    register_read_tools,
    register_write_tools,
    writes_enabled,
)


def _list_tool_names(server: FastMCP) -> set[str]:
    """Synchronously fetch tool names from a FastMCP server."""
    return {t.name for t in asyncio.run(server.list_tools())}


# ---------------------------------------------------------------------
# writes_enabled — strict env-var parsing (v0.5)
# ---------------------------------------------------------------------


@pytest.mark.parametrize("value", ["true", "TRUE", " true ", "True"])
def test_writes_enabled_true_accepts_case_and_whitespace(
    monkeypatch: pytest.MonkeyPatch, value: str
) -> None:
    monkeypatch.setenv("SP_ALLOW_WRITES", value)
    assert writes_enabled() is True


@pytest.mark.parametrize("value", ["false", "FALSE", " false "])
def test_writes_enabled_false_accepts_case_and_whitespace(
    monkeypatch: pytest.MonkeyPatch, value: str
) -> None:
    monkeypatch.setenv("SP_ALLOW_WRITES", value)
    assert writes_enabled() is False


@pytest.mark.parametrize("value", ["1", "yes", "on", "garbage", "", "0", "no", "off"])
def test_writes_enabled_strict_rejects_legacy_and_other_values(
    monkeypatch: pytest.MonkeyPatch, value: str
) -> None:
    """v0.5 breaking change: only exactly 'true' / 'false' accepted.
    Legacy v0.4 truthy values (1/yes/on) and any other string raise."""
    monkeypatch.setenv("SP_ALLOW_WRITES", value)
    with pytest.raises(SharepointConsentNotConfiguredError, match="SP_ALLOW_WRITES"):
        writes_enabled()


def test_writes_enabled_unset_raises(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("SP_ALLOW_WRITES", raising=False)
    with pytest.raises(SharepointConsentNotConfiguredError, match="not set"):
        writes_enabled()


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
        "sp_share_list",
        "sp_pages_list",
        "sp_page_read",
        "sp_changes",
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


def test_module_level_server_omits_writes_in_explicit_readonly_mode(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Explicit SP_ALLOW_WRITES=false → no write tools, read tools still there."""
    monkeypatch.setenv("SP_ALLOW_WRITES", "false")
    from sharepoint_mcp.server import _build_server

    server = _build_server()
    names = _list_tool_names(server)
    assert "sp_open" not in names
    assert "sp_search" in names


def test_module_level_server_refuses_when_env_unset(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """_build_server raises SharepointConsentNotConfiguredError when SP_ALLOW_WRITES unset."""
    monkeypatch.delenv("SP_ALLOW_WRITES", raising=False)
    from sharepoint_mcp.server import _build_server

    with pytest.raises(SharepointConsentNotConfiguredError, match="not set"):
        _build_server()


def test_module_level_server_refuses_on_legacy_truthy(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """v0.4 legacy `SP_ALLOW_WRITES=yes` is rejected — must be explicit true/false."""
    monkeypatch.setenv("SP_ALLOW_WRITES", "yes")
    from sharepoint_mcp.server import _build_server

    with pytest.raises(SharepointConsentNotConfiguredError, match="SP_ALLOW_WRITES"):
        _build_server()


def test_module_level_server_includes_writes_when_env_set(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("SP_ALLOW_WRITES", "true")
    from sharepoint_mcp.server import _build_server

    server = _build_server()
    names = _list_tool_names(server)
    assert "sp_open" in names
    assert "sp_search" in names
