# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for the MCP server's tool registration (v0.7.0).

Verifies that:
- Auth tools are always registered.
- Per-category register functions add exactly the expected `sp_<cat>_*` tools.
- Write tools are gated by `SP_ALLOW_WRITES`.
- `SP_TOOL_GROUPS` filters the registered surface; unknown groups raise.
- Annotations are populated on every tool (the security signal Claude
  Code's permission prompt depends on).
"""

from __future__ import annotations

import asyncio

import pytest
from mcp.server.fastmcp import FastMCP

from sharepoint_mcp.auth.flow import SharepointConsentNotConfiguredError
from sharepoint_mcp.server import (
    ALL_TOOL_GROUPS,
    SharepointToolGroupsError,
    parse_tool_groups,
    register_auth_tools,
    register_drive_read_tools,
    register_drive_write_tools,
    register_list_read_tools,
    register_list_write_tools,
    register_search_tools,
    register_share_read_tools,
    register_share_write_tools,
    register_site_tools,
    writes_enabled,
)


def _list_tool_names(server: FastMCP) -> set[str]:
    """Synchronously fetch tool names from a FastMCP server."""
    return {t.name for t in asyncio.run(server.list_tools())}


# ---------------------------------------------------------------------
# writes_enabled — strict env-var parsing
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
    """Only exactly 'true' / 'false' accepted. Legacy v0.4 truthy values
    (1/yes/on) and any other string raise."""
    monkeypatch.setenv("SP_ALLOW_WRITES", value)
    with pytest.raises(SharepointConsentNotConfiguredError, match="SP_ALLOW_WRITES"):
        writes_enabled()


def test_writes_enabled_unset_raises(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("SP_ALLOW_WRITES", raising=False)
    with pytest.raises(SharepointConsentNotConfiguredError, match="not set"):
        writes_enabled()


# ---------------------------------------------------------------------
# SP_TOOL_GROUPS parsing
# ---------------------------------------------------------------------


def test_parse_tool_groups_default_returns_all() -> None:
    assert parse_tool_groups(None) == set(ALL_TOOL_GROUPS)
    assert parse_tool_groups("") == set(ALL_TOOL_GROUPS)
    assert parse_tool_groups("   ") == set(ALL_TOOL_GROUPS)


def test_parse_tool_groups_subset() -> None:
    assert parse_tool_groups("drive,search") == {"drive", "search", "auth"}


def test_parse_tool_groups_auth_always_included() -> None:
    """Even if the operator omits `auth`, it's forced in — every other call needs it."""
    assert "auth" in parse_tool_groups("drive")


def test_parse_tool_groups_whitespace_and_case_tolerant() -> None:
    assert parse_tool_groups(" Drive , SEARCH ") == {"drive", "search", "auth"}


def test_parse_tool_groups_unknown_raises() -> None:
    with pytest.raises(SharepointToolGroupsError, match="unknown group"):
        parse_tool_groups("drive,bogus,search")


def test_parse_tool_groups_multiple_unknown_all_listed() -> None:
    """The error names every unknown group so the operator fixes them all at once."""
    with pytest.raises(SharepointToolGroupsError) as exc:
        parse_tool_groups("drive,foo,bar")
    msg = str(exc.value)
    assert "bar" in msg
    assert "foo" in msg


# ---------------------------------------------------------------------
# Per-category registration shapes
# ---------------------------------------------------------------------


def test_register_auth_tools_adds_auth_tools() -> None:
    server = FastMCP("test-auth")
    register_auth_tools(server)
    assert _list_tool_names(server) == {"sp_auth_begin", "sp_auth_status"}


def test_register_site_tools_adds_site_tools() -> None:
    server = FastMCP("test-site")
    register_site_tools(server)
    assert _list_tool_names(server) == {
        "sp_site_list",
        "sp_site_followed_list",
        "sp_site_drive_list",
        "sp_site_page_list",
        "sp_site_page_read",
        "sp_site_trash_list",
    }


def test_register_drive_read_tools_shape() -> None:
    server = FastMCP("test-drive-read")
    register_drive_read_tools(server)
    assert _list_tool_names(server) == {
        "sp_drive_folder_list",
        "sp_drive_file_read",
        "sp_drive_file_history",
        "sp_drive_file_version_get",
        "sp_drive_change_track",
        "sp_drive_checkout_list",
    }


def test_register_drive_write_tools_shape() -> None:
    server = FastMCP("test-drive-write")
    register_drive_write_tools(server)
    assert _list_tool_names(server) == {
        "sp_drive_folder_create",
        "sp_drive_file_upload",
        "sp_drive_file_delete",
        "sp_drive_file_move",
        "sp_drive_file_copy",
        "sp_drive_file_metadata",
        "sp_drive_file_checkout",
        "sp_drive_file_checkin",
        "sp_drive_file_checkout_discard",
        "sp_drive_file_checkout_bulk",
        "sp_drive_file_checkin_bulk",
    }


def test_register_list_read_tools_shape() -> None:
    server = FastMCP("test-list-read")
    register_list_read_tools(server)
    assert _list_tool_names(server) == {
        "sp_list_list",
        "sp_list_column_list",
        "sp_list_item_list",
        "sp_list_item_get",
    }


def test_register_list_write_tools_shape() -> None:
    server = FastMCP("test-list-write")
    register_list_write_tools(server)
    assert _list_tool_names(server) == {
        "sp_list_item_create",
        "sp_list_item_update",
        "sp_list_item_delete",
    }


def test_register_share_read_tools_shape() -> None:
    server = FastMCP("test-share-read")
    register_share_read_tools(server)
    assert _list_tool_names(server) == {
        "sp_share_link_list",
        "sp_share_permission_list",
    }


def test_register_share_write_tools_shape() -> None:
    server = FastMCP("test-share-write")
    register_share_write_tools(server)
    assert _list_tool_names(server) == {
        "sp_share_link_create",
        "sp_share_link_revoke",
    }


def test_register_search_tools_shape() -> None:
    server = FastMCP("test-search")
    register_search_tools(server)
    assert _list_tool_names(server) == {"sp_search_query"}


# ---------------------------------------------------------------------
# Annotations are populated everywhere
# ---------------------------------------------------------------------


def test_read_tools_have_readonly_annotation() -> None:
    """Every read tool must have readOnlyHint=True so Claude Code's prompt is right."""
    server = FastMCP("all-read")
    register_site_tools(server)
    register_drive_read_tools(server)
    register_list_read_tools(server)
    register_share_read_tools(server)
    register_search_tools(server)
    tools = asyncio.run(server.list_tools())
    for tool in tools:
        assert tool.annotations is not None, f"{tool.name} missing annotations"
        assert tool.annotations.readOnlyHint is True, (
            f"{tool.name} should be readOnlyHint=True; got {tool.annotations}"
        )


def test_checkout_tool_has_correct_annotation_pair() -> None:
    """sp_drive_file_checkout is non-destructive (creates a lock, doesn't damage
    data) but is NOT read-only — the annotation pair distinguishes."""
    server = FastMCP("test-checkout")
    register_drive_write_tools(server)
    [checkout_tool] = [
        t for t in asyncio.run(server.list_tools()) if t.name == "sp_drive_file_checkout"
    ]
    assert checkout_tool.annotations is not None
    assert checkout_tool.annotations.readOnlyHint is False
    assert checkout_tool.annotations.destructiveHint is False  # acquires a lock, not destruction


# ---------------------------------------------------------------------
# Module-level _build_server() honours env config
# ---------------------------------------------------------------------


def test_build_server_default_groups_includes_everything(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("SP_ALLOW_WRITES", "true")
    monkeypatch.delenv("SP_TOOL_GROUPS", raising=False)
    from sharepoint_mcp.server import _build_server

    server = _build_server()
    names = _list_tool_names(server)
    # Spot-check one from each category
    assert "sp_auth_begin" in names
    assert "sp_site_list" in names
    assert "sp_drive_file_checkout" in names
    assert "sp_list_item_create" in names
    assert "sp_share_link_create" in names
    assert "sp_search_query" in names


def test_build_server_omits_writes_in_explicit_readonly_mode(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Explicit SP_ALLOW_WRITES=false → no write tools, read tools still there."""
    monkeypatch.setenv("SP_ALLOW_WRITES", "false")
    monkeypatch.delenv("SP_TOOL_GROUPS", raising=False)
    from sharepoint_mcp.server import _build_server

    server = _build_server()
    names = _list_tool_names(server)
    assert "sp_drive_file_checkout" not in names
    assert "sp_search_query" in names
    assert "sp_drive_file_read" in names


def test_build_server_refuses_when_writes_env_unset(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("SP_ALLOW_WRITES", raising=False)
    from sharepoint_mcp.server import _build_server

    with pytest.raises(SharepointConsentNotConfiguredError, match="not set"):
        _build_server()


def test_build_server_refuses_on_legacy_truthy(monkeypatch: pytest.MonkeyPatch) -> None:
    """Legacy `SP_ALLOW_WRITES=yes` is rejected — must be explicit true/false."""
    monkeypatch.setenv("SP_ALLOW_WRITES", "yes")
    from sharepoint_mcp.server import _build_server

    with pytest.raises(SharepointConsentNotConfiguredError, match="SP_ALLOW_WRITES"):
        _build_server()


def test_build_server_with_tool_groups_filter(monkeypatch: pytest.MonkeyPatch) -> None:
    """SP_TOOL_GROUPS=drive,search → only those + auth registered."""
    monkeypatch.setenv("SP_ALLOW_WRITES", "true")
    monkeypatch.setenv("SP_TOOL_GROUPS", "drive,search")
    from sharepoint_mcp.server import _build_server

    server = _build_server()
    names = _list_tool_names(server)
    # In:
    assert "sp_auth_begin" in names
    assert "sp_drive_file_checkout" in names
    assert "sp_search_query" in names
    # Out:
    assert "sp_site_list" not in names
    assert "sp_list_item_create" not in names
    assert "sp_share_link_create" not in names


def test_build_server_unknown_tool_group_raises(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("SP_ALLOW_WRITES", "true")
    monkeypatch.setenv("SP_TOOL_GROUPS", "drive,typo")
    from sharepoint_mcp.server import _build_server

    with pytest.raises(SharepointToolGroupsError, match="unknown group"):
        _build_server()


def test_build_server_tool_groups_orthogonal_to_writes(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """SP_TOOL_GROUPS picks the group; SP_ALLOW_WRITES decides whether
    that group's write tools register."""
    monkeypatch.setenv("SP_ALLOW_WRITES", "false")
    monkeypatch.setenv("SP_TOOL_GROUPS", "drive")
    from sharepoint_mcp.server import _build_server

    server = _build_server()
    names = _list_tool_names(server)
    # Drive read tools in, drive write tools out:
    assert "sp_drive_file_read" in names
    assert "sp_drive_file_checkout" not in names
