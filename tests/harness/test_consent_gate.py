# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""End-to-end harness test for the v0.5 strict consent gate.

Confirms that the `SP_ALLOW_WRITES` env-var validation, OAuth scope
resolution, and `_build_server()` tool-registration all flow through
correctly when wired up against the real harness profile.

The harness here is more about the **integration of the consent
machinery with the real auth stack** than about hitting Graph: we
verify the harness token is still usable, the scope-resolution
matches the env-var decision, and the server build refuses unset
env vars verbatim.

Skips gracefully if the harness profile token cache is empty (same
contract as the rest of `tests/harness/`).
"""

from __future__ import annotations

import asyncio

import httpx
import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.auth.flow import (
    SharepointConsentNotConfiguredError,
    resolve_scopes,
)

HARNESS_PROFILE = "harness"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _harness_token_or_skip() -> str:
    try:
        return get_token(HARNESS_PROFILE)
    except AuthRequiredError as exc:
        pytest.skip(
            f"Harness credentials not available: {exc}. "
            "Run `uv run mcp-server-sharepoint login --profile harness` to populate.",
        )


def test_consent_gate_writes_true_resolves_to_readwrite_scopes(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """SP_ALLOW_WRITES=true → resolve_scopes() returns the ReadWrite
    variants. Pinned at the harness layer (and not just the unit
    layer) because the scope tuple is what gets sent to Microsoft
    Identity at login time — any drift between strings has real
    user-visible consequences on the consent screen."""
    monkeypatch.setenv("SP_ALLOW_WRITES", "true")
    scopes = resolve_scopes()
    assert "Files.ReadWrite.All" in scopes
    assert "Sites.ReadWrite.All" in scopes
    assert "Files.Read.All" not in scopes
    assert "Sites.Read.All" not in scopes


def test_consent_gate_writes_false_resolves_to_readonly_scopes(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """SP_ALLOW_WRITES=false → resolve_scopes() drops Files.ReadWrite.All
    and Sites.ReadWrite.All in favour of their `.Read.All` variants.
    The consent screen on a fresh login will reflect this — operators
    in read-only mode never see a "this app can modify your files"
    line."""
    monkeypatch.setenv("SP_ALLOW_WRITES", "false")
    scopes = resolve_scopes()
    assert "Files.Read.All" in scopes
    assert "Sites.Read.All" in scopes
    assert "Files.ReadWrite.All" not in scopes
    assert "Sites.ReadWrite.All" not in scopes


def test_consent_gate_unset_raises_with_actionable_message(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """The error message is the user-facing onboarding doc — it must
    name the env var, both accepted values, and the file the operator
    edits."""
    monkeypatch.delenv("SP_ALLOW_WRITES", raising=False)
    with pytest.raises(SharepointConsentNotConfiguredError) as exc_info:
        resolve_scopes()
    msg = str(exc_info.value)
    assert "SP_ALLOW_WRITES" in msg
    assert '"true"' in msg
    assert '"false"' in msg
    assert ".mcp.json" in msg


def test_consent_gate_harness_token_still_works_against_real_graph(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """The harness token was minted under v0.4 default scopes (the
    full ReadWrite set). After the v0.5 scope split it stays valid —
    Graph accepts a broader-scope token even when the client requests
    only narrower scopes on the NEXT refresh. Proves the upgrade
    doesn't strand existing operator setups."""
    monkeypatch.setenv("SP_ALLOW_WRITES", "true")
    token = _harness_token_or_skip()

    with httpx.Client(timeout=10.0) as client:
        response = client.get(
            f"{GRAPH_BASE}/me",
            headers={"Authorization": f"Bearer {token}"},
        )
    response.raise_for_status()
    assert response.json().get("id"), "Graph /me must return an id"


def test_consent_gate_build_server_writes_true_registers_write_tools(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """End-to-end: with the harness token cached and SP_ALLOW_WRITES=true,
    _build_server() succeeds and exposes the gated write tools. This
    is the closest we can get to "operator just installed the server"
    without re-running interactive login in CI."""
    _harness_token_or_skip()  # gate on harness availability
    monkeypatch.setenv("SP_ALLOW_WRITES", "true")
    from sharepoint_mcp.server import _build_server

    server = _build_server()
    names = {t.name for t in asyncio.run(server.list_tools())}
    assert "sp_open" in names
    assert "sp_save" in names
    assert "sp_release" in names
    # Read tools always there
    assert "sp_search" in names


def test_consent_gate_build_server_writes_false_omits_write_tools(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """End-to-end: SP_ALLOW_WRITES=false → no write tools, read tools intact."""
    _harness_token_or_skip()
    monkeypatch.setenv("SP_ALLOW_WRITES", "false")
    from sharepoint_mcp.server import _build_server

    server = _build_server()
    names = {t.name for t in asyncio.run(server.list_tools())}
    assert "sp_open" not in names
    assert "sp_save" not in names
    assert "sp_search" in names
