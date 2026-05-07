# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for the service-principal / client-credentials auth path (#40)."""

from __future__ import annotations

import time

import httpx
import pytest
import respx

from sharepoint_mcp.auth.flow import AUTHORITY_BASE
from sharepoint_mcp.auth.service_principal import (
    SERVICE_PRINCIPAL_SCOPE,
    ServicePrincipalConfigError,
    acquire_app_only_token,
    get_app_only_token,
    is_service_principal_mode,
    reset_cache,
)


@pytest.fixture(autouse=True)
def _clear_cache_between_tests() -> None:
    reset_cache()


# ---------------------------------------------------------------------
# Mode detection
# ---------------------------------------------------------------------


@pytest.mark.parametrize(
    "mode",
    [
        "service-principal",
        "service_principal",
        "client-credentials",
        "client_credentials",
        "app-only",
        "app_only",
        "SERVICE-PRINCIPAL",
    ],
)
def test_is_service_principal_mode_explicit_aliases(
    monkeypatch: pytest.MonkeyPatch, mode: str
) -> None:
    monkeypatch.setenv("SP_AUTH_MODE", mode)
    monkeypatch.delenv("SP_CLIENT_SECRET", raising=False)
    assert is_service_principal_mode() is True


@pytest.mark.parametrize("mode", ["delegated", "user", "device-code", "device_code"])
def test_is_service_principal_mode_explicit_delegated_wins(
    monkeypatch: pytest.MonkeyPatch, mode: str
) -> None:
    """Even with SP_CLIENT_SECRET set, explicit delegated wins."""
    monkeypatch.setenv("SP_AUTH_MODE", mode)
    monkeypatch.setenv("SP_CLIENT_SECRET", "anything")
    assert is_service_principal_mode() is False


def test_is_service_principal_mode_auto_detect_via_secret(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("SP_AUTH_MODE", raising=False)
    monkeypatch.setenv("SP_CLIENT_SECRET", "shh")
    assert is_service_principal_mode() is True


def test_is_service_principal_mode_default_false(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("SP_AUTH_MODE", raising=False)
    monkeypatch.delenv("SP_CLIENT_SECRET", raising=False)
    assert is_service_principal_mode() is False


def test_is_service_principal_mode_empty_secret_does_not_auto_detect(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.delenv("SP_AUTH_MODE", raising=False)
    monkeypatch.setenv("SP_CLIENT_SECRET", "")
    assert is_service_principal_mode() is False


# ---------------------------------------------------------------------
# acquire_app_only_token — wire shape
# ---------------------------------------------------------------------


@respx.mock
def test_acquire_app_only_token_uses_client_credentials_grant() -> None:
    tenant = "tenant-guid"
    route = respx.post(f"{AUTHORITY_BASE}/{tenant}/oauth2/v2.0/token").respond(
        json={
            "token_type": "Bearer",
            "expires_in": 3599,
            "access_token": "AT-app",
            "scope": "https://graph.microsoft.com/.default",
        }
    )
    cached = acquire_app_only_token(
        client_id="cid",
        client_secret="secret",
        tenant=tenant,
    )
    assert cached.access_token == "AT-app"
    assert cached.refresh_token is None
    body = route.calls.last.request.read().decode()
    assert "grant_type=client_credentials" in body
    assert "client_id=cid" in body
    assert "client_secret=secret" in body
    assert "scope=https" in body  # SERVICE_PRINCIPAL_SCOPE


@respx.mock
def test_acquire_app_only_token_propagates_4xx() -> None:
    tenant = "t"
    respx.post(f"{AUTHORITY_BASE}/{tenant}/oauth2/v2.0/token").respond(
        401, json={"error": "invalid_client"}
    )
    with pytest.raises(httpx.HTTPStatusError):
        acquire_app_only_token(client_id="cid", client_secret="bad", tenant=tenant)


def test_service_principal_scope_is_dot_default() -> None:
    """The /.default suffix is what tells AAD to issue a token covering all
    consented Application permissions. Don't change this without a code search."""
    assert SERVICE_PRINCIPAL_SCOPE == "https://graph.microsoft.com/.default"


# ---------------------------------------------------------------------
# get_app_only_token — env-var validation + caching
# ---------------------------------------------------------------------


def test_get_app_only_token_requires_client_id(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("SP_CLIENT_ID", raising=False)
    monkeypatch.setenv("SP_CLIENT_SECRET", "s")
    monkeypatch.setenv("SP_TENANT_ID", "t")
    with pytest.raises(ServicePrincipalConfigError, match="SP_CLIENT_ID"):
        get_app_only_token()


def test_get_app_only_token_requires_client_secret(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("SP_CLIENT_ID", "c")
    monkeypatch.delenv("SP_CLIENT_SECRET", raising=False)
    monkeypatch.setenv("SP_TENANT_ID", "t")
    with pytest.raises(ServicePrincipalConfigError, match="SP_CLIENT_SECRET"):
        get_app_only_token()


def test_get_app_only_token_requires_tenant(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("SP_CLIENT_ID", "c")
    monkeypatch.setenv("SP_CLIENT_SECRET", "s")
    monkeypatch.delenv("SP_TENANT_ID", raising=False)
    with pytest.raises(ServicePrincipalConfigError, match="SP_TENANT_ID"):
        get_app_only_token()


def test_get_app_only_token_lists_all_missing_at_once(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Missing all three env vars: error message names all three so the user
    fixes them in one go, not three round-trips."""
    monkeypatch.delenv("SP_CLIENT_ID", raising=False)
    monkeypatch.delenv("SP_CLIENT_SECRET", raising=False)
    monkeypatch.delenv("SP_TENANT_ID", raising=False)
    with pytest.raises(ServicePrincipalConfigError) as exc_info:
        get_app_only_token()
    msg = str(exc_info.value)
    assert "SP_CLIENT_ID" in msg
    assert "SP_CLIENT_SECRET" in msg
    assert "SP_TENANT_ID" in msg


@respx.mock
def test_get_app_only_token_caches_until_expiry(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("SP_CLIENT_ID", "c")
    monkeypatch.setenv("SP_CLIENT_SECRET", "s")
    monkeypatch.setenv("SP_TENANT_ID", "t")
    route = respx.post(f"{AUTHORITY_BASE}/t/oauth2/v2.0/token").respond(
        json={"access_token": "AT-1", "expires_in": 3600, "scope": ""},
    )
    assert get_app_only_token() == "AT-1"
    assert get_app_only_token() == "AT-1"
    assert get_app_only_token() == "AT-1"
    assert route.call_count == 1  # cached after first acquire


@respx.mock
def test_get_app_only_token_reacquires_after_expiry(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("SP_CLIENT_ID", "c")
    monkeypatch.setenv("SP_CLIENT_SECRET", "s")
    monkeypatch.setenv("SP_TENANT_ID", "t")
    # First response: token expired 60 seconds ago (negative expires_in
    # gives an expires_at in the past)
    route = respx.post(f"{AUTHORITY_BASE}/t/oauth2/v2.0/token").mock(
        side_effect=[
            httpx.Response(200, json={"access_token": "AT-old", "expires_in": -60, "scope": ""}),
            httpx.Response(200, json={"access_token": "AT-new", "expires_in": 3600, "scope": ""}),
        ],
    )
    first = get_app_only_token()
    second = get_app_only_token()
    assert first == "AT-old"
    assert second == "AT-new"
    assert route.call_count == 2


@respx.mock
def test_get_app_only_token_separate_cache_per_tenant(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Two distinct (client_id, tenant) pairs cache separately."""
    monkeypatch.setenv("SP_CLIENT_ID", "c1")
    monkeypatch.setenv("SP_CLIENT_SECRET", "s")
    monkeypatch.setenv("SP_TENANT_ID", "t1")
    respx.post(f"{AUTHORITY_BASE}/t1/oauth2/v2.0/token").respond(
        json={"access_token": "AT-T1", "expires_in": 3600, "scope": ""},
    )
    monkeypatch.setenv("SP_TENANT_ID", "t2")
    respx.post(f"{AUTHORITY_BASE}/t2/oauth2/v2.0/token").respond(
        json={"access_token": "AT-T2", "expires_in": 3600, "scope": ""},
    )
    monkeypatch.setenv("SP_TENANT_ID", "t1")
    assert get_app_only_token() == "AT-T1"
    monkeypatch.setenv("SP_TENANT_ID", "t2")
    assert get_app_only_token() == "AT-T2"
    monkeypatch.setenv("SP_TENANT_ID", "t1")
    # Hits cache, does not re-acquire
    assert get_app_only_token() == "AT-T1"


def test_reset_cache_clears_entries(monkeypatch: pytest.MonkeyPatch) -> None:
    """Sanity: reset_cache() drops everything so test isolation works."""
    from sharepoint_mcp.auth import service_principal as sp

    sp._app_token_cache[("c", "t")] = type(sp.acquire_app_only_token).__call__  # type: ignore[assignment]
    # The above is gibberish but proves the cache has at least one entry;
    # the reset must clear it.
    sp._app_token_cache[("c", "t")] = sp.CachedToken(  # type: ignore[attr-defined]
        access_token="x",
        refresh_token=None,
        expires_at=time.time() + 1000,
        scope="",
    )
    assert sp._app_token_cache  # populated
    reset_cache()
    assert sp._app_token_cache == {}
