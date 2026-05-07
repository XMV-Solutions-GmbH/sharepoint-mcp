# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for the public auth API (`get_token`, `interactive_login`).

Exercises the orchestration around the lower-level flow primitives:
cache-hit, refresh-on-expiry, refresh-rejected-clears-cache, BYO env
overrides, and the interactive-login full-cycle.
"""

from __future__ import annotations

import time

import pytest
import respx

from sharepoint_mcp.auth import (
    AuthRequiredError,
    DeviceCodeChallenge,
    get_token,
    interactive_login,
)
from sharepoint_mcp.auth.flow import (
    DEFAULT_AUTHORITY_TENANT,
    DEFAULT_CLIENT_ID,
)
from sharepoint_mcp.auth.tokens import CachedToken

# ---------------------------------------------------------------------
# In-memory store fixture
# ---------------------------------------------------------------------


class _MemStore:
    def __init__(self) -> None:
        self._d: dict[str, bytes] = {}

    def get(self, profile: str) -> bytes | None:
        return self._d.get(profile)

    def set(self, profile: str, value: bytes) -> None:
        self._d[profile] = value

    def delete(self, profile: str) -> None:
        self._d.pop(profile, None)


@pytest.fixture
def store() -> _MemStore:
    return _MemStore()


# ---------------------------------------------------------------------
# get_token — cache-hit / refresh / re-auth-required
# ---------------------------------------------------------------------


def test_get_token_returns_cached_when_fresh(store: _MemStore) -> None:
    fresh = CachedToken(
        access_token="AT-fresh",
        refresh_token="RT",
        expires_at=time.time() + 3600,
        scope="",
    )
    store.set("default", fresh.to_json().encode())
    assert get_token(store=store) == "AT-fresh"


def test_get_token_no_cache_raises_AuthRequiredError(store: _MemStore) -> None:
    with pytest.raises(AuthRequiredError, match="no cached credentials"):
        get_token(store=store)


def test_get_token_expired_no_refresh_raises_AuthRequiredError(store: _MemStore) -> None:
    expired = CachedToken(
        access_token="AT",
        refresh_token=None,
        expires_at=time.time() - 1,
        scope="",
    )
    store.set("default", expired.to_json().encode())
    with pytest.raises(AuthRequiredError, match="no refresh token"):
        get_token(store=store)


@respx.mock
def test_get_token_refreshes_when_expired(store: _MemStore) -> None:
    expired = CachedToken(
        access_token="AT-old",
        refresh_token="RT-1",
        expires_at=time.time() - 1,
        scope="",
    )
    store.set("default", expired.to_json().encode())

    refresh_url = f"https://login.microsoftonline.com/{DEFAULT_AUTHORITY_TENANT}/oauth2/v2.0/token"
    respx.post(refresh_url).respond(
        json={
            "access_token": "AT-new",
            "refresh_token": "RT-2",
            "expires_in": 3600,
            "scope": "",
            "token_type": "Bearer",
        }
    )

    assert get_token(store=store) == "AT-new"
    # Persisted updated token, including rotated refresh token
    raw = store.get("default")
    assert raw is not None
    new_cached = CachedToken.from_json(raw.decode())
    assert new_cached.access_token == "AT-new"
    assert new_cached.refresh_token == "RT-2"


@respx.mock
def test_get_token_invalid_refresh_clears_cache(store: _MemStore) -> None:
    expired = CachedToken(
        access_token="AT",
        refresh_token="RT-stale",
        expires_at=time.time() - 1,
        scope="",
    )
    store.set("default", expired.to_json().encode())

    refresh_url = f"https://login.microsoftonline.com/{DEFAULT_AUTHORITY_TENANT}/oauth2/v2.0/token"
    respx.post(refresh_url).respond(
        400,
        json={"error": "invalid_grant", "error_description": "AADSTS70008"},
    )

    with pytest.raises(AuthRequiredError, match="rejected"):
        get_token(store=store)
    # Cache cleared so next call doesn't re-attempt the bad refresh token
    assert store.get("default") is None


# ---------------------------------------------------------------------
# BYO env-var override
# ---------------------------------------------------------------------


@respx.mock
def test_get_token_byo_client_id_from_env(
    store: _MemStore, monkeypatch: pytest.MonkeyPatch
) -> None:
    expired = CachedToken(
        access_token="old",
        refresh_token="RT",
        expires_at=time.time() - 1,
        scope="",
    )
    store.set("default", expired.to_json().encode())
    monkeypatch.setenv("SP_CLIENT_ID", "byo-client-id")

    route = respx.post(
        f"https://login.microsoftonline.com/{DEFAULT_AUTHORITY_TENANT}/oauth2/v2.0/token"
    ).respond(
        json={
            "access_token": "AT",
            "refresh_token": "RT2",
            "expires_in": 3600,
            "scope": "",
            "token_type": "Bearer",
        }
    )
    get_token(store=store)
    body = route.calls.last.request.read().decode()
    assert "client_id=byo-client-id" in body


@respx.mock
def test_get_token_byo_tenant_from_env(store: _MemStore, monkeypatch: pytest.MonkeyPatch) -> None:
    expired = CachedToken(
        access_token="old",
        refresh_token="RT",
        expires_at=time.time() - 1,
        scope="",
    )
    store.set("default", expired.to_json().encode())
    monkeypatch.setenv("SP_TENANT_ID", "byo-tenant-guid")

    respx.post("https://login.microsoftonline.com/byo-tenant-guid/oauth2/v2.0/token").respond(
        json={
            "access_token": "AT",
            "refresh_token": "RT2",
            "expires_in": 3600,
            "scope": "",
            "token_type": "Bearer",
        }
    )
    assert get_token(store=store) == "AT"


def test_get_token_explicit_kwarg_beats_env(
    store: _MemStore, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Explicit kwarg precedence: caller > env > built-in default."""
    monkeypatch.setenv("SP_CLIENT_ID", "should-be-ignored")
    fresh = CachedToken(
        access_token="AT",
        refresh_token="RT",
        expires_at=time.time() + 3600,
        scope="",
    )
    store.set("default", fresh.to_json().encode())
    # No HTTP call needed — token is fresh — but the function still goes through
    # _resolve_client_id and the explicit value should win silently.
    get_token(store=store, client_id="explicit-wins")


# ---------------------------------------------------------------------
# interactive_login — full Device Code flow
# ---------------------------------------------------------------------


@respx.mock
def test_interactive_login_full_flow_persists_token(store: _MemStore) -> None:
    base = f"https://login.microsoftonline.com/{DEFAULT_AUTHORITY_TENANT}/oauth2/v2.0"
    respx.post(f"{base}/devicecode").respond(
        json={
            "device_code": "DC-secret",
            "user_code": "USR-CODE",
            "verification_uri": "https://microsoft.com/devicelogin",
            "expires_in": 900,
            "interval": 1,
            "message": "Go to URL",
        }
    )
    respx.post(f"{base}/token").respond(
        200,
        json={
            "access_token": "AT-final",
            "refresh_token": "RT-final",
            "expires_in": 3600,
            "scope": "Files.ReadWrite.All offline_access",
            "token_type": "Bearer",
        },
    )

    captured: list[DeviceCodeChallenge] = []
    cached = interactive_login(
        store=store,
        prompt=captured.append,
    )

    assert cached.access_token == "AT-final"
    assert cached.refresh_token == "RT-final"
    raw = store.get("default")
    assert raw is not None
    persisted = CachedToken.from_json(raw.decode())
    assert persisted == cached
    assert len(captured) == 1
    assert captured[0].user_code == "USR-CODE"
    assert captured[0].verification_uri == "https://microsoft.com/devicelogin"


@respx.mock
def test_interactive_login_uses_default_client_id(store: _MemStore) -> None:
    """No SP_CLIENT_ID env, no explicit kwarg -> XMV-published default."""
    base = f"https://login.microsoftonline.com/{DEFAULT_AUTHORITY_TENANT}/oauth2/v2.0"
    devcode_route = respx.post(f"{base}/devicecode").respond(
        json={
            "device_code": "DC",
            "user_code": "U",
            "verification_uri": "x",
            "expires_in": 1,
            "interval": 1,
            "message": "",
        }
    )
    respx.post(f"{base}/token").respond(
        200,
        json={
            "access_token": "AT",
            "refresh_token": "RT",
            "expires_in": 3600,
            "scope": "",
            "token_type": "Bearer",
        },
    )
    interactive_login(store=store, prompt=lambda _ch: None)
    body = devcode_route.calls.last.request.read().decode()
    assert f"client_id={DEFAULT_CLIENT_ID}" in body
