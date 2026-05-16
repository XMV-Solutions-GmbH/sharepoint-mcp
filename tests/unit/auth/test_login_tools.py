# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_login_begin / sp_login_status (#75, #76).

Edge cases deliberately covered:
- Idempotent re-call returns existing pending session unchanged.
- force=True cancels the prior task and returns a fresh session.
- Concurrent simultaneous calls — first-write-wins.
- Service-principal mode → sp_login_begin refuses with typed error.
- Microsoft refuses /devicecode → DeviceCodeRequestFailedError raised.
- public_view shape: device_code is NEVER in the output (security).
- login_status: all five states (signed_in via active probe, pending,
  none, terminal-success surfaces, terminal-failed surfaces error).
- login_status: token cache is corrupt / unreadable → falls through to none.
- login_status: refresh-on-expired-access path → status=signed_in.
- login_status: refresh fails → status=none.
- login_status: missing refresh_token on expired access → none (no infinite retry).
- _extract_upn_from_jwt: handles malformed JWTs without raising.
"""

from __future__ import annotations

import asyncio
import base64
import json
import time
from collections.abc import Iterator
from pathlib import Path
from types import SimpleNamespace

import httpx
import pytest
from mcp_microsoft_graph_auth import LoginSessionRegistry, TokenStoreLockTimeoutError

from sharepoint_mcp.auth import login_tools
from sharepoint_mcp.auth.login_tools import (
    ConcurrentLoginAttemptError,
    DeviceCodeRequestFailedError,
    ServicePrincipalActiveError,
    _extract_upn_from_jwt,
    _none,
    login_begin,
    login_status,
)
from sharepoint_mcp.auth.tokens import CachedToken


@pytest.fixture(autouse=True)
def _isolate_registry() -> Iterator[None]:
    """Each test gets a fresh registry — the module-level _REGISTRY is
    process-shared which would leak state between tests."""
    fresh = LoginSessionRegistry()
    saved = login_tools._REGISTRY
    login_tools._REGISTRY = fresh
    try:
        yield
    finally:
        login_tools._REGISTRY = saved


@pytest.fixture(autouse=True)
def _clear_sp_env(monkeypatch: pytest.MonkeyPatch) -> None:
    """Reset SP_AUTH_MODE / SP_CLIENT_SECRET so service-principal mode
    detection is deterministic."""
    monkeypatch.delenv("SP_AUTH_MODE", raising=False)
    monkeypatch.delenv("SP_CLIENT_SECRET", raising=False)


def _make_jwt(claims: dict[str, str]) -> str:
    """Build a minimal JWT-shaped token with the given claims for tests."""
    header = base64.urlsafe_b64encode(b'{"alg":"none"}').rstrip(b"=").decode()
    payload = base64.urlsafe_b64encode(json.dumps(claims).encode()).rstrip(b"=").decode()
    return f"{header}.{payload}.signature"


# ---------------------------------------------------------------------
# login_begin — happy path + idempotency
# ---------------------------------------------------------------------


@pytest.mark.asyncio
async def test_login_begin_returns_user_facing_fields_on_first_call(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        login_tools,
        "request_device_code",
        lambda **kw: (
            "secret-device-code",
            type(
                "C",
                (),
                {
                    "user_code": "ABC-123",
                    "verification_uri": "https://x/devicelogin",
                    "verification_uri_complete": None,
                    "expires_at": time.time() + 900,
                    "interval": 5,
                },
            )(),
        ),
    )
    # Stub the poll loop so the asyncio task doesn't actually run
    monkeypatch.setattr(login_tools, "_poll_loop", _stub_poll_loop)
    out = await login_begin(profile="default")
    assert out["status"] == "pending"
    assert out["profile"] == "default"
    assert out["user_code"] == "ABC-123"
    assert out["verification_url"] == "https://x/devicelogin"
    assert out["interval_s"] == 5
    assert "expires_at" in out
    assert out["session_id"]


@pytest.mark.asyncio
async def test_login_begin_does_not_leak_device_code_in_output(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """The single most important security invariant of public_view."""
    monkeypatch.setattr(
        login_tools,
        "request_device_code",
        lambda **kw: (
            "DC-VERY-SECRET-DO-NOT-LEAK",
            type(
                "C",
                (),
                {
                    "user_code": "U",
                    "verification_uri": "x",
                    "verification_uri_complete": None,
                    "expires_at": time.time() + 900,
                    "interval": 5,
                },
            )(),
        ),
    )
    monkeypatch.setattr(login_tools, "_poll_loop", _stub_poll_loop)
    out = await login_begin(profile="default")
    assert "device_code" not in out
    assert "VERY-SECRET" not in str(out)


@pytest.mark.asyncio
async def test_login_begin_idempotent_returns_existing_pending(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    counter = {"n": 0}

    def _request_device_code(**kw: object) -> tuple[str, object]:
        counter["n"] += 1
        return (
            f"DC-{counter['n']}",
            type(
                "C",
                (),
                {
                    "user_code": f"USER-{counter['n']}",
                    "verification_uri": "x",
                    "verification_uri_complete": None,
                    "expires_at": time.time() + 900,
                    "interval": 5,
                },
            )(),
        )

    monkeypatch.setattr(login_tools, "request_device_code", _request_device_code)
    monkeypatch.setattr(login_tools, "_poll_loop", _stub_poll_loop)
    first = await login_begin(profile="p")
    second = await login_begin(profile="p")
    assert first["session_id"] == second["session_id"]
    assert first["user_code"] == second["user_code"] == "USER-1"
    assert counter["n"] == 1  # second call did NOT trigger a fresh /devicecode


@pytest.mark.asyncio
async def test_login_begin_force_cancels_existing_and_starts_fresh(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    counter = {"n": 0}

    def _request_device_code(**kw: object) -> tuple[str, object]:
        counter["n"] += 1
        return (
            f"DC-{counter['n']}",
            type(
                "C",
                (),
                {
                    "user_code": f"USER-{counter['n']}",
                    "verification_uri": "x",
                    "verification_uri_complete": None,
                    "expires_at": time.time() + 900,
                    "interval": 5,
                },
            )(),
        )

    monkeypatch.setattr(login_tools, "request_device_code", _request_device_code)
    monkeypatch.setattr(login_tools, "_poll_loop", _stub_poll_loop)
    first = await login_begin(profile="p")
    second = await login_begin(profile="p", force=True)
    assert first["session_id"] != second["session_id"]
    assert first["user_code"] != second["user_code"]
    assert counter["n"] == 2


# ---------------------------------------------------------------------
# login_begin — service-principal mode + Graph errors
# ---------------------------------------------------------------------


@pytest.mark.asyncio
async def test_login_begin_in_service_principal_mode_raises(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setenv("SP_AUTH_MODE", "service-principal")
    with pytest.raises(ServicePrincipalActiveError):
        await login_begin(profile="default")


@pytest.mark.asyncio
async def test_login_begin_propagates_microsoft_devicecode_failure(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def _failing(**kw: object) -> object:
        raise httpx.ConnectError("network down")

    monkeypatch.setattr(login_tools, "request_device_code", _failing)
    with pytest.raises(DeviceCodeRequestFailedError, match="refused"):
        await login_begin(profile="default")


# ---------------------------------------------------------------------
# login_status — three states + recently-terminal surfaces
# ---------------------------------------------------------------------


@pytest.mark.asyncio
async def test_login_status_none_when_no_session_no_cached_token(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(login_tools, "get_token_store", lambda: _NullStore())
    out = await login_status(profile="x")
    assert out == {
        "status": "none",
        "profile": "x",
        "signed_in_user_upn": None,
        "error": None,
    }


@pytest.mark.asyncio
async def test_login_status_signed_in_when_cached_valid_token(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A user who logged in via CLI days ago has a valid token on disk;
    sp_login_status must surface signed_in directly, not 'none'."""
    cached = CachedToken(
        access_token=_make_jwt({"upn": "alice@x.com"}),
        refresh_token="rt",
        expires_at=time.time() + 3600,
        scope="",
    )
    store = _DictStore()
    store.set("default", cached.to_json().encode())
    monkeypatch.setattr(login_tools, "get_token_store", lambda: store)
    out = await login_status(profile="default")
    assert out["status"] == "signed_in"
    assert out["signed_in_user_upn"] == "alice@x.com"


@pytest.mark.asyncio
async def test_login_status_pending_from_in_memory_session(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        login_tools,
        "request_device_code",
        lambda **kw: (
            "dc",
            type(
                "C",
                (),
                {
                    "user_code": "MOBILE-OPTIMISED",
                    "verification_uri": "https://login/devicelogin",
                    "verification_uri_complete": None,
                    "expires_at": time.time() + 900,
                    "interval": 5,
                },
            )(),
        ),
    )
    monkeypatch.setattr(login_tools, "_poll_loop", _stub_poll_loop)
    await login_begin(profile="p")
    out = await login_status(profile="p")
    assert out["status"] == "pending"
    assert out["user_code"] == "MOBILE-OPTIMISED"
    assert out["verification_url"] == "https://login/devicelogin"
    assert out["time_remaining_s"] > 0


@pytest.mark.asyncio
async def test_login_status_signed_in_after_session_reaches_success(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Once the polling task finishes (status=success), sp_login_status
    surfaces 'signed_in' (not 'success' as an internal state)."""
    monkeypatch.setattr(
        login_tools,
        "request_device_code",
        lambda **kw: (
            "dc",
            type(
                "C",
                (),
                {
                    "user_code": "U",
                    "verification_uri": "x",
                    "verification_uri_complete": None,
                    "expires_at": time.time() + 900,
                    "interval": 5,
                },
            )(),
        ),
    )
    monkeypatch.setattr(login_tools, "_poll_loop", _stub_poll_loop)
    await login_begin(profile="p")
    # Simulate the poll task completing successfully
    session = login_tools._REGISTRY.get("p")
    assert session is not None
    session.status = "success"
    session.signed_in_user_upn = "bob@x.com"
    out = await login_status(profile="p")
    assert out["status"] == "signed_in"
    assert out["signed_in_user_upn"] == "bob@x.com"


@pytest.mark.asyncio
async def test_login_status_surfaces_terminal_expired_with_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Recently-terminal session: surface error, distinct from 'none'."""
    monkeypatch.setattr(
        login_tools,
        "request_device_code",
        lambda **kw: (
            "dc",
            type(
                "C",
                (),
                {
                    "user_code": "U",
                    "verification_uri": "x",
                    "verification_uri_complete": None,
                    "expires_at": time.time() + 900,
                    "interval": 5,
                },
            )(),
        ),
    )
    monkeypatch.setattr(login_tools, "_poll_loop", _stub_poll_loop)
    await login_begin(profile="p")
    session = login_tools._REGISTRY.get("p")
    assert session is not None
    session.status = "expired"
    session.error = {"code": "expired", "message": "device code expired"}
    out = await login_status(profile="p")
    assert out["status"] == "expired"
    assert out["error"] == {"code": "expired", "message": "device code expired"}
    assert out["signed_in_user_upn"] is None


@pytest.mark.asyncio
async def test_login_status_surfaces_terminal_failed(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        login_tools,
        "request_device_code",
        lambda **kw: (
            "dc",
            type(
                "C",
                (),
                {
                    "user_code": "U",
                    "verification_uri": "x",
                    "verification_uri_complete": None,
                    "expires_at": time.time() + 900,
                    "interval": 5,
                },
            )(),
        ),
    )
    monkeypatch.setattr(login_tools, "_poll_loop", _stub_poll_loop)
    await login_begin(profile="p")
    session = login_tools._REGISTRY.get("p")
    assert session is not None
    session.status = "failed"
    session.error = {"code": "access_denied", "message": "user refused"}
    out = await login_status(profile="p")
    assert out["status"] == "failed"
    assert out["error"]["code"] == "access_denied"


# ---------------------------------------------------------------------
# login_status — refresh-on-expired path
# ---------------------------------------------------------------------


@pytest.mark.asyncio
async def test_login_status_refreshes_expired_access_token(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    expired = CachedToken(
        access_token=_make_jwt({"upn": "old@x.com"}),
        refresh_token="rt-good",
        expires_at=time.time() - 60,
        scope="",
    )
    store = _DictStore()
    store.set("default", expired.to_json().encode())
    monkeypatch.setattr(login_tools, "get_token_store", lambda: store)

    new_token = CachedToken(
        access_token=_make_jwt({"upn": "fresh@x.com"}),
        refresh_token="rt-rotated",
        expires_at=time.time() + 3600,
        scope="",
    )
    monkeypatch.setattr(
        login_tools,
        "refresh_access_token",
        lambda **kw: new_token,
    )

    out = await login_status(profile="default")
    assert out["status"] == "signed_in"
    assert out["signed_in_user_upn"] == "fresh@x.com"


@pytest.mark.asyncio
async def test_login_status_returns_none_when_refresh_fails(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    expired = CachedToken(
        access_token=_make_jwt({"upn": "x"}),
        refresh_token="rt-stale",
        expires_at=time.time() - 60,
        scope="",
    )
    store = _DictStore()
    store.set("default", expired.to_json().encode())
    monkeypatch.setattr(login_tools, "get_token_store", lambda: store)

    def _fail(**kw: object) -> object:
        raise RuntimeError("refresh rejected")

    monkeypatch.setattr(login_tools, "refresh_access_token", _fail)
    out = await login_status(profile="default")
    assert out["status"] == "none"


@pytest.mark.asyncio
async def test_login_status_returns_none_when_expired_with_no_refresh_token(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """If access expired AND no refresh token is cached, don't loop —
    just return none so the agent prompts a fresh login."""
    expired_no_refresh = CachedToken(
        access_token=_make_jwt({"upn": "x"}),
        refresh_token=None,
        expires_at=time.time() - 60,
        scope="",
    )
    store = _DictStore()
    store.set("default", expired_no_refresh.to_json().encode())
    monkeypatch.setattr(login_tools, "get_token_store", lambda: store)
    out = await login_status(profile="default")
    assert out["status"] == "none"


# ---------------------------------------------------------------------
# login_status — defensive parsing
# ---------------------------------------------------------------------


@pytest.mark.asyncio
async def test_login_status_corrupt_token_cache_falls_through_to_none(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A corrupt JSON blob in the token cache shouldn't crash the agent.
    Falls through to 'none'; subsequent login_begin → store.set will
    overwrite the corruption."""
    store = _DictStore()
    store.set("default", b"this is not valid json {{{")
    monkeypatch.setattr(login_tools, "get_token_store", lambda: store)
    out = await login_status(profile="default")
    assert out["status"] == "none"


@pytest.mark.asyncio
async def test_login_status_token_store_init_failure_returns_none(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Token-store init can itself raise (e.g. encrypted-file with no
    passphrase). Agent gets 'none' so it can prompt; the underlying
    error will surface properly during login_begin."""

    def _raises() -> object:
        raise RuntimeError("no usable token store")

    monkeypatch.setattr(login_tools, "get_token_store", _raises)
    out = await login_status(profile="default")
    assert out["status"] == "none"


# ---------------------------------------------------------------------
# _extract_upn_from_jwt
# ---------------------------------------------------------------------


def test_extract_upn_prefers_upn_claim_over_preferred_username() -> None:
    token = _make_jwt({"upn": "primary@x", "preferred_username": "alt@x"})
    assert _extract_upn_from_jwt(token) == "primary@x"


def test_extract_upn_falls_back_to_preferred_username() -> None:
    token = _make_jwt({"preferred_username": "fallback@x"})
    assert _extract_upn_from_jwt(token) == "fallback@x"


def test_extract_upn_returns_none_when_neither_claim_present() -> None:
    token = _make_jwt({"sub": "some-sub-id"})
    assert _extract_upn_from_jwt(token) is None


def test_extract_upn_handles_malformed_jwt_without_raising() -> None:
    assert _extract_upn_from_jwt("not.a.jwt.really") is None
    assert _extract_upn_from_jwt("only-one-part") is None
    assert _extract_upn_from_jwt("") is None


def test_extract_upn_handles_invalid_base64_payload() -> None:
    assert _extract_upn_from_jwt("header.!!notbase64!!.sig") is None


def test_extract_upn_handles_non_json_payload() -> None:
    payload = base64.urlsafe_b64encode(b"plain text").rstrip(b"=").decode()
    assert _extract_upn_from_jwt(f"hdr.{payload}.sig") is None


# ---------------------------------------------------------------------
# login_begin — concurrent-lock guard (#77)
# ---------------------------------------------------------------------


@pytest.mark.asyncio
async def test_login_begin_raises_concurrent_error_when_lock_held(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """If the lock probe sees a concurrent holder, login_begin raises
    ConcurrentLoginAttemptError before touching /devicecode."""

    def _locked(path: object) -> None:
        raise TokenStoreLockTimeoutError(tmp_path / "login.lock", 0.1)

    monkeypatch.setattr(login_tools, "_probe_login_lock", _locked)
    with pytest.raises(ConcurrentLoginAttemptError, match="Another process"):
        await login_begin(profile="default")


@pytest.mark.asyncio
async def test_login_begin_concurrent_error_does_not_call_devicecode(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """No /devicecode HTTP round-trip is made when the lock probe fails."""
    called = {"n": 0}

    def _locked(path: object) -> None:
        raise TokenStoreLockTimeoutError(tmp_path / "login.lock", 0.1)

    def _count(**kw: object) -> object:
        called["n"] += 1
        raise AssertionError("should not be reached")

    monkeypatch.setattr(login_tools, "_probe_login_lock", _locked)
    monkeypatch.setattr(login_tools, "request_device_code", _count)
    with pytest.raises(ConcurrentLoginAttemptError):
        await login_begin(profile="default")
    assert called["n"] == 0


# ---------------------------------------------------------------------
# _login_lock_path
# ---------------------------------------------------------------------


def test_login_lock_path_structure(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """Lock file must be inside DEFAULT_CACHE_DIR / profile / login.lock,
    and the parent directory must be created on demand."""
    monkeypatch.setattr(login_tools, "DEFAULT_CACHE_DIR", tmp_path)
    path = login_tools._login_lock_path("my-profile")
    assert path == tmp_path / "my-profile" / "login.lock"
    assert path.parent.is_dir()


def test_login_lock_path_profile_isolation(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """Different profiles must produce different lock paths."""
    monkeypatch.setattr(login_tools, "DEFAULT_CACHE_DIR", tmp_path)
    assert login_tools._login_lock_path("a") != login_tools._login_lock_path("b")


# ---------------------------------------------------------------------
# _poll_loop — TokenStoreLockTimeoutError surfaces correctly
# ---------------------------------------------------------------------


@pytest.mark.asyncio
async def test_poll_loop_sets_failed_status_on_lock_timeout(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Narrow race: probe succeeded but another process grabbed the lock
    before the task started. _poll_loop must record status=failed with
    code=concurrent_login_attempt (not leave the session stuck as pending)."""

    def _locked(session: object, client_id: str, tenant: str) -> object:
        raise TokenStoreLockTimeoutError(Path("/tmp/login.lock"), 1200.0)

    monkeypatch.setattr(login_tools, "_sync_poll_with_lock", _locked)

    session = SimpleNamespace(
        status="pending",
        error=None,
        signed_in_user_upn=None,
        profile="p",
        device_code="dc",
        interval_s=5,
    )
    await login_tools._poll_loop(session, "client-id", "tenant")
    assert session.status == "failed"
    assert session.error is not None
    assert session.error["code"] == "concurrent_login_attempt"


# ---------------------------------------------------------------------
# _sync_poll_with_lock — happy path
# ---------------------------------------------------------------------


def test_sync_poll_with_lock_stores_token_and_returns_it(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """Happy path: poll_for_token returns a token; the token is written to
    the store and the same object is returned to the caller."""
    expected = CachedToken(
        access_token=_make_jwt({"upn": "stored@x.com"}),
        refresh_token="rt",
        expires_at=time.time() + 3600,
        scope="",
    )

    monkeypatch.setattr(login_tools, "poll_for_token", lambda **kw: expected)
    store = _DictStore()
    monkeypatch.setattr(login_tools, "get_token_store", lambda: store)
    monkeypatch.setattr(login_tools, "DEFAULT_CACHE_DIR", tmp_path)

    session = SimpleNamespace(profile="q", device_code="dc2", interval_s=5)
    result = login_tools._sync_poll_with_lock(session, "cid", "tenant")

    assert result is expected
    assert store.get("q") == expected.to_json().encode()


# ---------------------------------------------------------------------
# _none helper shape pin
# ---------------------------------------------------------------------


def test_none_helper_returns_canonical_shape() -> None:
    """Pin the shape so accidental field renames break a test."""
    out = _none("any-profile")
    assert set(out.keys()) == {"status", "profile", "signed_in_user_upn", "error"}
    assert out["status"] == "none"
    assert out["profile"] == "any-profile"
    assert out["signed_in_user_upn"] is None
    assert out["error"] is None


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------


async def _stub_poll_loop(
    session: object,
    client_id: str,
    tenant: str,
) -> None:
    """Replacement for the real _poll_loop that just sleeps until cancelled."""
    try:
        await asyncio.sleep(3600)
    except asyncio.CancelledError:
        raise


class _NullStore:
    def get(self, profile: str) -> bytes | None:
        return None

    def set(self, profile: str, value: bytes) -> None:
        pass

    def delete(self, profile: str) -> None:
        pass


class _DictStore:
    def __init__(self) -> None:
        self._d: dict[str, bytes] = {}

    def get(self, profile: str) -> bytes | None:
        return self._d.get(profile)

    def set(self, profile: str, value: bytes) -> None:
        self._d[profile] = value

    def delete(self, profile: str) -> None:
        self._d.pop(profile, None)


