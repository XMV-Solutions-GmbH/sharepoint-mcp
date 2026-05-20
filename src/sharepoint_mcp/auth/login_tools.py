# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Integrated MCP-tool login flow (sp_auth_begin + sp_auth_status).

Provides the two MCP tools the agent calls to drive Device Code login
without shelling out to the CLI:

- `login_begin(profile, force=False)` — non-blocking. Returns within
  ~1 second with `user_code` + `verification_url`. An asyncio task
  continues polling Microsoft Identity in the background and writes
  the resulting token to the configured TokenStore on success.
- `login_status(profile)` — three states the agent can act on
  directly: `signed_in` (valid token on disk, regardless of how it
  got there — CLI or tool-flow), `pending` (Device Code flow in
  progress), `none` (neither — agent should call `login_begin`).
  Critically, `signed_in` is determined by an active probe of the
  token cache, not just by checking the in-memory session — so a
  user who logged in via CLI days ago shows up correctly.

The session registry is the lib's `LoginSessionRegistry` (process-
local, thread-safe). The asyncio task lifecycle and the JWT-claim
extraction for `signed_in_user_upn` live here because they depend
on the consumer's TokenStore + CLI-resolution conventions, which the
lib stays agnostic of.

**Limitation: pending sessions live in this process only.** If the
MCP server restarts mid-flow (Claude Code session ends, server
crashes, container redeployed), the session is lost and the agent
must call `login_begin` again. Persisting them is non-trivial (the
asyncio task can't be serialised; the new task on restart would
have to resume polling against the original device_code, which
typically has expired anyway in ~15 min) and is deferred.
"""

from __future__ import annotations

import asyncio
import base64
import binascii
import json
from datetime import UTC, datetime, timezone
from pathlib import Path
from typing import Any

import httpx

# Symbols come from the shared library; we keep the import surface
# narrow to make it obvious which bits this module orchestrates.
from mcp_microsoft_graph_auth import (
    AuthorizationDeniedError,
    DeviceCodeExpiredError,
    LoginSession,
    LoginSessionRegistry,
    TokenStoreLockTimeoutError,
)
from mcp_microsoft_graph_auth._filelock import exclusive_lock
from mcp_microsoft_graph_auth.tokens import CachedToken

from sharepoint_mcp.auth import (
    _resolve_client_id,
    _resolve_tenant,
)
from sharepoint_mcp.auth.flow import (
    DEFAULT_SCOPES,
    poll_for_token,
    refresh_access_token,
    request_device_code,
)
from sharepoint_mcp.auth.service_principal import is_service_principal_mode
from sharepoint_mcp.auth.store import DEFAULT_CACHE_DIR, get_token_store

# Module-level registry instance. Process-scoped: shared across all
# tool calls within a single MCP server lifetime. Test isolation
# happens by either using `_REGISTRY.clear()` or, in tests that need
# a clean dict, by patching the module-level binding.
_REGISTRY = LoginSessionRegistry()


class ConcurrentLoginAttemptError(RuntimeError):
    """Another process already holds the login-flow lock for this profile.

    Raised when `sp_auth_begin` detects a cross-process Device Code
    flow in progress (CLI or another MCP server instance). The agent
    should surface this to the user and ask them to wait before retrying.
    """


class DeviceCodeRequestFailedError(RuntimeError):
    """Microsoft Identity refused the /devicecode request.

    Common causes: network failure, BYO `SP_CLIENT_ID` misconfigured
    (app not registered as public client / Device Code flow not
    enabled), tenant guards. The wrapped HTTP error is the cause.
    """


class ServicePrincipalActiveError(RuntimeError):
    """The server is configured for service-principal auth.

    `sp_auth_begin` is only meaningful in delegated mode. In
    service-principal mode the consumer's `get_app_only_token()`
    auto-acquires tokens from the configured `SP_CLIENT_SECRET`;
    no Device Code flow is involved.
    """


# ---------------------------------------------------------------------
# login_begin
# ---------------------------------------------------------------------


async def login_begin(
    *,
    profile: str = "default",
    force: bool = False,
) -> dict[str, Any]:
    """Initiate a Device Code login for `profile`. Non-blocking.

    Returns within ~1 second with the user-facing fields
    (`user_code`, `verification_url`, `expires_at`, etc.). An
    asyncio task continues to poll Microsoft Identity in the
    background and writes the token + sets `status="success"`
    when the user completes sign-in.

    Idempotency:
    - If a `pending` session already exists and `force=False`, the
      existing session is returned unchanged (same `user_code`).
    - If `force=True`, the existing task is cancelled and a fresh
      session is created.
    - If a terminal session (`success` / `expired` / etc.) exists,
      a fresh session is started regardless of `force`.

    Raises:
        ServicePrincipalActiveError: server is in service-principal
            mode (no Device Code flow needed).
        DeviceCodeRequestFailedError: Microsoft Identity refused
            the /devicecode request. Wrapped HTTP error in `__cause__`.
    """
    if is_service_principal_mode():
        raise ServicePrincipalActiveError(
            "sp_auth_begin is for delegated user auth. The server is "
            "configured for service-principal mode (SP_AUTH_MODE / "
            "SP_CLIENT_SECRET) — tokens are auto-acquired from the "
            "client_secret without user interaction.",
        )

    existing = _REGISTRY.get(profile)
    if existing is not None and existing.status == "pending":
        if not force:
            return _public_view(existing)
        # force=True: cancel the existing poller and fall through to
        # start a fresh flow.
        if existing.task is not None and not existing.task.done():
            existing.task.cancel()
        _REGISTRY.remove(profile)

    # Existing session in a terminal state (or none): start fresh.
    client_id = _resolve_client_id(None)
    tenant = _resolve_tenant(None)

    # Cross-process guard: fail fast if the CLI or another MCP server
    # instance already holds the login-flow lock for this profile.
    # The full-duration lock is held inside _poll_loop; this probe
    # prevents an unnecessary /devicecode round-trip and gives the
    # agent an immediate, actionable error instead of a silent race.
    _flow_lock_path = _login_lock_path(profile)
    try:
        await asyncio.to_thread(_probe_login_lock, _flow_lock_path)
    except TokenStoreLockTimeoutError:
        raise ConcurrentLoginAttemptError(
            f"Another process is already running a Device Code login for "
            f"profile {profile!r}. Wait for it to complete and then retry.",
        ) from None

    try:
        device_code, challenge = await asyncio.to_thread(
            request_device_code,
            client_id=client_id,
            tenant=tenant,
            scopes=DEFAULT_SCOPES,
        )
    except httpx.HTTPError as exc:
        raise DeviceCodeRequestFailedError(
            f"Microsoft Identity refused the device-code request: {exc}",
        ) from exc

    session = LoginSession.new(
        profile=profile,
        device_code=device_code,
        user_code=challenge.user_code,
        verification_url=challenge.verification_uri,
        verification_url_complete=challenge.verification_uri_complete,
        expires_at=datetime.fromtimestamp(challenge.expires_at, tz=UTC),
        interval_s=challenge.interval,
    )
    # First-write-wins for two concurrent callers — the lib's
    # put_if_absent guarantees that.
    final = _REGISTRY.put_if_absent(session)
    if final is not session:
        # Another caller raced and won: return their session, drop
        # the device_code we just requested unused.
        return _public_view(final)

    # Spawn the polling task. Stored on the session for later
    # cancellation via `force=True` or shutdown.
    final.task = asyncio.create_task(_poll_loop(final, client_id, tenant))
    return _public_view(final)


async def _poll_loop(
    session: LoginSession,
    client_id: str,
    tenant: str,
) -> None:
    """Drive the Device Code poll until success / expiry / cancel.

    Runs the blocking poll inside a thread via `asyncio.to_thread`.
    The login-flow file lock is held for the entire poll duration so
    that a concurrent CLI login or a second MCP server process can
    detect the in-progress flow and fail fast with a clear error
    instead of starting a competing Device Code flow.

    Updates `session.status` in place.
    """
    try:
        cached = await asyncio.to_thread(
            _sync_poll_with_lock,
            session,
            client_id,
            tenant,
        )
        session.signed_in_user_upn = _extract_upn_from_jwt(cached.access_token)
        session.status = "success"
    except asyncio.CancelledError:
        session.status = "cancelled"
        raise
    except TokenStoreLockTimeoutError as exc:
        # Another process grabbed the lock between the probe in
        # login_begin and the task start — rare but possible.
        session.status = "failed"
        session.error = {"code": "concurrent_login_attempt", "message": str(exc)}
    except DeviceCodeExpiredError as exc:
        session.status = "expired"
        session.error = {"code": "expired", "message": str(exc)}
    except AuthorizationDeniedError as exc:
        session.status = "failed"
        session.error = {"code": "access_denied", "message": str(exc)}
    except (httpx.HTTPError, RuntimeError) as exc:
        session.status = "failed"
        session.error = {"code": "unknown", "message": str(exc)}


def _sync_poll_with_lock(
    session: LoginSession,
    client_id: str,
    tenant: str,
) -> CachedToken:
    """Blocking poll under the login-flow file lock.

    Acquires the cross-process login-flow lock for this profile for
    the entire poll duration (~up to 15 min). The lock prevents two
    processes from driving simultaneous Device Code flows for the
    same profile, which would confuse the user with two prompts.

    Raises `TokenStoreLockTimeoutError` if another process already
    holds the lock (the probe in `login_begin` already caught this in
    the common case; this handles the narrow race window between probe
    and task start).
    """
    lock_path = _login_lock_path(session.profile)
    # Timeout is generous: the Device Code flow itself expires in
    # ~15 minutes. We want to hold the lock until the poll resolves,
    # not time out while the user is completing sign-in.
    with exclusive_lock(lock_path, timeout=1200.0):
        cached = poll_for_token(
            device_code=session.device_code,
            client_id=client_id,
            tenant=tenant,
            interval=session.interval_s,
        )
        store = get_token_store()
        store.set(session.profile, cached.to_json().encode())
    return cached


# ---------------------------------------------------------------------
# File-lock helpers
# ---------------------------------------------------------------------


def _login_lock_path(profile: str) -> Path:
    """Sidecar lock file for cross-process login-flow serialisation.

    Always file-system based (even when the token store uses keyring)
    so the lock path is deterministic and OS-level, not process-local.
    The file is created on demand; its contents are irrelevant.
    """
    path = DEFAULT_CACHE_DIR / profile / "login.lock"
    path.parent.mkdir(parents=True, exist_ok=True)
    return path


def _probe_login_lock(path: Path) -> None:
    """Non-blocking lock probe — raises `TokenStoreLockTimeoutError` if
    another process currently holds the login-flow lock.

    Used by `login_begin` to fail fast before making the /devicecode
    round-trip. Acquires and immediately releases the lock (no side
    effects on the winning caller).
    """
    with exclusive_lock(path, timeout=0.1):
        pass


# ---------------------------------------------------------------------
# login_status
# ---------------------------------------------------------------------


async def login_status(*, profile: str = "default") -> dict[str, Any]:
    """Return the active state for `profile`.

    Three resolved states:

    - `signed_in` — valid token on disk OR in-memory session
      reached `success`. `signed_in_user_upn` populated.
    - `pending` — Device Code flow in progress. `user_code` /
      `verification_url` / `time_remaining_s` populated.
    - `none` — neither: agent should call `sp_auth_begin`.

    Recently-terminal sessions (`expired` / `failed` / `cancelled`)
    surface their error once via the `error` field but do NOT advance
    to `none` — they keep their terminal state until `sp_auth_begin`
    is called again. This lets the agent surface a clear "code
    expired, please try again" message instead of "you're not signed
    in" (which would be ambiguous between never-tried and just-failed).
    """
    session = _REGISTRY.get(profile)
    if session is not None:
        if session.status == "pending":
            return {
                "status": "pending",
                "profile": profile,
                "session_id": session.session_id,
                "user_code": session.user_code,
                "verification_url": session.verification_url,
                "verification_url_complete": session.verification_url_complete,
                "time_remaining_s": session.time_remaining_s(),
                "signed_in_user_upn": None,
                "error": None,
            }
        if session.status == "success":
            return {
                "status": "signed_in",
                "profile": profile,
                "signed_in_user_upn": session.signed_in_user_upn,
                "error": None,
            }
        # Terminal failure — surface error, distinct from `none` so
        # the agent can render a clear "what went wrong" message.
        return {
            "status": session.status,  # "expired" | "failed" | "cancelled"
            "profile": profile,
            "signed_in_user_upn": None,
            "error": session.error,
        }

    # No in-memory session — actively probe the token cache.
    try:
        store = get_token_store()
    except Exception:
        # Token-store init itself failed (e.g. encrypted-file with no
        # passphrase). Surface as `none` so the agent prompts a fresh
        # login; the underlying error will surface again at login_begin
        # time with better context.
        return _none(profile)

    try:
        cached_raw = store.get(profile)
    except Exception:
        return _none(profile)
    if cached_raw is None:
        return _none(profile)

    try:
        cached = CachedToken.from_json(cached_raw.decode())
    except Exception:
        # Corrupt cache entry. Treat as `none` and let login_begin sort
        # it out — store.set on success will overwrite the corruption.
        return _none(profile)

    if not cached.is_expired():
        return {
            "status": "signed_in",
            "profile": profile,
            "signed_in_user_upn": _extract_upn_from_jwt(cached.access_token),
            "error": None,
        }

    # Access expired; try to refresh silently.
    if not cached.refresh_token:
        return _none(profile)
    try:
        new = await asyncio.to_thread(
            refresh_access_token,
            refresh_token=cached.refresh_token,
            client_id=_resolve_client_id(None),
            tenant=_resolve_tenant(None),
            scopes=DEFAULT_SCOPES,
        )
    except Exception:
        return _none(profile)
    try:
        store.set(profile, new.to_json().encode())
    except Exception:
        # Could not persist the refreshed token. The token is still
        # valid in memory but won't survive process restart. Surface
        # as signed_in with a best-effort upn — better than `none`.
        return {
            "status": "signed_in",
            "profile": profile,
            "signed_in_user_upn": _extract_upn_from_jwt(new.access_token),
            "error": None,
        }
    return {
        "status": "signed_in",
        "profile": profile,
        "signed_in_user_upn": _extract_upn_from_jwt(new.access_token),
        "error": None,
    }


def _none(profile: str) -> dict[str, Any]:
    """Standard 'no auth state for this profile' response."""
    return {
        "status": "none",
        "profile": profile,
        "signed_in_user_upn": None,
        "error": None,
    }


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------


def _public_view(session: LoginSession) -> dict[str, Any]:
    """Tool-output-friendly dict — drops the secret device_code and the
    opaque task handle. Mirrors the lib's `public_view` shape but with
    sharepoint-mcp-flavoured field selection."""
    return {
        "status": session.status,
        "profile": session.profile,
        "session_id": session.session_id,
        "user_code": session.user_code,
        "verification_url": session.verification_url,
        "verification_url_complete": session.verification_url_complete,
        "time_remaining_s": session.time_remaining_s(),
        "expires_at": session.expires_at.isoformat(),
        "interval_s": session.interval_s,
        "signed_in_user_upn": session.signed_in_user_upn,
        "error": session.error,
    }


def _extract_upn_from_jwt(access_token: str) -> str | None:
    """Pull `upn` (or `preferred_username`) from a Microsoft Identity JWT.

    Microsoft Identity v2.0 access tokens are JWTs whose middle segment
    is the base64url-encoded claims payload. This extraction is purely
    local (no Graph round-trip) and cheap — typical agent UX wants the
    upn rendered in `sp_auth_status`'s response immediately.

    Returns None on any parse failure. The consumer can fall back to a
    `/me` lookup if precise display-name matters more than latency.
    """
    try:
        parts = access_token.split(".")
        if len(parts) != 3:
            return None
        payload = parts[1]
        # Restore base64url padding.
        payload += "=" * (-len(payload) % 4)
        decoded = base64.urlsafe_b64decode(payload)
        claims = json.loads(decoded)
    except (binascii.Error, ValueError, json.JSONDecodeError):
        return None
    upn = claims.get("upn") or claims.get("preferred_username") or ""
    return str(upn) or None


# Suppress unused-import warning for the timezone alias — kept for
# downstream code that imports it from this module.
_ = timezone
