# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Authentication public API.

Two entry points:

- `get_token(profile)` — silent path. Returns a fresh access token
  from the cached refresh token. Refreshes through Microsoft Identity
  if needed. Never blocks on user interaction. Raises
  `AuthRequiredError` if the cache is empty or the refresh token has
  been invalidated; the caller is expected to surface that to the
  human and arrange `interactive_login()` to be invoked separately.
- `interactive_login(profile)` — out-of-band path. Drives the full
  Device Code flow, blocks until the human completes (or refuses) the
  prompt, persists the resulting tokens to the configured TokenStore.
  Intended to be invoked from a CLI subcommand (`sharepoint-mcp
  login`), not from inside an MCP tool call.

This split mirrors how `gh auth login` separates from `gh` runtime
calls: the MCP server does not pause to do interactive auth in the
middle of a tool call.

`SP_CLIENT_ID` and `SP_TENANT_ID` env vars override the bundled
multi-tenant defaults — see `docs/app-concept.md` § Authentication.
"""

from __future__ import annotations

import os
import sys
import webbrowser
from collections.abc import Callable
from urllib.parse import urlencode, urlsplit, urlunsplit

import httpx

from sharepoint_mcp.auth.flow import (
    DEFAULT_AUTHORITY_TENANT,
    DEFAULT_CLIENT_ID,
    AuthorizationDeniedError,
    DeviceCodeChallenge,
    DeviceCodeError,
    DeviceCodeExpiredError,
    RefreshTokenInvalidError,
    poll_for_token,
    refresh_access_token,
    request_device_code,
)
from sharepoint_mcp.auth.store import TokenStore, get_token_store
from sharepoint_mcp.auth.tokens import CachedToken

CLIENT_ID_ENV = "SP_CLIENT_ID"
TENANT_ENV = "SP_TENANT_ID"

__all__ = [
    "AuthRequiredError",
    "AuthorizationDeniedError",
    "CachedToken",
    "DeviceCodeChallenge",
    "DeviceCodeError",
    "DeviceCodeExpiredError",
    "RefreshTokenInvalidError",
    "get_token",
    "interactive_login",
]


class AuthRequiredError(RuntimeError):
    """No usable cached token; the caller must trigger `interactive_login`.

    Raised by `get_token` when (a) the cache is empty for this profile,
    (b) the cached access token expired and there is no refresh token
    to use, or (c) Microsoft Identity rejected the refresh token (the
    expired-after-too-long-idle case).

    The MCP tool layer should catch this, surface a clear message to
    the agent (and through it to the user), and stop. Re-authentication
    happens out of band via `uvx sharepoint-mcp login`.
    """

    def __init__(self, profile: str, reason: str) -> None:
        super().__init__(
            f"No usable credentials for profile {profile!r}: {reason}. "
            f"Run `uvx sharepoint-mcp login --profile {profile}` to sign in.",
        )
        self.profile = profile
        self.reason = reason


def _resolve_client_id(client_id: str | None) -> str:
    if client_id:
        return client_id
    env = os.environ.get(CLIENT_ID_ENV, "").strip()
    return env or DEFAULT_CLIENT_ID


def _resolve_tenant(tenant: str | None) -> str:
    if tenant:
        return tenant
    env = os.environ.get(TENANT_ENV, "").strip()
    return env or DEFAULT_AUTHORITY_TENANT


def _verification_uri_with_code(uri: str, code: str) -> str:
    """Append `?otc=<code>` so Microsoft's device-login page pre-fills the code.

    Microsoft's device-login page accepts an `otc` (one-time code) query
    parameter and auto-fills the input field with it, so the user only
    has to click and pick an account — no manual code typing.
    """
    parts = urlsplit(uri)
    existing = parts.query
    new_query = (existing + "&" if existing else "") + urlencode({"otc": code})
    return urlunsplit((parts.scheme, parts.netloc, parts.path, new_query, parts.fragment))


def _has_desktop_session() -> bool:
    """Heuristic: true iff there's likely a usable graphical browser."""
    if sys.platform == "darwin" or sys.platform.startswith("win"):
        return True
    return bool(os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY"))


def _default_prompt(challenge: DeviceCodeChallenge) -> None:
    """Show the Device Code challenge to the human running login.

    On a desktop session (macOS/Windows always, Linux with DISPLAY or
    WAYLAND_DISPLAY), tries to open the pre-filled verification URL in
    the default browser. In every case, also prints the URL + code to
    stderr so headless users (and our background-task captures) get
    them in the log.
    """
    pre_filled = _verification_uri_with_code(challenge.verification_uri, challenge.user_code)

    opened = False
    if _has_desktop_session():
        try:
            opened = webbrowser.open(pre_filled, new=2)
        except webbrowser.Error:
            opened = False

    if opened:
        msg = (
            "\nOpening your browser to complete sign-in.\n"
            f"  If it didn't open, visit: {pre_filled}\n"
            f"  (Code is pre-filled: {challenge.user_code})\n"
            "\nWaiting for sign-in..."
        )
    else:
        msg = (
            "\nTo sign in to sharepoint-mcp:\n"
            f"  1. Open: {pre_filled}\n"
            f"     (the code {challenge.user_code} is already in the URL)\n"
            "  2. Pick the M365 account you want to authorise.\n"
            "\nWaiting for sign-in..."
        )

    print(msg, file=sys.stderr, flush=True)


def get_token(
    profile: str = "default",
    *,
    client_id: str | None = None,
    tenant: str | None = None,
    store: TokenStore | None = None,
    http: httpx.Client | None = None,
) -> str:
    """Return a valid access token for `profile`.

    Reads from the configured TokenStore. If the cached access token
    is still valid (with a small refresh buffer), returns it directly.
    If expired and a refresh token is present, exchanges it through
    Microsoft Identity silently and persists the new tokens.

    Raises:
        AuthRequiredError: no cached entry, no refresh token, or the
            refresh token was rejected. The caller must trigger
            `interactive_login` to recover.
    """
    resolved_client = _resolve_client_id(client_id)
    resolved_tenant = _resolve_tenant(tenant)
    resolved_store = store if store is not None else get_token_store()

    raw = resolved_store.get(profile)
    if raw is None:
        raise AuthRequiredError(profile, "no cached credentials found")

    cached = CachedToken.from_json(raw.decode())
    if not cached.is_expired():
        return cached.access_token

    if cached.refresh_token is None:
        raise AuthRequiredError(
            profile, "cached token has expired and no refresh token is available"
        )

    try:
        new_token = refresh_access_token(
            refresh_token=cached.refresh_token,
            client_id=resolved_client,
            tenant=resolved_tenant,
            http=http,
        )
    except RefreshTokenInvalidError as exc:
        resolved_store.delete(profile)
        raise AuthRequiredError(
            profile, f"refresh token rejected by Microsoft Identity ({exc})"
        ) from exc

    resolved_store.set(profile, new_token.to_json().encode())
    return new_token.access_token


def interactive_login(
    profile: str = "default",
    *,
    client_id: str | None = None,
    tenant: str | None = None,
    store: TokenStore | None = None,
    prompt: Callable[[DeviceCodeChallenge], None] | None = None,
    http: httpx.Client | None = None,
) -> CachedToken:
    """Run the Device Code flow end-to-end. Blocks until completion.

    On success: persists the issued tokens to `store` (or the
    auto-detected store) under `profile`, and returns the
    `CachedToken`.

    Raises:
        AuthorizationDeniedError: user refused the prompt.
        DeviceCodeExpiredError: device code expired before sign-in.

    Intended to be invoked from a CLI subcommand or test. Do not call
    from inside an MCP tool handler — it blocks the server for up to
    ~15 minutes.
    """
    resolved_client = _resolve_client_id(client_id)
    resolved_tenant = _resolve_tenant(tenant)
    resolved_store = store if store is not None else get_token_store()
    resolved_prompt = prompt if prompt is not None else _default_prompt

    device_code, challenge = request_device_code(
        client_id=resolved_client,
        tenant=resolved_tenant,
        http=http,
    )
    resolved_prompt(challenge)

    cached = poll_for_token(
        device_code=device_code,
        client_id=resolved_client,
        tenant=resolved_tenant,
        interval=challenge.interval,
        http=http,
    )
    resolved_store.set(profile, cached.to_json().encode())
    return cached
