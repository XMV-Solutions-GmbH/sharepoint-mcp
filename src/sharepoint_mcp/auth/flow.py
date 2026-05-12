# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""OAuth 2.0 Device Code Flow + refresh-token client against Microsoft Identity.

Thin shim over `mcp-microsoft-graph-auth`'s `device_code` module
that supplies SharePoint-specific defaults: the bundled multi-tenant
Entra app's client_id, the SharePoint-flavoured Graph scopes, and
the multi-tenant `organizations` authority.

Each function delegates to the shared library after applying the
defaults. Existing imports (`sharepoint_mcp.auth.flow.poll_for_token`)
keep working without source-level changes.
"""

from __future__ import annotations

import os
import time
from collections.abc import Callable

import httpx
from mcp_microsoft_graph_auth.device_code import (
    AUTHORITY_BASE,
    AuthorizationDeniedError,
    DeviceCodeChallenge,
    DeviceCodeError,
    DeviceCodeExpiredError,
    RefreshTokenInvalidError,
)
from mcp_microsoft_graph_auth.device_code import (
    poll_for_token as _lib_poll_for_token,
)
from mcp_microsoft_graph_auth.device_code import (
    refresh_access_token as _lib_refresh_access_token,
)
from mcp_microsoft_graph_auth.device_code import (
    request_device_code as _lib_request_device_code,
)
from mcp_microsoft_graph_auth.tokens import CachedToken

# ---------------------------------------------------------------------
# SharePoint-specific defaults — see app-concept § Authentication.
# The client_id is the XMV-published multi-tenant Entra app
# registration; users override via SP_CLIENT_ID for BYO scenarios.
# ---------------------------------------------------------------------

DEFAULT_CLIENT_ID = "cb7cf68d-90d5-4841-90a7-de3a40be280b"
DEFAULT_AUTHORITY_TENANT = "organizations"

# v0.5 — split scopes so the consent screen reflects the operator's
# actual decision. With SP_ALLOW_WRITES=false the consent prompt
# reads "this app can read your SharePoint files" only; with =true
# it shows the ReadWrite variants too. Microsoft's scope model lets
# us request narrower scopes when we don't need writes.
_BASE_SCOPES: tuple[str, ...] = (
    "Files.Read.All",
    "Sites.Read.All",
    "User.Read",
    "offline_access",
)
_WRITES_REPLACEMENTS = {
    "Files.Read.All": "Files.ReadWrite.All",
    "Sites.Read.All": "Sites.ReadWrite.All",
}

# Env var that opts the running MCP server into requesting the
# ReadWrite scopes at OAuth time (and registering the write tools as
# MCP tools). v0.5 made this strict — must be exactly `"true"` or
# `"false"`; unset / empty / legacy truthy (`1`/`yes`/`on`) all raise
# `SharepointConsentNotConfiguredError` at startup. Rationale:
# operators silently landing in read-only mode without realising
# writes are a separately-opt-in feature was the dominant onboarding
# failure mode in v0.4.x.
ALLOW_WRITES_ENV = "SP_ALLOW_WRITES"
_STRICT_TRUE = "true"
_STRICT_FALSE = "false"


class SharepointConsentNotConfiguredError(RuntimeError):
    """Raised at server-build / CLI-login time when `SP_ALLOW_WRITES`
    is unset or has a non-`true`/`false` value.

    The exception message is the user-facing onboarding hint —
    callers re-raise without wrapping so the operator sees it
    verbatim on stderr.
    """


def _strict_bool_env(name: str) -> bool:
    """Read `name` from the environment and parse strictly.

    Returns `True` for "true", `False` for "false" (case-insensitive,
    leading/trailing whitespace ignored). Raises
    `SharepointConsentNotConfiguredError` with the documented
    onboarding-help message for anything else, including unset / empty.
    """
    raw = os.environ.get(name)
    if raw is not None:
        normalised = raw.strip().lower()
        if normalised == _STRICT_TRUE:
            return True
        if normalised == _STRICT_FALSE:
            return False
    raise SharepointConsentNotConfiguredError(_consent_help_text(name, raw))


def _consent_help_text(name: str, raw: str | None) -> str:
    """Format the onboarding-help message for an unset / invalid
    consent env var."""
    got = "(not set)" if raw is None else f"{raw!r}"
    return (
        f"ERROR: mcp-server-sharepoint requires an explicit "
        f"{ALLOW_WRITES_ENV} decision (got {got}).\n\n"
        f"This server can check out, edit, and check in SharePoint files "
        f"on the signed-in user's behalf (opt-in) or operate in read-only "
        f"mode. There is no implicit default — the operator must "
        f"consciously decide.\n\n"
        f"Set in your MCP client config (.mcp.json env section):\n\n"
        f'  "{ALLOW_WRITES_ENV}": "true"    — enable checkout / save / '
        f"share / list-item / page-edit tools\n"
        f'  "{ALLOW_WRITES_ENV}": "false"   — read-only (no write tools)\n\n'
        f'With "false", the OAuth consent screen requests only '
        f'`Files.Read.All` + `Sites.Read.All`. With "true", it requests '
        f"the `.ReadWrite.All` variants. The decision flows through to "
        f"both the tool surface AND the consent prompt.\n\n"
        f"See README §Authentication for the design rationale."
    )


def validate_consent_config() -> bool:
    """Validate the consent env var at startup.

    Returns `writes_enabled` (True/False). Raises
    `SharepointConsentNotConfiguredError` with a clear, actionable
    error message if `SP_ALLOW_WRITES` is unset or has a
    non-`true`/`false` value.
    """
    return _strict_bool_env(ALLOW_WRITES_ENV)


def resolve_scopes() -> tuple[str, ...]:
    """Return the OAuth scopes to request at this moment.

    With `SP_ALLOW_WRITES=false`: returns the read-only base
    (`Files.Read.All`, `Sites.Read.All`, plus the always-needed
    `User.Read` and `offline_access`). With `=true`: swaps in the
    `ReadWrite.All` variants. Resolved at call time, not at module
    load, so test-time `monkeypatch.setenv` flips behaviour without
    re-importing.

    Raises `SharepointConsentNotConfiguredError` if the env var
    is not configured strictly.
    """
    if validate_consent_config():
        return tuple(_WRITES_REPLACEMENTS.get(s, s) for s in _BASE_SCOPES)
    return _BASE_SCOPES


# Backwards-compat alias for callers that imported the old constant
# at module load. They get the writes-enabled scope set so existing
# scripts that hard-coded `DEFAULT_SCOPES` for write flows still
# work; the strict-validation step in `_build_server` and `cli.main`
# catches operators who never set the env var.
DEFAULT_SCOPES: tuple[str, ...] = tuple(_WRITES_REPLACEMENTS.get(s, s) for s in _BASE_SCOPES)

DEVICE_CODE_GRANT_TYPE = "urn:ietf:params:oauth:grant-type:device_code"

__all__ = [
    "ALLOW_WRITES_ENV",
    "AUTHORITY_BASE",
    "DEFAULT_AUTHORITY_TENANT",
    "DEFAULT_CLIENT_ID",
    "DEFAULT_SCOPES",
    "DEVICE_CODE_GRANT_TYPE",
    "AuthorizationDeniedError",
    "CachedToken",
    "DeviceCodeChallenge",
    "DeviceCodeError",
    "DeviceCodeExpiredError",
    "RefreshTokenInvalidError",
    "SharepointConsentNotConfiguredError",
    "poll_for_token",
    "refresh_access_token",
    "request_device_code",
    "resolve_scopes",
    "validate_consent_config",
]


def request_device_code(
    *,
    client_id: str = DEFAULT_CLIENT_ID,
    tenant: str = DEFAULT_AUTHORITY_TENANT,
    scopes: tuple[str, ...] = DEFAULT_SCOPES,
    http: httpx.Client | None = None,
) -> tuple[str, DeviceCodeChallenge]:
    """Initiate the Device Code flow with SharePoint-flavoured defaults."""
    return _lib_request_device_code(
        client_id=client_id,
        tenant=tenant,
        scopes=scopes,
        http=http,
    )


def poll_for_token(
    *,
    device_code: str,
    client_id: str = DEFAULT_CLIENT_ID,
    tenant: str = DEFAULT_AUTHORITY_TENANT,
    interval: int = 5,
    http: httpx.Client | None = None,
    sleep: Callable[[float], None] = time.sleep,
    now: Callable[[], float] = time.time,
) -> CachedToken:
    """Poll `/token` until the user completes (or denies) sign-in."""
    return _lib_poll_for_token(
        device_code=device_code,
        client_id=client_id,
        tenant=tenant,
        interval=interval,
        http=http,
        sleep=sleep,
        now=now,
    )


def refresh_access_token(
    *,
    refresh_token: str,
    client_id: str = DEFAULT_CLIENT_ID,
    tenant: str = DEFAULT_AUTHORITY_TENANT,
    scopes: tuple[str, ...] = DEFAULT_SCOPES,
    http: httpx.Client | None = None,
) -> CachedToken:
    """Exchange a refresh token for a new access (and refresh) token."""
    return _lib_refresh_access_token(
        refresh_token=refresh_token,
        client_id=client_id,
        tenant=tenant,
        scopes=scopes,
        http=http,
    )
