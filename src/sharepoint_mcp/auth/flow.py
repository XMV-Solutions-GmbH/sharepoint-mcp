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
DEFAULT_SCOPES = (
    "Files.ReadWrite.All",
    "Sites.ReadWrite.All",
    "User.Read",
    "offline_access",
)

DEVICE_CODE_GRANT_TYPE = "urn:ietf:params:oauth:grant-type:device_code"

__all__ = [
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
    "poll_for_token",
    "refresh_access_token",
    "request_device_code",
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
