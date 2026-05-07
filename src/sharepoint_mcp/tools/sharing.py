# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Sharing-link tools (closes #47).

- `sp_share_list(url)` — list existing sharing links on an item
- `sp_share_create(url, type, scope, expires=None, password=None)` —
  create a sharing link, returns the share URL
- `sp_share_revoke(url, link_id)` — delete a sharing-link permission

**Security model.** Sharing links create a discoverable access path
on the URL itself: anyone who learns the URL can open the file
(within `scope`'s bounds). They are the most common ISMS-audit
finding ("anyone-with-link sharing enabled on a confidential
document"), so:

- Defaults are conservative: `type='view'`, `scope='organization'`.
  An explicit `scope='anonymous'` is required to create a public link.
- The tool description warns the agent that anonymous + edit is the
  worst combination and should only happen on explicit user request.
- All three tools are gated by `SP_ALLOW_WRITES` at the server layer
  (yes including `sp_share_list` — it doesn't mutate but it's
  thematically grouped with the dangerous siblings, and the agent
  shouldn't be using it without writes-enabled permission anyway).
  Actually no: `sp_share_list` is read-only. It stays in the read
  bucket. Only create + revoke are gated.

Wire shape for sharing-link create response (Microsoft Graph
`POST /drives/{drive-id}/items/{item-id}/createLink`):

    {
        "id": "<permission-id>",
        "roles": ["read"|"write"],
        "link": {
            "type": "view"|"edit"|"embed"|"blocksDownload",
            "scope": "anonymous"|"organization"|"users",
            "webUrl": "<the share URL>",
            "preventsDownload": <bool>
        },
        "expirationDateTime": "<ISO>"|null,
        "hasPassword": <bool>
    }

We normalise to:

    {
        "id": str,                      # permission id (use with revoke)
        "web_url": str,                 # the share URL
        "type": str,                    # view|edit|...
        "scope": str,                   # organization|anonymous|users
        "roles": [str, ...],
        "expiration_date_time": str | None,
        "has_password": bool,
        "prevents_download": bool,
    }
"""

from __future__ import annotations

from typing import Any

import httpx

from sharepoint_mcp.auth import get_token
from sharepoint_mcp.tools._common import (
    GRAPH_BASE,
    parse_sharepoint_url,
    resolve_drive_item,
    resolve_site_id,
)
from sharepoint_mcp.tools.permissions import permissions as _list_all_permissions

VALID_LINK_TYPES = frozenset({"view", "edit", "embed", "blocksDownload"})
VALID_LINK_SCOPES = frozenset({"anonymous", "organization", "users"})


def share_list(
    url: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> list[dict[str, Any]]:
    """List existing sharing links on a SharePoint file or folder.

    This is equivalent to `sp_permissions(url)` filtered to entries
    whose grantee is a sharing link. Returned shape matches what
    `sp_share_create` produces, so callers can correlate.

    Returns an empty list when no sharing links exist on the item.
    """
    if not url or not url.strip():
        raise ValueError("sp_share_list requires a non-empty url")
    _, _, item_path = parse_sharepoint_url(url)
    if not item_path:
        raise ValueError(
            f"sp_share_list requires a file/folder URL; got a site URL: {url!r}",
        )
    # Delegate the underlying permissions fetch + URL resolution; then
    # filter and re-shape.
    all_perms = _list_all_permissions(url, profile=profile, http=http)
    return [_normalise_existing_link(p) for p in all_perms if p["grantee"]["type"] == "link"]


def share_create(
    url: str,
    *,
    type: str = "view",
    scope: str = "organization",
    expires: str | None = None,
    password: str | None = None,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> dict[str, Any]:
    """Create a sharing link on a SharePoint file / folder.

    Defaults are intentionally conservative — `type='view'` and
    `scope='organization'` produce the lowest-risk link. The agent
    must explicitly pass `scope='anonymous'` to create a public link.

    Args:
        url: file or folder URL.
        type: 'view' (default), 'edit', 'embed', or 'blocksDownload'.
        scope: 'organization' (default), 'anonymous', or 'users'.
        expires: optional ISO 8601 datetime — when the link auto-expires.
        password: optional password (only meaningful for anonymous links;
            tenants can disable this).

    Returns the new sharing-link permission with id (for revoke),
    web_url (the share URL), type, scope, roles, etc.

    Raises:
        ValueError: empty url, type/scope outside the allowed set,
            or url points at a site rather than a file/folder.
        httpx.HTTPStatusError: any non-2xx from Graph. Common: 403 if
            the tenant has anonymous-sharing disabled or the user
            lacks the scope.
    """
    if not url or not url.strip():
        raise ValueError("sp_share_create requires a non-empty url")
    if type not in VALID_LINK_TYPES:
        raise ValueError(
            f"sp_share_create type must be one of {sorted(VALID_LINK_TYPES)}; got {type!r}",
        )
    if scope not in VALID_LINK_SCOPES:
        raise ValueError(
            f"sp_share_create scope must be one of {sorted(VALID_LINK_SCOPES)}; got {scope!r}",
        )
    hostname, site_path, item_path = parse_sharepoint_url(url)
    if not item_path:
        raise ValueError(
            f"sp_share_create requires a file/folder URL; got a site URL: {url!r}",
        )

    body: dict[str, Any] = {"type": type, "scope": scope}
    if expires:
        body["expirationDateTime"] = expires
    if password:
        body["password"] = password

    token = get_token(profile)
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        drive_id, item_id = resolve_drive_item(client, site_id, item_path, headers=headers)
        response = client.post(
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/createLink",
            headers=headers,
            json=body,
        )
        response.raise_for_status()
        return _normalise_create_response(response.json())
    finally:
        if http is None:
            client.close()


def share_revoke(
    url: str,
    link_id: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> None:
    """Revoke (delete) a sharing-link permission.

    `link_id` is the permission id returned by `sp_share_create` or
    `sp_share_list`. After this call the share URL stops working.
    """
    if not url or not url.strip():
        raise ValueError("sp_share_revoke requires a non-empty url")
    if not link_id or not str(link_id).strip():
        raise ValueError("sp_share_revoke requires a non-empty link_id")
    hostname, site_path, item_path = parse_sharepoint_url(url)
    if not item_path:
        raise ValueError(
            f"sp_share_revoke requires a file/folder URL; got a site URL: {url!r}",
        )

    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        drive_id, item_id = resolve_drive_item(client, site_id, item_path, headers=headers)
        response = client.delete(
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/permissions/{link_id}",
            headers=headers,
        )
        response.raise_for_status()
    finally:
        if http is None:
            client.close()


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------


def _normalise_create_response(payload: dict[str, Any]) -> dict[str, Any]:
    """Translate a Graph createLink response into our consistent shape."""
    link = payload.get("link") or {}
    if not isinstance(link, dict):
        link = {}
    roles_raw = payload.get("roles") or []
    roles = [str(r) for r in roles_raw] if isinstance(roles_raw, list) else []
    return {
        "id": str(payload.get("id") or ""),
        "web_url": str(link.get("webUrl") or ""),
        "type": str(link.get("type") or ""),
        "scope": str(link.get("scope") or ""),
        "roles": roles,
        "expiration_date_time": (
            str(payload["expirationDateTime"]) if payload.get("expirationDateTime") else None
        ),
        "has_password": bool(payload.get("hasPassword") or False),
        "prevents_download": bool(link.get("preventsDownload") or False),
    }


def _normalise_existing_link(perm: dict[str, Any]) -> dict[str, Any]:
    """Re-shape a sp_permissions entry into the same shape as create.

    sp_share_list calls sp_permissions and filters to link grantees;
    we project here so callers see consistent fields whether they
    just created the link or are listing pre-existing ones. The
    grantee.link_web_url comes from the underlying Graph permission's
    link.webUrl field — see permissions.py.
    """
    grantee_raw = perm.get("grantee")
    grantee = grantee_raw if isinstance(grantee_raw, dict) else {}
    return {
        "id": str(perm.get("id") or ""),
        "web_url": str(grantee.get("link_web_url") or ""),
        "type": str(grantee.get("link_type") or ""),
        "scope": str(grantee.get("link_scope") or ""),
        "roles": list(perm.get("roles") or []),
        "expiration_date_time": None,
        "has_password": False,
        "prevents_download": False,
    }
