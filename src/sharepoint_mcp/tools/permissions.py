# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_permissions — list permissions on a file / folder / site (closes #46).

A single tool that handles three cases by URL shape:

- Site URL (`https://x/sites/foo`) -> `GET /sites/{site-id}/permissions`
- File URL or folder URL inside a drive -> `GET /drives/{drive-id}/items/{item-id}/permissions`

Microsoft Graph's permission shape is heterogeneous (users vs.
groups vs. sharing links vs. legacy SharePoint permissions). We
normalise into one consistent structure:

    {
        "id": "<permission-id>",
        "roles": ["read", "write", ...],
        "grantee": {
            "type": "user" | "group" | "link" | "siteUser" |
                    "siteGroup" | "application" | "unknown",
            "display_name": "<name or empty>",
            "email": "<email or empty>",
            "link_type": "<view|edit|... or empty for non-link grantees>",
            "link_scope": "<anonymous|organization|users or empty>",
        },
        "inherited": <bool>,
    }

For sharing-link permissions, `grantee.type == "link"` and the
link-specific fields populate; the user-facing fields are empty.
For "shared with everyone in the org" links, `link_scope` is
"organization".
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


def permissions(
    url: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> list[dict[str, Any]]:
    """List permissions on a SharePoint file / folder / site.

    `url` is a site URL (lists site permissions) or any item URL
    inside a drive (lists permissions on that item). Folders and
    files are both driveItems and use the same endpoint.

    Returns a list of permission entries — see module docstring for
    the normalised shape.

    Raises:
        ValueError: empty URL.
        httpx.HTTPStatusError: any non-2xx from Graph (404 if not
            visible to the caller; 403 if scope is missing).
    """
    if not url or not url.strip():
        raise ValueError("sp_permissions requires a non-empty url")

    hostname, site_path, item_path = parse_sharepoint_url(url)
    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        if item_path:
            drive_id, item_id = resolve_drive_item(client, site_id, item_path, headers=headers)
            target = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/permissions"
        else:
            target = f"{GRAPH_BASE}/sites/{site_id}/permissions"
        response = client.get(target, headers=headers)
        response.raise_for_status()
        return _extract_permissions(response.json())
    finally:
        if http is None:
            client.close()


def _extract_permissions(payload: dict[str, Any]) -> list[dict[str, Any]]:
    raw = payload.get("value", []) or []
    if not isinstance(raw, list):
        return []
    return [_one_permission(entry) for entry in raw if isinstance(entry, dict)]


def _one_permission(entry: dict[str, Any]) -> dict[str, Any]:
    roles_raw = entry.get("roles", []) or []
    roles = [str(r) for r in roles_raw] if isinstance(roles_raw, list) else []
    return {
        "id": str(entry.get("id") or ""),
        "roles": roles,
        "grantee": _extract_grantee(entry),
        "inherited": "inheritedFrom" in entry,
    }


def _extract_grantee(entry: dict[str, Any]) -> dict[str, Any]:
    """Pick the most specific grantee description Graph populated.

    Graph's shape varies — newer endpoints use grantedToV2 /
    grantedToIdentitiesV2 (for sharing-link grants to multiple
    principals), older ones use grantedTo / grantedToIdentities.
    Sharing-link permissions have a `link` facet instead of a
    grantee identity.
    """
    link = entry.get("link")
    if isinstance(link, dict):
        return {
            "type": "link",
            "display_name": "",
            "email": "",
            "link_type": str(link.get("type") or ""),
            "link_scope": str(link.get("scope") or ""),
        }

    candidates: list[dict[str, Any]] = []
    for key in ("grantedToV2", "grantedTo"):
        v = entry.get(key)
        if isinstance(v, dict):
            candidates.append(v)
    for key in ("grantedToIdentitiesV2", "grantedToIdentities"):
        v = entry.get(key)
        if isinstance(v, list):
            for cand in v:
                if isinstance(cand, dict):
                    candidates.append(cand)

    for cand in candidates:
        normalised = _normalise_identity(cand)
        if normalised["display_name"] or normalised["email"] or normalised["type"] != "unknown":
            return normalised
    return {
        "type": "unknown",
        "display_name": "",
        "email": "",
        "link_type": "",
        "link_scope": "",
    }


def _normalise_identity(identity: dict[str, Any]) -> dict[str, Any]:
    """Translate one identitySet into our consistent grantee shape."""
    for key, type_label in (
        ("user", "user"),
        ("group", "group"),
        ("siteUser", "siteUser"),
        ("siteGroup", "siteGroup"),
        ("application", "application"),
    ):
        principal = identity.get(key)
        if isinstance(principal, dict):
            display_name = str(
                principal.get("displayName")
                or principal.get("loginName")
                or principal.get("LookupValue")
                or ""
            )
            email = str(principal.get("email") or "")
            return {
                "type": type_label,
                "display_name": display_name,
                "email": email,
                "link_type": "",
                "link_scope": "",
            }
    return {
        "type": "unknown",
        "display_name": "",
        "email": "",
        "link_type": "",
        "link_scope": "",
    }
