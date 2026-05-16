# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Site discovery (closes #49).

Two read-only tools:

- `sp_sites(query=None)` — search across sites the user can see.
  Wraps `GET /sites?search=...`. Empty query lists everything visible
  via the multi-tenant default (typically sites under the user's
  primary tenant).
- `sp_followed_sites()` — the "my SharePoint" entrypoint, wrapping
  `GET /me/followedSites`. Useful for an agent that wants to start
  from the user's curated list rather than guess at site URLs.

Returned dict shape per site (consistent across both tools):

    {
        "id": "<graph-site-id>",
        "name": "<displayName>",
        "web_url": "<webUrl>",
        "description": "<description or empty>",
        "last_modified": "<ISO datetime or empty>",
    }

The `id` is the Graph composite ID (`hostname,siteCollectionId,webId`)
that other Graph endpoints accept. Callers don't usually need it —
the `web_url` is what they pass to `sp_list_folder` / `sp_search` etc. —
but it's exposed for advanced use.
"""

from __future__ import annotations

from typing import Any

import httpx

from sharepoint_mcp.auth import get_token
from sharepoint_mcp.tools._common import (
    GRAPH_BASE,
    list_site_drives,
    parse_sharepoint_url,
    resolve_site_id,
)

__all__ = ["drives", "followed_sites", "sites"]


def sites(
    query: str | None = None,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> list[dict[str, Any]]:
    """Search SharePoint sites visible to the signed-in user.

    `query` is a free-text site-name search. `None` or empty string
    lists all visible sites (Microsoft's documented behaviour for
    `?search=*`).

    Returns up to ~25 results — Microsoft caps `/sites?search` and
    doesn't currently support `$top` / `$skip` for that endpoint.
    """
    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        # Empty / None: search="*" returns all sites the user can see.
        # Microsoft requires the search parameter to be present.
        search_value = (query or "").strip() or "*"
        response = client.get(
            f"{GRAPH_BASE}/sites",
            headers=headers,
            params={"search": search_value},
        )
        response.raise_for_status()
        return _extract_sites(response.json())
    finally:
        if http is None:
            client.close()


def drives(
    site_url: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> list[dict[str, Any]]:
    """List the document libraries (drives) on a SharePoint site.

    `site_url` is the human-readable site URL
    (`https://contoso.sharepoint.com/sites/foo`). Returns each drive
    with id, name, web_url, description, drive_type ("documentLibrary"
    for the typical case), and quota info when present.

    Together with `sp_sites`, this lets the agent discover which
    library to read from on a given site without having to know the
    library names upfront.
    """
    if not site_url or not site_url.strip():
        raise ValueError("sp_drives requires a non-empty site_url")
    hostname, site_path, item_path = parse_sharepoint_url(site_url)
    if item_path:
        raise ValueError(
            f"sp_drives expects a site URL, not a file/folder URL "
            f"(got {site_url!r}; item path {item_path!r}).",
        )

    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        raw = list_site_drives(client, site_id, headers=headers)
    finally:
        if http is None:
            client.close()
    return [_one_drive(d) for d in raw]


def _one_drive(entry: dict[str, Any]) -> dict[str, Any]:
    quota = entry.get("quota") or {}
    return {
        "id": str(entry.get("id") or ""),
        "name": str(entry.get("name") or ""),
        "web_url": str(entry.get("webUrl") or ""),
        "description": str(entry.get("description") or ""),
        "drive_type": str(entry.get("driveType") or ""),
        "quota_total": int(quota.get("total") or 0) if isinstance(quota, dict) else 0,
        "quota_used": int(quota.get("used") or 0) if isinstance(quota, dict) else 0,
    }


def followed_sites(
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> list[dict[str, Any]]:
    """List sites the signed-in user has followed in SharePoint.

    Wraps `GET /me/followedSites`. The list reflects the user's own
    "Following" list in the SharePoint web UI. In service-principal
    mode there's no `/me`, so this raises a clear error rather than
    silently returning empty — the agent should fall back to
    `sp_sites()` for a tenant-wide view.
    """
    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        response = client.get(
            f"{GRAPH_BASE}/me/followedSites",
            headers=headers,
        )
        if response.status_code == 400:
            # /me/followedSites isn't valid in app-only mode; surface
            # a helpful error rather than the raw Graph 400.
            raise RuntimeError(
                "sp_followed_sites is not available in service-principal "
                "(app-only) auth mode — there's no signed-in user. Use "
                "sp_sites() for a tenant-wide view.",
            )
        response.raise_for_status()
        return _extract_sites(response.json())
    finally:
        if http is None:
            client.close()


def _extract_sites(payload: dict[str, Any]) -> list[dict[str, Any]]:
    """Normalise a Graph site-list response into our consistent shape."""
    raw = payload.get("value", [])
    if not isinstance(raw, list):
        return []
    return [_one_site(entry) for entry in raw if isinstance(entry, dict)]


def _one_site(entry: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": str(entry.get("id") or ""),
        "name": str(entry.get("displayName") or entry.get("name") or ""),
        "web_url": str(entry.get("webUrl") or ""),
        "description": str(entry.get("description") or ""),
        "last_modified": str(entry.get("lastModifiedDateTime") or ""),
    }
