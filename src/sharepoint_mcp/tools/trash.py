# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Recycle-bin tools (closes #50).

- `sp_trash_list(site_url)` — list items in the site's recycle bin.
- `sp_trash_restore(site_url, item_id)` — restore an item from the
  recycle bin to its original location.

**Both tools currently use Microsoft Graph's `/beta` endpoints.**
The site-level recycle-bin API has not yet been promoted to v1.0
(as of 2026-05-07). Beta endpoints are stable enough that other
tools rely on them (e.g. SharePoint web UI, SharePoint admin
center), but Microsoft reserves the right to change the schema.
We pin to the documented beta shape and will migrate to v1.0
when it lands. See <https://learn.microsoft.com/en-us/graph/api/recyclebin-list-items?view=graph-rest-beta>.

Result shape for `sp_trash_list`:

    {
        "id": "<recycle-bin-item-id>",
        "name": "<filename>",
        "size": <int bytes>,
        "deleted_date_time": "<ISO datetime>",
        "deleted_from_location": "<original folder path>",
        "deleted_by": "<display name or empty>",
    }

The `id` is what you pass to `sp_trash_restore`. The original location
helps the agent confirm "yes this is the file the user meant".
"""

from __future__ import annotations

from typing import Any

import httpx

from sharepoint_mcp.auth import get_token
from sharepoint_mcp.tools._common import parse_sharepoint_url, resolve_site_id

GRAPH_BETA_BASE = "https://graph.microsoft.com/beta"


def trash_list(
    site_url: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> list[dict[str, Any]]:
    """List items in the recycle bin of `site_url`.

    Returns a list of recycle-bin items with id, name, size,
    deleted_date_time, deleted_from_location, deleted_by. Empty list
    when the recycle bin is empty.

    Raises:
        ValueError: empty / blank site_url, or URL points at a file
            rather than a site.
        httpx.HTTPStatusError: any non-2xx from Graph beta. Common
            ones: 403 if the user lacks site-collection scope; 404 if
            the tenant has the recycle-bin endpoint disabled.
    """
    if not site_url or not site_url.strip():
        raise ValueError("sp_trash_list requires a non-empty site_url")
    hostname, site_path, item_path = parse_sharepoint_url(site_url)
    if item_path:
        raise ValueError(
            f"sp_trash_list expects a site URL, not a file/folder URL "
            f"(got {site_url!r}; item path {item_path!r}).",
        )

    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        response = client.get(
            f"{GRAPH_BETA_BASE}/sites/{site_id}/recycleBin/items",
            headers=headers,
        )
        response.raise_for_status()
        return _extract_trash_items(response.json())
    finally:
        if http is None:
            client.close()


def trash_restore(
    site_url: str,
    item_id: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> dict[str, Any]:
    """Restore a single item from the recycle bin to its original location.

    `item_id` is the id field returned by `sp_trash_list`.

    Returns a dict with the restored item's id, name, web_url (when
    Graph populates them on restore — beta sometimes returns 204 No
    Content for a successful restore, in which case we return an
    empty dict — the agent should follow up with the appropriate
    read tool to confirm the file is back).

    Raises:
        ValueError: empty inputs.
        httpx.HTTPStatusError: any non-2xx from Graph beta. 404 means
            the item-id is unknown (already restored, or wrong site).
    """
    if not site_url or not site_url.strip():
        raise ValueError("sp_trash_restore requires a non-empty site_url")
    if not item_id or not item_id.strip():
        raise ValueError("sp_trash_restore requires a non-empty item_id")

    hostname, site_path, item_path = parse_sharepoint_url(site_url)
    if item_path:
        raise ValueError(
            f"sp_trash_restore expects a site URL, not a file/folder URL "
            f"(got {site_url!r}; item path {item_path!r}).",
        )

    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        response = client.post(
            f"{GRAPH_BETA_BASE}/sites/{site_id}/recycleBin/items/{item_id}/restore",
            headers=headers,
        )
        response.raise_for_status()
        # 200/201 returns the restored item; 204 returns nothing.
        if response.status_code in (200, 201) and response.content:
            try:
                payload = response.json()
            except ValueError:
                return {}
            return _restored_item_summary(payload)
        return {}
    finally:
        if http is None:
            client.close()


def _extract_trash_items(payload: dict[str, Any]) -> list[dict[str, Any]]:
    raw = payload.get("value", [])
    if not isinstance(raw, list):
        return []
    return [_one_trash_item(entry) for entry in raw if isinstance(entry, dict)]


def _one_trash_item(entry: dict[str, Any]) -> dict[str, Any]:
    deleted_by_raw = entry.get("deletedBy") or {}
    deleted_by_user = deleted_by_raw.get("user") if isinstance(deleted_by_raw, dict) else None
    if isinstance(deleted_by_user, dict):
        deleted_by = str(deleted_by_user.get("displayName") or deleted_by_user.get("email") or "")
    else:
        deleted_by = ""
    return {
        "id": str(entry.get("id") or ""),
        "name": str(entry.get("name") or ""),
        "size": int(entry.get("size") or 0),
        "deleted_date_time": str(entry.get("deletedDateTime") or ""),
        "deleted_from_location": str(entry.get("deletedFromLocation") or ""),
        "deleted_by": deleted_by,
    }


def _restored_item_summary(payload: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": str(payload.get("id") or ""),
        "name": str(payload.get("name") or ""),
        "web_url": str(payload.get("webUrl") or ""),
    }
