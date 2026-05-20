# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_drive_change_track — delta-query change tracking on a SharePoint drive (closes #51).

Microsoft Graph's delta endpoint returns the set of driveItems
changed since a stored cursor:

- First call: `GET /drives/{drive-id}/root/delta` — returns
  every item currently visible plus an initial cursor.
- Subsequent calls: pass the cursor back, get only the items that
  changed (created / modified / deleted) since the cursor was issued.

Pagination: Graph splits large responses across pages
(`@odata.nextLink`); the final page carries `@odata.deltaLink`,
which is the cursor for the next round. We follow the chain
internally and surface a single (items, cursor) tuple to callers.

Cursor opacity: callers store the cursor blob and pass it back; we
make no assumption about its format and don't parse it. The blob
can be a Graph URL (most common) or anything Microsoft chooses to
emit in future revisions.

Result shape:

    {
        "items": [
            {
                "id": str,
                "name": str,                # "" for deleted items
                "type": "file"|"folder"|"deleted",
                "web_url": str,
                "parent_path": str,         # parent driveItem path or ""
                "size": int,                # 0 for folders / deleted
                "last_modified": str,       # ISO datetime or ""
                "deleted": bool,            # True when type == "deleted"
            },
            ...
        ],
        "cursor": str,                      # opaque; pass back via `since=`
    }

The `deleted` facet on driveItem signals the item was removed (or
moved out of scope) since the last cursor. `name` and other fields
are typically empty for deletions; the agent should rely on `id`
and a side-channel mapping (e.g. its own previous results) to know
what was actually removed.
"""

from __future__ import annotations

from typing import Any

import httpx

from sharepoint_mcp.auth import get_token
from sharepoint_mcp.tools._common import (
    GRAPH_BASE,
    parse_sharepoint_url,
    resolve_site_id,
)


def changes(
    scope_url: str,
    *,
    since: str | None = None,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> dict[str, Any]:
    """Return items changed in a SharePoint drive since `since`, plus a new cursor.

    `scope_url` is a SharePoint site URL — delta runs on the site's
    default drive root.

    `since` is the opaque cursor from a previous `sp_drive_change_track` call,
    or None for the initial sync (which returns the full item set).

    Returns `{"items": [...], "cursor": str}`. Always returns a new
    cursor — store it for the next call.

    Raises:
        ValueError: empty scope_url, or scope_url points at a file/folder.
        httpx.HTTPStatusError: any non-2xx from Graph. 410 Gone happens
            when the cursor has expired (Graph rolls cursors after long
            inactivity); caller should drop the stale cursor and call
            again with `since=None` for a full re-sync.
    """
    if not scope_url or not scope_url.strip():
        raise ValueError("sp_drive_change_track requires a non-empty scope_url")
    hostname, site_path, item_path = parse_sharepoint_url(scope_url)
    if item_path:
        raise ValueError(
            f"sp_drive_change_track expects a site URL, not a file/folder URL "
            f"(got {scope_url!r}; item path {item_path!r}). "
            "Folder-scoped delta is deferred to a follow-up.",
        )

    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        if since:
            # The cursor IS the URL Graph wants for the next call.
            next_url: str | None = since
        else:
            site_id = resolve_site_id(client, hostname, site_path, headers=headers)
            next_url = f"{GRAPH_BASE}/sites/{site_id}/drive/root/delta"

        items: list[dict[str, Any]] = []
        cursor = since or ""
        # Follow @odata.nextLink until @odata.deltaLink terminates the chain.
        while next_url:
            response = client.get(next_url, headers=headers)
            response.raise_for_status()
            payload = response.json()
            items.extend(_extract_items(payload))
            delta_link = payload.get("@odata.deltaLink")
            if delta_link:
                cursor = str(delta_link)
                next_url = None
            else:
                next_link = payload.get("@odata.nextLink")
                next_url = str(next_link) if next_link else None
        return {"items": items, "cursor": cursor}
    finally:
        if http is None:
            client.close()


def _extract_items(payload: dict[str, Any]) -> list[dict[str, Any]]:
    raw = payload.get("value", []) or []
    if not isinstance(raw, list):
        return []
    return [_one_item(entry) for entry in raw if isinstance(entry, dict)]


def _one_item(entry: dict[str, Any]) -> dict[str, Any]:
    is_deleted = isinstance(entry.get("deleted"), dict)
    if is_deleted:
        item_type = "deleted"
    elif "folder" in entry:
        item_type = "folder"
    else:
        item_type = "file"

    parent_ref = entry.get("parentReference") or {}
    parent_path = ""
    if isinstance(parent_ref, dict):
        parent_path = str(parent_ref.get("path") or "")

    return {
        "id": str(entry.get("id") or ""),
        "name": str(entry.get("name") or ""),
        "type": item_type,
        "web_url": str(entry.get("webUrl") or ""),
        "parent_path": parent_path,
        "size": int(entry.get("size") or 0),
        "last_modified": str(entry.get("lastModifiedDateTime") or ""),
        "deleted": is_deleted,
    }
