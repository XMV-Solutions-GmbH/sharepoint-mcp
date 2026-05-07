# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Recycle-bin tools (closes #50, partial).

- `sp_trash_list(site_url)` — list items in the site's recycle bin.

`sp_trash_restore` is **deferred**. Microsoft Graph's beta endpoint
at `/beta/sites/{id}/recycleBin/items` exposes the listing, but no
`/restore` action is currently documented or available on
recycleBinItem at the site scope (only on SharePoint Embedded
fileStorageContainer recycleBin items). Until Microsoft surfaces a
restore action — or we add a SharePoint REST API fallback — restore
must be done via the SharePoint web UI. Tracked in a follow-up
ticket; see CHANGELOG / issue #50 thread.

`sp_trash_list` uses Microsoft Graph's `/beta` endpoint. The
site-level recycle-bin API has not been promoted to v1.0 (as of
2026-05-07). Beta is stable enough that production tools rely on
it (SharePoint web UI, admin center), but Microsoft reserves the
right to change the schema. We pin to the documented beta shape
and will migrate to v1.0 when it lands. See
<https://learn.microsoft.com/en-us/graph/api/recyclebin-list-items?view=graph-rest-beta>.

Result shape for `sp_trash_list`:

    {
        "id": "<recycle-bin-item-id>",
        "name": "<filename>",
        "size": <int bytes>,
        "deleted_date_time": "<ISO datetime>",
        "deleted_from_location": "<original folder path>",
        "deleted_by": "<display name or empty>",
    }
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
    limit: int = 200,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> list[dict[str, Any]]:
    """List items in the recycle bin of `site_url`, newest first.

    Returns up to `limit` items. The recycle bin can hold many
    thousands of items on busy sites; we paginate via `@odata.nextLink`
    until either `limit` is reached or there are no more pages. Default
    200 = Microsoft's default page size, so the typical case is a
    single round-trip.

    Each item: id, name, size, deleted_date_time,
    deleted_from_location, deleted_by. Empty list when the recycle
    bin is empty.

    Raises:
        ValueError: empty / blank site_url, URL points at a file, or
            limit < 1.
        httpx.HTTPStatusError: any non-2xx from Graph beta. Common:
            403 if the user lacks site-collection scope; 404 if the
            tenant has the recycle-bin endpoint disabled.
    """
    if not site_url or not site_url.strip():
        raise ValueError("sp_trash_list requires a non-empty site_url")
    if limit < 1:
        raise ValueError(f"limit must be >= 1, got {limit!r}")
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
        next_url: str | None = (
            f"{GRAPH_BETA_BASE}/sites/{site_id}/recycleBin/items"
            "?$orderby=deletedDateTime+desc&$top=200"
        )
        results: list[dict[str, Any]] = []
        while next_url and len(results) < limit:
            response = client.get(next_url, headers=headers)
            response.raise_for_status()
            payload = response.json()
            results.extend(_extract_trash_items(payload))
            next_link = payload.get("@odata.nextLink")
            next_url = str(next_link) if next_link else None
        return results[:limit]
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


