# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_file_history — list version history of a SharePoint file.

Read-only. Wraps `GET /drives/{id}/items/{id}/versions`.

Three Graph calls per invocation: site lookup → driveItem lookup →
versions list. All read-only.

The returned shape mirrors what the agent typically wants for an
audit-trail walk: id, lastModifiedDateTime, the user who saved that
version, and the version's size.

`comment` is **NOT** populated — Microsoft Graph's `driveItemVersion`
resource doesn't expose the per-version checkin comment via the v1.0
endpoints. The comment landing in SharePoint's audit log on a
sp_save_file call is real (visible in the SharePoint web UI's version
history), but reading it back through Graph is a documented
limitation. If/when Microsoft adds the field, we'll surface it here.
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


def history(
    url: str,
    *,
    limit: int = 20,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> list[dict[str, Any]]:
    """List the version history of a SharePoint file.

    Returns up to `limit` versions, newest first. Each entry has:

    - `id` — version identifier (e.g. "3.0", "3.1") — pass to
      `sp_get_file_version` to fetch that version's content.
    - `last_modified` — ISO datetime when the version was created.
    - `last_modified_by` — display-name or email of the user who
      created the version.
    - `size` — bytes.

    Raises:
        ValueError: empty URL or URL doesn't point at a file.
        httpx.HTTPStatusError: non-2xx Graph response.
        sharepoint_mcp.auth.AuthRequiredError: no cached token.
    """
    if not url or not url.strip():
        raise ValueError("sp_file_history requires a non-empty url")
    if limit <= 0:
        raise ValueError(f"limit must be positive, got {limit}")

    hostname, site_path, item_path = parse_sharepoint_url(url)
    if not item_path:
        raise ValueError(f"sp_file_history needs a file URL, got {url!r}")

    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        drive_id, item_id = resolve_drive_item(client, site_id, item_path, headers=headers)

        versions_response = client.get(
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/versions",
            headers=headers,
            params={"$top": limit, "$orderby": "lastModifiedDateTime desc"},
        )
        versions_response.raise_for_status()
        return _extract_versions(versions_response.json())
    finally:
        if http is None:
            client.close()


def _extract_versions(payload: dict[str, Any]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for v in payload.get("value", []):
        out.append(
            {
                "id": str(v.get("id") or ""),
                "last_modified": str(v.get("lastModifiedDateTime") or ""),
                "last_modified_by": _extract_user(v.get("lastModifiedBy")),
                "size": int(v.get("size") or 0),
            },
        )
    return out


def _extract_user(modified_by: dict[str, Any] | None) -> str | None:
    if not modified_by:
        return None
    user = modified_by.get("user")
    if not isinstance(user, dict):
        return None
    return user.get("displayName") or user.get("email")
