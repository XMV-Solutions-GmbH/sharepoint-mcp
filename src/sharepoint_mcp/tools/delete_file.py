# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_drive_file_delete — delete a drive file or folder (soft-delete to recycle bin).

Mirrors sp_list_item_delete (list items) for drive files. Graph's DELETE on a
driveItem sends it to the site recycle bin, matching SharePoint's native
behaviour — recoverable for ~93 days via sp_site_trash_list.

Graph API:
    DELETE /drives/{drive_id}/items/{item_id}

SharePoint never hard-deletes on a plain DELETE; permanent deletion would
require an additional API that is deliberately not exposed here (too
dangerous in an LLM context).

Implements GitHub issue #92.
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


def delete_file(
    site_url: str,
    path: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> dict[str, Any]:
    """Delete a drive file or folder at ``path``, sending it to the recycle bin.

    ``path`` is drive-relative (e.g. ``"2026/Q2/report.md"`` or
    ``"Shared Documents/2026/Q2/report.md"`` — the library prefix is handled
    by the existing ``resolve_drive_item`` helper).

    Returns a dict with:
    - ``deleted``: ``True``
    - ``path``: the normalised drive-relative path that was deleted

    Raises:
        ValueError: empty ``site_url`` or ``path``.
        httpx.HTTPStatusError: non-2xx Graph response, including 404 if the
            path does not exist.
        sharepoint_mcp.auth.AuthRequiredError: no cached token for ``profile``.
    """
    if not site_url or not site_url.strip():
        raise ValueError("sp_drive_file_delete requires a non-empty site_url")
    if not path or not path.strip():
        raise ValueError("sp_drive_file_delete requires a non-empty path")

    hostname, site_path, item_path = parse_sharepoint_url(site_url)

    # Accept the path either as part of site_url or as a separate ``path``
    # argument. When site_url already contains a file path (e.g. copied from
    # the browser), item_path is non-empty and we use it; otherwise we use the
    # explicit ``path`` argument.
    drive_path = item_path if item_path else path.strip().strip("/")
    if not drive_path:
        raise ValueError("sp_drive_file_delete: could not determine a drive-relative path")

    token = get_token(profile)
    headers: dict[str, str] = {"Authorization": f"Bearer {token}"}

    owned = http is None
    client = http if http is not None else httpx.Client(timeout=30.0, follow_redirects=True)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        drive_id, item_id = resolve_drive_item(client, site_id, drive_path, headers=headers)

        response = client.delete(
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}",
            headers=headers,
        )
        response.raise_for_status()

        return {"deleted": True, "path": drive_path}
    finally:
        if owned:
            client.close()
