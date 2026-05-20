# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_drive_file_move — move or rename a drive file or folder.

Combines two Graph capabilities in a single PATCH call:
- Reparent: supply a new ``parentReference.id`` to move to a different folder.
- Rename: supply a new ``name`` to rename in place.
- Both: supply both to move and rename simultaneously.

Graph API:
    PATCH /drives/{drive_id}/items/{item_id}
    body: {"parentReference": {"id": "<dest_folder_id>"}, "name": "<new_name>"}

The Graph response is the updated driveItem. We return the resolved source path,
the destination path, and the item's new webUrl.

Implements GitHub issue #95.
"""

from __future__ import annotations

from typing import Any

import httpx

from sharepoint_mcp.auth import get_token
from sharepoint_mcp.tools._common import (
    GRAPH_BASE,
    parse_sharepoint_url,
    resolve_drive_item,
    resolve_drive_item_full,
    resolve_site_id,
)


def move_file(
    site_url: str,
    source_path: str,
    destination_path: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> dict[str, Any]:
    """Move (and optionally rename) a drive file or folder.

    ``source_path`` and ``destination_path`` are both drive-relative paths
    (e.g. ``"2026/Q2/old.md"`` → ``"Archive/2026/Q2/old.md"``).

    The destination is interpreted as the **full path of the item after the
    move**, not the destination folder. The last segment is used as the
    (possibly new) name; all preceding segments must refer to an existing
    folder.

    Returns a dict with:
    - ``moved``: ``True``
    - ``source``: normalised source path
    - ``destination``: normalised destination path
    - ``web_url``: new SharePoint web URL of the item

    Raises:
        ValueError: empty inputs or paths.
        httpx.HTTPStatusError: non-2xx Graph response, including 404 if
            source or destination parent folder does not exist.
        sharepoint_mcp.auth.AuthRequiredError: no cached token for ``profile``.
    """
    if not site_url or not site_url.strip():
        raise ValueError("sp_drive_file_move requires a non-empty site_url")
    if not source_path or not source_path.strip():
        raise ValueError("sp_drive_file_move requires a non-empty source_path")
    if not destination_path or not destination_path.strip():
        raise ValueError("sp_drive_file_move requires a non-empty destination_path")

    src = source_path.strip().strip("/")
    dst = destination_path.strip().strip("/")

    if not src:
        raise ValueError("sp_drive_file_move: source_path contains no path segments")
    if not dst:
        raise ValueError("sp_drive_file_move: destination_path contains no path segments")

    hostname, site_path, _ = parse_sharepoint_url(site_url)
    token = get_token(profile)
    headers: dict[str, str] = {"Authorization": f"Bearer {token}"}
    json_headers = {**headers, "Content-Type": "application/json"}

    owned = http is None
    client = http if http is not None else httpx.Client(timeout=30.0, follow_redirects=True)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)

        # Resolve source to (drive_id, item_id).
        drive_id, item_id = resolve_drive_item(client, site_id, src, headers=headers)

        # Split destination into parent folder path + new name.
        if "/" in dst:
            dest_parent_path, new_name = dst.rsplit("/", 1)
        else:
            # Destination is at drive root — rename only, no reparent needed
            # unless the item is already at root.
            dest_parent_path = ""
            new_name = dst

        # Resolve destination parent folder to get its item id.
        if dest_parent_path:
            dest_parent_item = resolve_drive_item_full(
                client, site_id, dest_parent_path, headers=headers
            )
            dest_folder_id = str(dest_parent_item["id"])
            dest_folder_drive_id = str(dest_parent_item["parentReference"]["driveId"])
        else:
            # Root of the default drive — fetch drive root id.
            root_resp = client.get(
                f"{GRAPH_BASE}/sites/{site_id}/drive/root",
                headers=headers,
            )
            root_resp.raise_for_status()
            root = root_resp.json()
            dest_folder_id = str(root["id"])
            parent_ref = root.get("parentReference") or {}
            dest_folder_drive_id = str(parent_ref["driveId"]) if parent_ref else drive_id

        patch_body: dict[str, Any] = {
            "parentReference": {
                "driveId": dest_folder_drive_id,
                "id": dest_folder_id,
            },
            "name": new_name,
        }

        response = client.patch(
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}",
            headers=json_headers,
            json=patch_body,
        )
        response.raise_for_status()
        updated_item = response.json()

        return {
            "moved": True,
            "source": src,
            "destination": dst,
            "web_url": str(updated_item.get("webUrl") or ""),
        }
    finally:
        if owned:
            client.close()
