# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_publish — upload a NEW local file as a new SharePoint document.

The "draft + promote" use case from `docs/app-concept.md`: agent
drafts a document locally, then publishes it to a SharePoint folder
as a brand-new file. Distinct from `sp_save`, which checks in an
edited copy of an *existing* checked-out file.

Refuses if the target path already exists — the caller should use
`sp_open` + `sp_save` to update existing files (gives them an audit
comment + version history). Distinct semantics, explicit error
message rather than silent overwrite.

Two Graph calls:

1. `GET /sites/{id}/drive/root:/{path}` — check if target exists
   (404 means free to publish).
2. `PUT /sites/{id}/drive/root:/{path}:/content` — upload the file.

The created driveItem has `eTag` and `webUrl` populated by Microsoft
on the response; we surface them so the caller can persist a
reference.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import httpx

from sharepoint_mcp.auth import get_token
from sharepoint_mcp.tools._common import (
    GRAPH_BASE,
    parse_sharepoint_url,
    resolve_drive_item_full,
    resolve_site_id,
)


def publish(
    local_path: str,
    target_folder_url: str,
    *,
    name: str | None = None,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> dict[str, Any]:
    """Upload `local_path` as a new file under `target_folder_url`.

    `target_folder_url` is the human-readable URL of a SharePoint
    folder (e.g. `https://contoso.sharepoint.com/sites/foo/Shared
    Documents/drafts`). The new file's name defaults to
    `basename(local_path)` unless `name` is given.

    Returns a dict with `name`, `web_url`, `etag`, `size`,
    `last_modified` populated from Microsoft's response.

    Raises:
        ValueError: empty / blank inputs.
        FileNotFoundError: `local_path` doesn't exist or isn't a file.
        FileExistsError: the target file is already present at the
            SharePoint URL — use sp_open + sp_save to edit existing
            files; sp_publish is for new ones only.
        httpx.HTTPStatusError: any other non-2xx from Graph.
        sharepoint_mcp.auth.AuthRequiredError: no cached token for
            `profile`.
    """
    if not local_path or not local_path.strip():
        raise ValueError("sp_publish requires a non-empty local_path")
    if not target_folder_url or not target_folder_url.strip():
        raise ValueError("sp_publish requires a non-empty target_folder_url")

    src = Path(local_path)
    if not src.exists():
        raise FileNotFoundError(f"Local file not found: {local_path!r}")
    if not src.is_file():
        raise FileNotFoundError(f"Local path is not a file: {local_path!r}")

    filename = name or src.name
    if not filename or "/" in filename or "\\" in filename:
        raise ValueError(
            f"`name` must be a bare filename without path separators, got {filename!r}",
        )

    hostname, site_path, folder_path = parse_sharepoint_url(target_folder_url)

    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=120.0, follow_redirects=True)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)

        # Resolve the target folder. If folder_path is empty, the target
        # is the default drive's root on the site. Library fallback in
        # resolve_drive_item_full transparently handles non-default
        # libraries (Site Assets, custom document libraries).
        if folder_path:
            folder = resolve_drive_item_full(client, site_id, folder_path, headers=headers)
            drive_id = folder["parentReference"]["driveId"]
            folder_id = folder["id"]
            existence_url = f"{GRAPH_BASE}/drives/{drive_id}/items/{folder_id}:/{filename}"
            upload_url = f"{GRAPH_BASE}/drives/{drive_id}/items/{folder_id}:/{filename}:/content"
        else:
            existence_url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{filename}"
            upload_url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{filename}:/content"

        # Check if target file already exists.
        existence_response = client.get(existence_url, headers=headers)
        if existence_response.status_code == 200:
            raise FileExistsError(
                f"Target already exists at {target_folder_url!r}/{filename!r}. "
                "Use sp_open + sp_save to update existing files (gives proper "
                "version history with audit comment).",
            )
        if existence_response.status_code != 404:
            existence_response.raise_for_status()

        # Upload via PUT /content. Microsoft creates the driveItem on the
        # fly. Response is the full driveItem.
        upload_response = client.put(
            upload_url,
            headers=headers,
            content=src.read_bytes(),
        )
        upload_response.raise_for_status()
        item = upload_response.json()

        return {
            "name": str(item.get("name") or filename),
            "web_url": str(item.get("webUrl") or ""),
            "etag": str(item.get("eTag") or ""),
            "size": int(item.get("size") or 0),
            "last_modified": str(item.get("lastModifiedDateTime") or ""),
        }
    finally:
        if http is None:
            client.close()
