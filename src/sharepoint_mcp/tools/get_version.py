# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_get_file_version — download a specific historical version of a file.

Read-only. Wraps `GET /drives/{id}/items/{id}/versions/{version-id}/content`.

Same overall shape as `sp_read_file` (writes the bytes to a temp file with
the original extension preserved, returns the path) but addressed by
version-id instead of "current". Use `sp_file_history` first to discover
which version-id to fetch.

Three Graph calls per invocation: site lookup → driveItem lookup →
versions content GET (which Microsoft serves from a CDN URL after a
302; httpx follows redirects).
"""

from __future__ import annotations

import os
import tempfile
from pathlib import Path

import httpx

from sharepoint_mcp.auth import get_token
from sharepoint_mcp.tools._common import (
    GRAPH_BASE,
    parse_sharepoint_url,
    resolve_drive_item,
    resolve_site_id,
)

TEMP_FILE_PREFIX = "sharepoint-mcp-version-"


def get_version(
    url: str,
    version_id: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> str:
    """Download a specific historical version's content. Return local temp path.

    `version_id` is the id from `sp_file_history`'s response (e.g. "3.0").
    The downloaded bytes are written to a temp file with the original
    file's extension preserved, suffixed with `_v<version-id>` to make
    it obvious which version it is when multiple are downloaded.

    Raises:
        ValueError: empty inputs / URL points at site or folder.
        httpx.HTTPStatusError: 404 if the version doesn't exist; other
            non-2xx as raised by Microsoft.
        sharepoint_mcp.auth.AuthRequiredError: no cached token.
    """
    if not url or not url.strip():
        raise ValueError("sp_get_file_version requires a non-empty url")
    if not version_id or not version_id.strip():
        raise ValueError("sp_get_file_version requires a non-empty version_id")

    hostname, site_path, item_path = parse_sharepoint_url(url)
    if not item_path:
        raise ValueError(f"sp_get_file_version needs a file URL, got {url!r}")

    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=60.0, follow_redirects=True)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        drive_id, item_id = resolve_drive_item(client, site_id, item_path, headers=headers)

        content_response = client.get(
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/versions/{version_id}/content",
            headers=headers,
        )
        content_response.raise_for_status()
        return _write_temp(
            content_response.content,
            base_name=Path(item_path).stem,
            suffix=Path(item_path).suffix,
            version_id=version_id,
        )
    finally:
        if http is None:
            client.close()


def _write_temp(content: bytes, *, base_name: str, suffix: str, version_id: str) -> str:
    """Write bytes to a temp file with a `_v<id>` infix in the name."""
    safe_version = version_id.replace("/", "_").replace(" ", "_")
    fd, temp_path = tempfile.mkstemp(
        suffix=suffix,
        prefix=f"{TEMP_FILE_PREFIX}{base_name}_v{safe_version}_",
    )
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(content)
    except OSError:
        try:
            Path(temp_path).unlink()
        except FileNotFoundError:
            pass
        raise
    return temp_path
