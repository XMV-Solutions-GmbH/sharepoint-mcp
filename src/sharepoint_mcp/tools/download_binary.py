# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_download_binary — download a SharePoint file's bytes as base64.

Wraps `GET /drives/{id}/items/{id}/content` (same CDN-redirect path as
`sp_read_file`) but returns the bytes base64-encoded in a JSON envelope
rather than writing them to a temp file.  Intended for small non-text
assets (images, PDFs, Office files) that an agent needs to embed or
inspect inline.

A hard 10 MB guard prevents accidentally routing large files through
the agent context window.  For larger files use `sp_read_file` and
process the temp-file path out-of-band.

Three Graph calls per invocation: site lookup → driveItem lookup (which
also yields the MIME type and size) → content GET (follows CDN redirect).
"""

from __future__ import annotations

import base64
from typing import Any

import httpx

from sharepoint_mcp.auth import get_token
from sharepoint_mcp.tools._common import (
    GRAPH_BASE,
    parse_sharepoint_url,
    resolve_drive_item_full,
    resolve_site_id,
)

MAX_BYTES = 10 * 1024 * 1024  # 10 MB


def download_binary(
    url: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> dict[str, Any]:
    """Download a SharePoint file's bytes and return them base64-encoded.

    `url` is the human-readable URL of a file inside a SharePoint
    document library.

    Returns a dict with:
    - ``filename``  — original file name from SharePoint metadata
    - ``mime_type`` — from Graph's ``file.mimeType``; falls back to
      ``"application/octet-stream"`` when Graph omits it
    - ``size_bytes`` — byte length of the downloaded content
    - ``base64``    — standard base64 (RFC 4648) of the raw bytes;
      decode with ``base64.b64decode(result["base64"])``

    Raises:
        ValueError: `url` is empty/relative/not a file, or the file
            exceeds the 10 MB size guard.
        httpx.HTTPStatusError: Graph returned non-2xx (e.g. 404 if the
            file doesn't exist or the user can't see it).
        sharepoint_mcp.auth.AuthRequiredError: no usable cached token.
    """
    if not url or not url.strip():
        raise ValueError("sp_download_binary requires a non-empty url")

    hostname, site_path, item_path = parse_sharepoint_url(url)
    if not item_path:
        raise ValueError(
            f"sp_download_binary needs a file URL, got a site/folder URL: {url!r}",
        )

    token = get_token(profile)
    auth_header = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=60.0, follow_redirects=True)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=auth_header)
        item = resolve_drive_item_full(client, site_id, item_path, headers=auth_header)
        drive_id = str(item["parentReference"]["driveId"])
        item_id = str(item["id"])
        filename = str(item.get("name", item_path.rsplit("/", 1)[-1]))
        declared_size: int = int(item.get("size", 0))

        if declared_size > MAX_BYTES:
            raise ValueError(
                f"sp_download_binary: file '{filename}' is {declared_size:,} bytes "
                f"which exceeds the 10 MB limit ({MAX_BYTES:,} bytes). "
                "Use sp_read_file to write the file to a local temp path instead."
            )

        mime_type: str = (item.get("file") or {}).get("mimeType") or "application/octet-stream"

        content_resp = client.get(
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content",
            headers=auth_header,
        )
        content_resp.raise_for_status()
        content = content_resp.content

        if len(content) > MAX_BYTES:
            raise ValueError(
                f"sp_download_binary: downloaded content for '{filename}' is "
                f"{len(content):,} bytes which exceeds the 10 MB limit ({MAX_BYTES:,} bytes). "
                "Use sp_read_file to write the file to a local temp path instead."
            )

        return {
            "filename": filename,
            "mime_type": mime_type,
            "size_bytes": len(content),
            "base64": base64.b64encode(content).decode("ascii"),
        }
    finally:
        if http is None:
            client.close()
