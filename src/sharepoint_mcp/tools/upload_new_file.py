# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_upload_new_file — upload a new file to SharePoint from inline base64 content.

Complementary to sp_publish (which takes a local file path): this tool accepts
the file content as a base64-encoded string, so agents can create small text
files, JSON documents, CSV exports, etc. without first writing them to disk.

For large files (> 4 MB) or binary blobs that the agent already has on disk,
use sp_publish instead — it streams from the local file and supports resumable
uploads for very large files.

Graph API:
1. GET  /sites/{site_id}/drive/root:/{path}          — existence check
2. PUT  /sites/{site_id}/drive/root:/{path}/content  — create

If the file already exists (step 1 returns 200), raises FileAlreadyExistsError
with instructions to use sp_open + sp_save instead.

Implements GitHub issue #87.
"""

from __future__ import annotations

import base64
import binascii
import mimetypes
from typing import Any

import httpx

from sharepoint_mcp.auth import get_token
from sharepoint_mcp.tools._common import GRAPH_BASE, parse_sharepoint_url, resolve_site_id

# Inline base64 uploads are capped at 4 MB (decoded). For larger files
# agents should write to disk and call sp_publish, which supports resumable
# uploads for files over the chunked-upload threshold.
MAX_INLINE_BYTES = 4 * 1024 * 1024

# Paths may start with these default-library prefixes — strip for convenience.
_LIBRARY_PREFIXES = ("shared documents/", "documents/")


class FileAlreadyExistsError(RuntimeError):
    """Raised when the target path already contains a file.

    Downstream MCP handlers should surface this as a recoverable condition:
    tell the agent to use ``sp_open`` + ``sp_save`` if it wants to edit the
    existing file, or choose a different path.
    """


def _normalize_path(path: str) -> str:
    """Return the drive-relative path without default library prefix or leading slash."""
    p = path.strip()
    lower = p.lower()
    for prefix in _LIBRARY_PREFIXES:
        if lower.startswith(prefix):
            p = p[len(prefix):]
            break
    return p.strip("/")


def upload_new_file(
    site_url: str,
    path: str,
    content: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> dict[str, Any]:
    """Upload ``content`` (base64-encoded) as a new file at ``path`` in SharePoint.

    ``path`` is drive-relative (e.g. ``"2026/Q2/report.md"`` or
    ``"Shared Documents/2026/Q2/report.md"`` — the library prefix is stripped).
    The parent folder must already exist; create it first with sp_create_folder if
    needed.

    Returns a dict with:
    - ``item_id``: Graph driveItem id
    - ``etag``: eTag of the newly created file (use with sp_save for conflict detection)
    - ``web_url``: SharePoint web URL
    - ``size``: file size in bytes

    Raises:
        ValueError: empty inputs, invalid base64, content exceeds 4 MB, or
            path resolves to no filename.
        FileAlreadyExistsError: a file already exists at ``path`` — use sp_open
            + sp_save to edit it.
        httpx.HTTPStatusError: any other non-2xx Graph response (e.g. 404 if
            the parent folder doesn't exist — create it with sp_create_folder first).
        sharepoint_mcp.auth.AuthRequiredError: no cached token for ``profile``.
    """
    if not site_url or not site_url.strip():
        raise ValueError("sp_upload_new_file requires a non-empty site_url")
    if not path or not path.strip():
        raise ValueError("sp_upload_new_file requires a non-empty path")
    if content is None:
        raise ValueError("sp_upload_new_file requires content (base64-encoded string)")

    try:
        raw_bytes = base64.b64decode(content, validate=True)
    except (ValueError, binascii.Error) as exc:
        raise ValueError(f"sp_upload_new_file: content is not valid base64: {exc}") from exc

    if len(raw_bytes) > MAX_INLINE_BYTES:
        raise ValueError(
            f"sp_upload_new_file: decoded content is {len(raw_bytes):,} bytes, "
            f"which exceeds the 4 MB inline limit. Write the content to a local "
            "file and use sp_publish for large uploads.",
        )

    normalized = _normalize_path(path)
    if not normalized:
        raise ValueError("sp_upload_new_file: path must include a filename")

    filename = normalized.rsplit("/", 1)[-1]
    if not filename:
        raise ValueError("sp_upload_new_file: path must end with a filename, not a slash")

    content_type, _ = mimetypes.guess_type(filename)
    if not content_type:
        content_type = "application/octet-stream"

    hostname, site_path, _ = parse_sharepoint_url(site_url)
    token = get_token(profile)
    auth_headers: dict[str, str] = {"Authorization": f"Bearer {token}"}

    owned = http is None
    client = http if http is not None else httpx.Client(timeout=60.0, follow_redirects=True)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=auth_headers)

        # Check whether the target file already exists.
        exist_url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{normalized}"
        exist_resp = client.get(exist_url, headers=auth_headers)
        if exist_resp.status_code == 200:
            raise FileAlreadyExistsError(
                f"A file already exists at {path!r} in this SharePoint site. "
                "To edit it, use sp_open to check it out, make your changes, "
                "then sp_save to commit with an audit comment.",
            )
        if exist_resp.status_code != 404:
            exist_resp.raise_for_status()

        # Upload the file content. Graph creates the driveItem on the fly
        # and returns the full driveItem JSON on success (200 or 201).
        upload_url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{normalized}:/content"
        upload_resp = client.put(
            upload_url,
            headers={**auth_headers, "Content-Type": content_type},
            content=raw_bytes,
        )
        upload_resp.raise_for_status()
        item = upload_resp.json()

        return {
            "item_id": str(item.get("id") or ""),
            "etag": str(item.get("eTag") or ""),
            "web_url": str(item.get("webUrl") or ""),
            "size": int(item.get("size") or len(raw_bytes)),
        }
    finally:
        if owned:
            client.close()
