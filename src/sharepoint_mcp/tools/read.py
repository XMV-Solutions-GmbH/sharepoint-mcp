# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_read — download a SharePoint file's content to a local temp file.

Read-only. Does NOT acquire a checkout / lock — that's `sp_open`'s job.
The temp file is written to the OS temp directory, with the original
file's extension preserved so editors / viewers can pick the right
handler. The agent (or test) is responsible for deciding what to do
with the path; the OS will eventually clean up the temp file as part
of normal `/tmp` hygiene.

Two Graph calls per invocation: site lookup → content GET. The
content GET follows redirects because Microsoft serves file content
from CDN URLs returned via 302.
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
    resolve_drive_item_full,
    resolve_site_id,
)

TEMP_FILE_PREFIX = "sharepoint-mcp-"


def read_file(
    url: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> str:
    """Download SharePoint file content to a local temp file. Return its path.

    `url` is the human-readable URL of a file inside a SharePoint
    document library, e.g.
    `https://contoso.sharepoint.com/sites/foo/Shared Documents/policy.docx`.

    Returns the absolute path of a temp file containing the downloaded
    bytes (extension preserved). The caller is responsible for using
    or removing the file; the OS cleans `/tmp` eventually.

    Raises:
        ValueError: input URL is empty / relative / points at a site
            or folder rather than a file.
        httpx.HTTPStatusError: Graph returned non-2xx (e.g. 404 if the
            file doesn't exist or the user can't see it).
        sharepoint_mcp.auth.AuthRequiredError: no usable cached token
            for `profile`.
    """
    if not url or not url.strip():
        raise ValueError("sp_read requires a non-empty url")

    hostname, site_path, item_path = parse_sharepoint_url(url)
    if not item_path:
        raise ValueError(
            f"sp_read needs a file URL, got a site/folder URL with no item path: {url!r}",
        )

    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=60.0, follow_redirects=True)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        item = resolve_drive_item_full(client, site_id, item_path, headers=headers)
        drive_id = item["parentReference"]["driveId"]
        item_id = item["id"]
        content_response = client.get(
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content",
            headers=headers,
        )
        content_response.raise_for_status()
        return _write_temp(content_response.content, suffix=Path(item_path).suffix)
    finally:
        if http is None:
            client.close()


def _write_temp(content: bytes, *, suffix: str) -> str:
    """Atomically write `content` to a fresh temp file and return its path.

    Cleans up the partial file if the write itself fails (the
    `mkstemp` already created a file on disk that needs removing).
    """
    fd, temp_path = tempfile.mkstemp(suffix=suffix, prefix=TEMP_FILE_PREFIX)
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
