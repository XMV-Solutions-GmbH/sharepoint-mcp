# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_drive_file_copy — copy a drive file to a new path (async Graph operation).

Graph's copy endpoint is asynchronous: it returns 202 Accepted with a
``Location`` header pointing to an operation-status URL. We poll that URL
until the operation reaches ``completed`` or ``failed``.

Graph API:
    POST /drives/{drive_id}/items/{item_id}/copy
    body: {"parentReference": {"driveId": "...", "id": "<dest_folder_id>"}, "name": "<new_name>"}

    → 202 Accepted, Location: <operation_url>

    GET <operation_url>
    → {"status": "inProgress"|"completed"|"failed", "resourceLink": "..."}

The ``resourceLink`` in the completed status response is the webUrl of the new
item. If the operation doesn't complete within the timeout, we raise
``TimeoutError`` — the copy may still complete on the server side, but the
caller can verify with sp_drive_folder_list.

Implements GitHub issue #96.
"""

from __future__ import annotations

import time
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

_DEFAULT_TIMEOUT_SECONDS = 60
_POLL_INTERVAL_SECONDS = 1.0


def copy_file(
    site_url: str,
    source_path: str,
    destination_path: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
    timeout: int = _DEFAULT_TIMEOUT_SECONDS,
) -> dict[str, Any]:
    """Copy a drive file to ``destination_path``.

    ``source_path`` and ``destination_path`` are drive-relative paths
    (e.g. ``"Templates/contract.docx"`` → ``"Projects/ACME/contract.docx"``).

    The destination is interpreted as the **full path of the copy after
    creation**, not the destination folder. The last segment is the name
    of the new file; all preceding segments must refer to an existing folder.

    The Graph copy operation is asynchronous. This function polls until the
    operation completes (up to ``timeout`` seconds, default 60).

    Returns a dict with:
    - ``copied``: ``True``
    - ``source``: normalised source path
    - ``destination``: normalised destination path
    - ``web_url``: SharePoint web URL of the new copy

    Raises:
        ValueError: empty inputs or paths.
        TimeoutError: copy operation did not complete within ``timeout`` seconds.
        RuntimeError: Graph reports the copy operation as ``failed``.
        httpx.HTTPStatusError: non-2xx Graph response, including 404 if source
            or destination parent folder does not exist.
        sharepoint_mcp.auth.AuthRequiredError: no cached token for ``profile``.
    """
    if not site_url or not site_url.strip():
        raise ValueError("sp_drive_file_copy requires a non-empty site_url")
    if not source_path or not source_path.strip():
        raise ValueError("sp_drive_file_copy requires a non-empty source_path")
    if not destination_path or not destination_path.strip():
        raise ValueError("sp_drive_file_copy requires a non-empty destination_path")

    src = source_path.strip().strip("/")
    dst = destination_path.strip().strip("/")

    if not src:
        raise ValueError("sp_drive_file_copy: source_path contains no path segments")
    if not dst:
        raise ValueError("sp_drive_file_copy: destination_path contains no path segments")

    hostname, site_path, _ = parse_sharepoint_url(site_url)
    token = get_token(profile)
    headers: dict[str, str] = {"Authorization": f"Bearer {token}"}
    json_headers = {**headers, "Content-Type": "application/json"}

    owned = http is None
    client = http if http is not None else httpx.Client(timeout=30.0, follow_redirects=False)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)

        # Resolve source to (drive_id, item_id).
        drive_id, item_id = resolve_drive_item(client, site_id, src, headers=headers)

        # Split destination into parent folder path + new name.
        if "/" in dst:
            dest_parent_path, new_name = dst.rsplit("/", 1)
        else:
            dest_parent_path = ""
            new_name = dst

        # Resolve destination parent folder.
        if dest_parent_path:
            dest_parent_item = resolve_drive_item_full(
                client, site_id, dest_parent_path, headers=headers
            )
            dest_folder_id = str(dest_parent_item["id"])
            dest_folder_drive_id = str(dest_parent_item["parentReference"]["driveId"])
        else:
            root_resp = client.get(
                f"{GRAPH_BASE}/sites/{site_id}/drive/root",
                headers=headers,
            )
            root_resp.raise_for_status()
            root = root_resp.json()
            dest_folder_id = str(root["id"])
            dest_folder_drive_id = drive_id

        # Initiate async copy.
        copy_resp = client.post(
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/copy",
            headers=json_headers,
            json={
                "parentReference": {
                    "driveId": dest_folder_drive_id,
                    "id": dest_folder_id,
                },
                "name": new_name,
            },
        )
        # Handle all expected Graph copy-endpoint response codes before
        # calling raise_for_status, because httpx raises HTTPStatusError on
        # 3xx responses when follow_redirects=False.
        #
        # 200/201 — synchronous completion with item body
        # 202     — async copy; Location header points to operation-status URL
        # 303     — CDN redirect; Location header is the new item's URL directly
        # 4xx/5xx — error, raise
        status = copy_resp.status_code

        if status in (200, 201):
            item = copy_resp.json()
            return {
                "copied": True,
                "source": src,
                "destination": dst,
                "web_url": str(item.get("webUrl") or ""),
            }

        if status == 303:
            return {
                "copied": True,
                "source": src,
                "destination": dst,
                "web_url": copy_resp.headers.get("Location", ""),
            }

        if status != 202:
            copy_resp.raise_for_status()

        operation_url = copy_resp.headers.get("Location", "")
        if not operation_url:
            raise RuntimeError(
                "sp_drive_file_copy: Graph returned 202 but no Location header — "
                "cannot poll operation status"
            )

        # Poll for completion.
        web_url = _poll_copy_operation(client, operation_url, headers=headers, timeout=timeout)

        return {
            "copied": True,
            "source": src,
            "destination": dst,
            "web_url": web_url,
        }
    finally:
        if owned:
            client.close()


def _poll_copy_operation(
    client: httpx.Client,
    operation_url: str,
    *,
    headers: dict[str, str],
    timeout: int,
) -> str:
    """Poll ``operation_url`` until the copy completes. Returns the new item's webUrl."""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        resp = client.get(operation_url, headers=headers)

        # Graph sometimes responds to the operation-status poll with 303 See Other
        # (CDN redirect pattern) — treat it as synchronous completion.
        if resp.status_code == 303:
            return str(resp.headers.get("Location", ""))

        resp.raise_for_status()
        data = resp.json()
        status = str(data.get("status") or "")

        if status == "completed":
            resource_link = str(data.get("resourceLink") or "")
            return resource_link

        if status == "failed":
            error = data.get("error") or {}
            code = error.get("code") or "unknown"
            message = error.get("message") or "no details"
            raise RuntimeError(
                f"sp_drive_file_copy: Graph copy operation failed — {code}: {message}"
            )

        time.sleep(_POLL_INTERVAL_SECONDS)

    raise TimeoutError(
        f"sp_drive_file_copy: copy operation did not complete within {timeout} seconds. "
        "The copy may still be in progress on the server — check sp_drive_folder_list to verify."
    )
