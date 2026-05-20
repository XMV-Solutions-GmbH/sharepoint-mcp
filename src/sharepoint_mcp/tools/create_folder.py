# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_drive_folder_create — create a folder hierarchy in a SharePoint document library.

Complementary to the checkout/checkin lifecycle: sp_drive_file_checkout and
sp_drive_file_checkin require an existing item to check out. This tool creates
new folders where none exist yet.

Graph API: one POST per segment that doesn't exist yet:

    POST /sites/{site_id}/drive/root/children           (for root-level segments)
    POST /sites/{site_id}/drive/root:/{parent}/children (for nested segments)
    body: {"name": "…", "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}

409 + error.code == "nameAlreadyExists" means the folder already exists — we
skip it and continue (idempotent). A 409 with any other code is propagated.

Implements GitHub issue #86.
"""

from __future__ import annotations

from typing import Any

import httpx

from sharepoint_mcp.auth import get_token
from sharepoint_mcp.tools._common import GRAPH_BASE, parse_sharepoint_url, resolve_site_id

# Library name prefixes that users may include in the path but that refer to
# the default document library — we strip them so both
#   path="Shared Documents/2026/Q2"  and  path="2026/Q2"  work identically.
_LIBRARY_PREFIXES = ("shared documents/", "documents/")


def _normalize_path(path: str) -> list[str]:
    """Return non-empty folder segments, stripping default library prefix."""
    p = path.strip()
    lower = p.lower()
    for prefix in _LIBRARY_PREFIXES:
        if lower.startswith(prefix):
            p = p[len(prefix) :]
            break
    return [s for s in p.strip("/").split("/") if s]


def create_folder(
    site_url: str,
    path: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> dict[str, Any]:
    """Create a folder hierarchy at `path` in the site's default document library.

    `path` is relative to the document library root (e.g. ``"2026/Q2/Reports"``).
    A leading ``"Shared Documents/"`` prefix is stripped for convenience. Intermediate
    folders that do not yet exist are created automatically (recursive mkdir semantics).
    Existing folders are skipped without error (idempotent).

    Returns a dict with:
    - ``created``: list of path segments successfully created this call
    - ``already_existed``: list of segments that already existed
    - ``web_url``: web URL of the deepest folder (from Graph response or a
      follow-up GET when all segments already existed)

    Raises:
        ValueError: empty ``site_url`` or ``path``, or path resolves to no segments.
        httpx.HTTPStatusError: any non-2xx Graph response other than a
            nameAlreadyExists 409 (e.g. a name collision with a *file*).
        sharepoint_mcp.auth.AuthRequiredError: no cached token for ``profile``.
    """
    if not site_url or not site_url.strip():
        raise ValueError("sp_drive_folder_create requires a non-empty site_url")
    if not path or not path.strip():
        raise ValueError("sp_drive_folder_create requires a non-empty path")

    segments = _normalize_path(path)
    if not segments:
        raise ValueError(
            f"sp_drive_folder_create: path {path!r} contains no valid folder segments "
            "after stripping the library prefix and slashes",
        )

    hostname, site_path, _ = parse_sharepoint_url(site_url)
    token = get_token(profile)
    auth_headers: dict[str, str] = {"Authorization": f"Bearer {token}"}
    json_headers = {**auth_headers, "Content-Type": "application/json"}

    owned = http is None
    client = http if http is not None else httpx.Client(timeout=60.0, follow_redirects=True)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=auth_headers)

        created: list[str] = []
        already_existed: list[str] = []
        last_web_url: str = ""

        for i, segment in enumerate(segments):
            parent_parts = segments[:i]
            current_full_path = "/".join(segments[: i + 1])

            if parent_parts:
                parent_str = "/".join(parent_parts)
                endpoint = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{parent_str}:/children"
            else:
                endpoint = f"{GRAPH_BASE}/sites/{site_id}/drive/root/children"

            resp = client.post(
                endpoint,
                headers=json_headers,
                json={
                    "name": segment,
                    "folder": {},
                    "@microsoft.graph.conflictBehavior": "fail",
                },
            )

            if resp.status_code == 201:
                created.append(current_full_path)
                last_web_url = str(resp.json().get("webUrl") or "")
            elif resp.status_code == 409:
                err_code = resp.json().get("error", {}).get("code", "")
                if err_code == "nameAlreadyExists":
                    already_existed.append(current_full_path)
                else:
                    resp.raise_for_status()
            else:
                resp.raise_for_status()

        # When all segments already existed we don't have a web_url from a
        # creation response — fetch it with one extra GET.
        if not last_web_url and segments:
            full_path = "/".join(segments)
            get_resp = client.get(
                f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{full_path}",
                headers=auth_headers,
            )
            if get_resp.status_code == 200:
                last_web_url = str(get_resp.json().get("webUrl") or "")

        return {
            "created": created,
            "already_existed": already_existed,
            "web_url": last_web_url,
        }
    finally:
        if owned:
            client.close()
