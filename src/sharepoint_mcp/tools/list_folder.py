# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_list — list children of a SharePoint folder by URL.

Resolves a SharePoint URL into a Microsoft Graph drive item via two
calls:

1. `GET /sites/{hostname}:{site-path}` — site lookup by URL.
2. `GET /sites/{site-id}/drive/root[:/path:]/children` — list the
   default drive's root or a sub-folder.

The `/shares` endpoint would be a single-call shortcut, but it
requires the URL to have been explicitly shared via a sharing-link
— site-membership alone is not enough. The two-call site/drive
path works for any site URL the user is a member of, which is what
we want.

For v0.1, only the site's **default drive** ("Shared Documents") is
supported; URLs pointing at other libraries (SiteAssets, Style
Library, custom libraries) are out of scope and return a clear
error after lookup. Custom-library support can land in v0.2.

Module is named `list_folder` instead of `list` to avoid shadowing
the Python builtin.
"""

from __future__ import annotations

from typing import Any
from urllib.parse import unquote, urlparse

import httpx

from sharepoint_mcp.auth import get_token

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# SharePoint URLs that point at the default drive ("Shared Documents")
# may include the library name as a path segment in any of these
# common spellings. We strip it before passing the rest to Graph as
# a folder path within the drive.
_DEFAULT_LIBRARY_SEGMENTS = frozenset(
    {"Shared Documents", "Shared%20Documents", "Documents"},
)
_SITE_ROOT_SEGMENTS = frozenset({"sites", "teams", "personal"})


def _parse_sharepoint_url(url: str) -> tuple[str, str, str]:
    """Split a SharePoint URL into (hostname, site_path, folder_path).

    Examples (after URL-decoding):

        https://contoso.sharepoint.com/sites/foo
            → ("contoso.sharepoint.com", "/sites/foo", "")

        https://contoso.sharepoint.com/sites/foo/Shared Documents
            → ("contoso.sharepoint.com", "/sites/foo", "")

        https://contoso.sharepoint.com/sites/foo/Shared Documents/policies
            → ("contoso.sharepoint.com", "/sites/foo", "policies")

        https://contoso.sharepoint.com/sites/foo/policies
            → ("contoso.sharepoint.com", "/sites/foo", "policies")
              (assumes the path is already drive-relative)

    Folder path is drive-relative (NOT site-relative), with no
    leading slash. An empty folder path means "root of the default
    drive".
    """
    parsed = urlparse(url)
    if not parsed.netloc:
        raise ValueError(f"sp_list expects an absolute URL, got {url!r}")
    parts = [p for p in unquote(parsed.path).split("/") if p]

    if len(parts) >= 2 and parts[0] in _SITE_ROOT_SEGMENTS:
        site_path = f"/{parts[0]}/{parts[1]}"
        rest = parts[2:]
    elif not parts:
        site_path = ""
        rest = []
    else:
        # No /sites or /teams prefix — assume the host is a OneDrive
        # personal site or similar; the rest is drive-relative.
        site_path = ""
        rest = parts

    if rest and rest[0] in _DEFAULT_LIBRARY_SEGMENTS:
        rest = rest[1:]

    folder_path = "/".join(rest)
    return parsed.netloc, site_path, folder_path


def list_folder(
    url: str,
    *,
    limit: int = 100,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> list[dict[str, Any]]:
    """List the children of the SharePoint/OneDrive folder at `url`.

    `url` is the human-readable URL of a SharePoint folder, e.g.
    `https://contoso.sharepoint.com/sites/foo/Shared Documents/policies`,
    or just the site URL to list the root of the default drive.

    Returns at most `limit` items, each with `name`, `type`
    (`"folder"` / `"file"`), `size`, `last_modified`, `web_url`.

    Raises `ValueError` for invalid input,
    `httpx.HTTPStatusError` on non-2xx Graph responses (e.g. 404 if
    the site doesn't exist or the user can't see it),
    `sharepoint_mcp.auth.AuthRequiredError` if no usable cached
    token exists for `profile`.
    """
    if not url or not url.strip():
        raise ValueError("sp_list requires a non-empty url")
    if limit <= 0:
        raise ValueError(f"limit must be positive, got {limit}")

    hostname, site_path, folder_path = _parse_sharepoint_url(url)

    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_response = client.get(
            f"{GRAPH_BASE}/sites/{hostname}:{site_path}",
            headers=headers,
        )
        site_response.raise_for_status()
        site_id = site_response.json()["id"]

        if folder_path:
            children_url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{folder_path}:/children"
        else:
            children_url = f"{GRAPH_BASE}/sites/{site_id}/drive/root/children"

        children_response = client.get(
            children_url,
            params={"$top": limit},
            headers=headers,
        )
        children_response.raise_for_status()
        return _extract_items(children_response.json())
    finally:
        if http is None:
            client.close()


def _extract_items(payload: dict[str, Any]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for item in payload.get("value", []):
        out.append(
            {
                "name": item.get("name"),
                "type": "folder" if "folder" in item else "file",
                "size": item.get("size"),
                "last_modified": item.get("lastModifiedDateTime"),
                "web_url": item.get("webUrl"),
            },
        )
    return out
