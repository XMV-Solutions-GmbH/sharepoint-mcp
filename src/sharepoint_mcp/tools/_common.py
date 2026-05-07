# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Shared helpers across `sp_*` tool modules.

URL parsing + site-ID resolution. Kept here (rather than in
`tools/__init__.py`) so the package's public API doesn't accidentally
expose internals.

Naming uses leading underscore at the module level to flag "internal
to the tools/ subpackage".
"""

from __future__ import annotations

from urllib.parse import unquote, urlparse

import httpx

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# SharePoint URLs that point at the default drive ("Shared Documents")
# may include the library name as a path segment in any of these
# common spellings. We strip it before passing the rest to Graph as
# a folder path within the drive.
_DEFAULT_LIBRARY_SEGMENTS = frozenset(
    {"Shared Documents", "Shared%20Documents", "Documents"},
)
_SITE_ROOT_SEGMENTS = frozenset({"sites", "teams", "personal"})


def parse_sharepoint_url(url: str) -> tuple[str, str, str]:
    """Split a SharePoint URL into (hostname, site_path, item_path).

    Examples (after URL-decoding):

        https://contoso.sharepoint.com/sites/foo
            → ("contoso.sharepoint.com", "/sites/foo", "")

        https://contoso.sharepoint.com/sites/foo/Shared Documents
            → ("contoso.sharepoint.com", "/sites/foo", "")

        https://contoso.sharepoint.com/sites/foo/Shared Documents/policies
            → ("contoso.sharepoint.com", "/sites/foo", "policies")

        https://contoso.sharepoint.com/sites/foo/policies/iso.docx
            → ("contoso.sharepoint.com", "/sites/foo", "policies/iso.docx")

    `item_path` is drive-relative (NOT site-relative), with no
    leading slash. Empty `item_path` means "root of the default drive".
    """
    parsed = urlparse(url)
    if not parsed.netloc:
        raise ValueError(f"expected an absolute URL, got {url!r}")
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

    item_path = "/".join(rest)
    return parsed.netloc, site_path, item_path


def resolve_site_id(
    client: httpx.Client,
    hostname: str,
    site_path: str,
    *,
    headers: dict[str, str],
) -> str:
    """Resolve a SharePoint site URL to its Graph site-id.

    Wraps `GET /sites/{hostname}:{site_path}`. Raises
    `httpx.HTTPStatusError` on non-2xx (e.g. 404 if the site doesn't
    exist or the user can't see it).
    """
    response = client.get(
        f"{GRAPH_BASE}/sites/{hostname}:{site_path}",
        headers=headers,
    )
    response.raise_for_status()
    return str(response.json()["id"])
