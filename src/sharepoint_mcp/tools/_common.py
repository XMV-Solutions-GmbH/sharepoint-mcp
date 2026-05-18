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

import base64
from typing import Any
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


def resolve_drive_item(
    client: httpx.Client,
    site_id: str,
    item_path: str,
    *,
    headers: dict[str, str],
    allow_library_fallback: bool = True,
) -> tuple[str, str]:
    """Resolve a drive-relative path to (drive_id, item_id).

    Default behaviour: tries the site's default drive first
    (`/sites/{site_id}/drive/root:/{item_path}`). If that 404s and
    `allow_library_fallback=True` (the default), splits the path on
    its first segment, looks that segment up as a library/drive name
    on the site, and retries against that drive's root. This is what
    makes URLs into non-default libraries (Site Assets, custom
    document libraries, etc.) work transparently.

    Set `allow_library_fallback=False` to disable the fallback —
    useful when the caller already knows the path is in the default
    drive and wants the original 404 to propagate without an extra
    Graph round-trip.
    """
    item = resolve_drive_item_full(
        client,
        site_id,
        item_path,
        headers=headers,
        allow_library_fallback=allow_library_fallback,
    )
    drive_id = str(item["parentReference"]["driveId"])
    item_id = str(item["id"])
    return drive_id, item_id


def resolve_drive_item_full(
    client: httpx.Client,
    site_id: str,
    item_path: str,
    *,
    headers: dict[str, str],
    allow_library_fallback: bool = True,
) -> dict[str, Any]:
    """Resolve a drive-relative path and return the full Graph driveItem.

    Like `resolve_drive_item` but returns the parsed driveItem dict
    so callers can read `name`, `eTag`, `size`, etc. without a
    second round-trip.

    See `resolve_drive_item` for the library-fallback semantics.
    """
    primary_url = (
        f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{item_path}"
        if item_path
        else f"{GRAPH_BASE}/sites/{site_id}/drive/root"
    )
    primary = client.get(primary_url, headers=headers)
    if primary.status_code != 404 or not allow_library_fallback or not item_path:
        primary.raise_for_status()
        return dict(primary.json())

    library_segment, _, rest = item_path.partition("/")

    # Fallback 1 (#79): the first segment may be the LOCALIZED display
    # name of the site's default library — German `Freigegebene
    # Dokumente`, Italian `Documenti condivisi`, etc. Microsoft Graph
    # addresses items in the default drive without the library
    # segment, so stripping it and retrying against `/drive/root:/{rest}`
    # often succeeds without any further lookup. Cheap (one extra
    # GET); preserves existing English-default-tenant behaviour because
    # primary already resolved there.
    if rest:
        default_retry_url = f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{rest}"
        default_retry = client.get(default_retry_url, headers=headers)
        if default_retry.status_code != 404:
            default_retry.raise_for_status()
            return dict(default_retry.json())

    # Fallback 2: first segment may be a non-default library name
    # (e.g. `SiteAssets`, a custom library). List the site's drives,
    # find a match by display name, retry against that drive's root.
    drive_id = _find_drive_id_by_name(client, site_id, library_segment, headers=headers)
    if drive_id is None:
        # No such library — re-raise the original 404 for clarity.
        primary.raise_for_status()
        return dict(primary.json())  # unreachable; mypy

    fallback_url = (
        f"{GRAPH_BASE}/drives/{drive_id}/root:/{rest}"
        if rest
        else f"{GRAPH_BASE}/drives/{drive_id}/root"
    )
    fallback = client.get(fallback_url, headers=headers)
    fallback.raise_for_status()
    return dict(fallback.json())


def resolve_drive_item_by_share_url(
    client: httpx.Client,
    web_url: str,
    *,
    headers: dict[str, str],
) -> dict[str, Any]:
    """Resolve any SharePoint / OneDrive webUrl to a canonical driveItem.

    Uses Microsoft Graph's `/shares/{shareId}/driveItem` endpoint with
    a `u!`-prefixed base64url-encoded URL. This works for **any** URL
    the caller has read access to — including URLs with **localized
    library names** (`Freigegebene Dokumente` on German tenants,
    `Documenti condivisi` on Italian, etc.), nested folders, and
    custom (non-default) document libraries — in one round-trip,
    without needing a separate site-id lookup or library-name
    enumeration first.

    Returns the full driveItem dict. The two stable identifiers the
    caller usually wants are `item["parentReference"]["driveId"]`
    and `item["id"]` — subsequent operations against
    `/drives/{driveId}/items/{itemId}/...` work regardless of locale.

    Encoding rules (per Microsoft Graph docs): standard base64url
    (`urlsafe_b64encode`) of the UTF-8 URL bytes, strip trailing `=`
    padding, prefix with `u!`. The `/` and `+` characters in the
    base64 alphabet are not used (urlsafe variant uses `-` and `_`).

    Fixes [#79](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues/79):
    `sp_list_folder` / `sp_read_file` no longer 404 on localized library names.
    """
    encoded = base64.urlsafe_b64encode(web_url.encode("utf-8")).rstrip(b"=").decode("ascii")
    share_id = f"u!{encoded}"
    response = client.get(
        f"{GRAPH_BASE}/shares/{share_id}/driveItem",
        headers=headers,
    )
    response.raise_for_status()
    return dict(response.json())


def _find_drive_id_by_name(
    client: httpx.Client,
    site_id: str,
    name: str,
    *,
    headers: dict[str, str],
) -> str | None:
    """Return the drive id whose display name matches `name` (case-insensitive),
    or None if the site has no such library.

    One Graph round-trip per call. Bulk operations across many distinct
    library URLs in one process pay this cost per URL; for the common
    case of repeated access to the same library, callers can plumb
    their own cache, but for v0.3 we keep it stateless for simplicity.
    """
    response = client.get(f"{GRAPH_BASE}/sites/{site_id}/drives", headers=headers)
    response.raise_for_status()
    drives = response.json().get("value", []) or []
    needle = name.casefold()
    for drive in drives:
        candidate = (drive.get("name") or "").casefold()
        if candidate == needle:
            return str(drive["id"])
    return None


def list_site_drives(
    client: httpx.Client,
    site_id: str,
    *,
    headers: dict[str, str],
) -> list[dict[str, Any]]:
    """List all drives (libraries) on a site. Used by sp_drives."""
    response = client.get(f"{GRAPH_BASE}/sites/{site_id}/drives", headers=headers)
    response.raise_for_status()
    raw = response.json().get("value", [])
    if not isinstance(raw, list):
        return []
    return [dict(d) for d in raw if isinstance(d, dict)]
