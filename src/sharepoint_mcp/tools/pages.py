# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""SharePoint modern Pages (closes #45).

Two read-only tools for the canonical wiki/knowledge-base format on
SharePoint Online (modern site pages — distinct from documents in
libraries):

- `sp_pages_list(site_url)` — list pages on a site
- `sp_page_read(page_url)` — fetch a page including canvas layout

URL convention: page URLs follow SharePoint's canonical form
`https://<host>/sites/<name>/SitePages/<filename>.aspx`. We parse
the trailing `.aspx` filename and look the page up via Graph's
`?$filter=name eq '<filename>'` query (Graph's `/sites/{id}/pages/
{id}` requires a GUID, not the name).

Canvas layout shape: Microsoft Graph exposes the page's web-parts
via `canvasLayout` (sections > columns > webParts). `sp_page_read`
returns the canvas as raw JSON for lossless inspection. Writing
pages back is intentionally out of scope: a metadata-only write
would be a half-tool (read full content + canvas, write only title)
that misleads agents into reaching for it expecting full edits, and
canvas writes need a clearer agent UX before they're safe. Today
modern Pages have to be edited via the SharePoint web UI.

Item shape (sp_pages_list, sp_page_read):

    {
        "id": "<sitePage GUID>",
        "name": "<filename>.aspx",
        "title": "<page title>",
        "web_url": "<browser URL>",
        "description": "<page description>",
        "page_layout": "<article|home|...>",
        "thumbnail_web_url": "<image URL or empty>",
        "last_modified": "<ISO datetime>",
        "last_modified_by": "<display name or empty>",
        # sp_page_read only:
        "canvas_layout": {<raw Graph canvasLayout JSON>},
    }
"""

from __future__ import annotations

from typing import Any
from urllib.parse import unquote, urlparse

import httpx

from sharepoint_mcp.auth import get_token
from sharepoint_mcp.tools._common import (
    GRAPH_BASE,
    parse_sharepoint_url,
    resolve_site_id,
)

PAGE_LAYOUT_NAME_LOWER = "sitepages"


def parse_page_url(url: str) -> tuple[str, str, str]:
    """Split a SharePoint page URL into (hostname, site_path, page_name).

    Recognises the canonical `/sites/<site>/SitePages/<name>.aspx`
    shape, case-insensitive on the SitePages segment to match
    SharePoint's URL casing tolerance.

    Examples:
        https://contoso.sharepoint.com/sites/foo/SitePages/Onboarding.aspx
            -> ("contoso.sharepoint.com", "/sites/foo", "Onboarding.aspx")

        https://contoso.sharepoint.com/sites/foo/sitepages/My%20Page.aspx
            -> ("contoso.sharepoint.com", "/sites/foo", "My Page.aspx")

    Raises:
        ValueError: URL is not absolute, doesn't have a /SitePages/<name>.aspx
            segment, or the page name is empty.
    """
    if not url or not url.strip():
        raise ValueError("page URL must be non-empty")
    parsed = urlparse(url)
    if not parsed.netloc:
        raise ValueError(f"expected an absolute URL, got {url!r}")
    parts = [p for p in unquote(parsed.path).split("/") if p]
    if len(parts) < 4 or parts[0] not in {"sites", "teams"}:
        raise ValueError(
            f"page URL must look like https://<host>/sites/<name>/SitePages/<page>.aspx; "
            f"got {url!r}",
        )
    site_path = f"/{parts[0]}/{parts[1]}"
    rest = parts[2:]
    sp_idx = next((i for i, p in enumerate(rest) if p.lower() == PAGE_LAYOUT_NAME_LOWER), -1)
    if sp_idx < 0 or sp_idx + 1 >= len(rest):
        raise ValueError(
            f"page URL must include a /SitePages/<name>.aspx segment; got {url!r}",
        )
    page_name = rest[sp_idx + 1]
    if not page_name:
        raise ValueError(f"page URL has empty page-name segment: {url!r}")
    return parsed.netloc, site_path, page_name


# ---------------------------------------------------------------------
# Read tools
# ---------------------------------------------------------------------


def pages_list(
    site_url: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> list[dict[str, Any]]:
    """List all modern pages on a SharePoint site.

    Returns each page with id, name, title, web_url, description,
    page_layout, thumbnail_web_url, last_modified, last_modified_by.
    """
    if not site_url or not site_url.strip():
        raise ValueError("sp_pages_list requires a non-empty site_url")
    hostname, site_path, item_path = parse_sharepoint_url(site_url)
    if item_path:
        raise ValueError(
            f"sp_pages_list expects a site URL, not a file/folder URL "
            f"(got {site_url!r}; item path {item_path!r}).",
        )

    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        response = client.get(f"{GRAPH_BASE}/sites/{site_id}/pages", headers=headers)
        response.raise_for_status()
        return _extract_pages(response.json(), include_canvas=False)
    finally:
        if http is None:
            client.close()


def page_read(
    page_url: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> dict[str, Any]:
    """Fetch a single SharePoint page including its canvas layout.

    Resolves the page URL to a page id via `?$filter=name eq '<name>'`
    (the Graph endpoint that takes a name segment requires a GUID,
    so we filter the list). Then expands canvasLayout for lossless
    web-part inspection.

    Returns the same shape as `sp_pages_list` plus `canvas_layout`
    (the raw Graph JSON for sections / columns / web-parts).
    """
    hostname, site_path, page_name = parse_page_url(page_url)
    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        page_id = _resolve_page_id(client, site_id, page_name, headers=headers)
        response = client.get(
            f"{GRAPH_BASE}/sites/{site_id}/pages/{page_id}/microsoft.graph.sitePage",
            headers=headers,
            params={"$expand": "canvasLayout"},
        )
        response.raise_for_status()
        return _one_page(response.json(), include_canvas=True)
    finally:
        if http is None:
            client.close()


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------


class PageNotFoundError(RuntimeError):
    """Raised when sp_page_read can't find a page by name.

    Different from a 404 — the list query succeeds but the response
    contains no matching page. Surfaced as a distinct exception so
    callers can handle "wrong page name" separately from "site
    doesn't exist" / "permission denied".
    """


def _resolve_page_id(
    client: httpx.Client,
    site_id: str,
    page_name: str,
    *,
    headers: dict[str, str],
) -> str:
    """Look up a page's GUID id by its filename via $filter."""
    # Microsoft's OData $filter expects single-quoted strings; embedded
    # apostrophes in the page name need escaping as ''.
    escaped = page_name.replace("'", "''")
    response = client.get(
        f"{GRAPH_BASE}/sites/{site_id}/pages",
        headers=headers,
        params={"$filter": f"name eq '{escaped}'"},
    )
    response.raise_for_status()
    payload = response.json()
    matches = payload.get("value") or []
    if not isinstance(matches, list) or not matches:
        raise PageNotFoundError(
            f"No SharePoint page named {page_name!r} on site {site_id!r}.",
        )
    first = matches[0]
    if not isinstance(first, dict) or not first.get("id"):
        raise PageNotFoundError(
            f"Page {page_name!r} found but the response is missing the id field.",
        )
    return str(first["id"])


def _extract_pages(payload: dict[str, Any], *, include_canvas: bool) -> list[dict[str, Any]]:
    raw = payload.get("value", []) or []
    if not isinstance(raw, list):
        return []
    return [
        _one_page(entry, include_canvas=include_canvas) for entry in raw if isinstance(entry, dict)
    ]


def _one_page(entry: dict[str, Any], *, include_canvas: bool) -> dict[str, Any]:
    last_modified_by = ""
    raw_lm = entry.get("lastModifiedBy")
    if isinstance(raw_lm, dict):
        user = raw_lm.get("user")
        if isinstance(user, dict):
            last_modified_by = str(user.get("displayName") or user.get("email") or "")

    out: dict[str, Any] = {
        "id": str(entry.get("id") or ""),
        "name": str(entry.get("name") or ""),
        "title": str(entry.get("title") or ""),
        "web_url": str(entry.get("webUrl") or ""),
        "description": str(entry.get("description") or ""),
        "page_layout": str(entry.get("pageLayout") or ""),
        "thumbnail_web_url": str(entry.get("thumbnailWebUrl") or ""),
        "last_modified": str(entry.get("lastModifiedDateTime") or ""),
        "last_modified_by": last_modified_by,
    }
    if include_canvas:
        canvas = entry.get("canvasLayout")
        out["canvas_layout"] = canvas if isinstance(canvas, dict) else {}
    return out
