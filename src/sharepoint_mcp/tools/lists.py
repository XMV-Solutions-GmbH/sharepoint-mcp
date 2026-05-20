# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""SharePoint Lists CRUD (closes #44).

Reads:

- `sp_list_list(site_url)` — list all lists on a site
- `sp_list_column_list(list_url)` — schema of a list (column definitions)
- `sp_list_item_list(list_url, filter=None, top=100)` — items with optional OData filter
- `sp_list_item_get(list_url, item_id)` — single item with all fields

Writes (gated by SP_ALLOW_WRITES at the server layer):

- `sp_list_item_create(list_url, fields)` — create a list item
- `sp_list_item_update(list_url, item_id, fields)` — patch fields
- `sp_list_item_delete(list_url, item_id)` — delete

URL conventions:

- `site_url` follows the same shape every other tool uses
  (`https://<tenant>/sites/<name>` or `/teams/<name>`).
- `list_url` is the SharePoint web URL of the list — a site URL with
  `/Lists/<list-name>` appended (`https://contoso.sharepoint.com/sites/foo/Lists/Issues`).

Microsoft Graph's `GET /sites/{site-id}/lists/{list-id-or-name}`
accepts the list's display name as the `{list-id-or-name}` segment,
so we can resolve a list URL to a list-id without an extra lookup
when the agent already has the name in the URL.

Item shape:

    {
        "id": "<list-item-id>",
        "created_date_time": "<ISO datetime>",
        "last_modified_date_time": "<ISO datetime>",
        "created_by": "<display name or empty>",
        "last_modified_by": "<display name or empty>",
        "web_url": "<browser URL>",
        "fields": {<column-name>: <value>, ...},
    }

`fields` is the SharePoint listItem's `fields` facet — a flat dict
keyed by internal column names. Read tools include the full fields
expansion; write tools accept any subset that matches the list's
schema.
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


def parse_list_url(url: str) -> tuple[str, str, str]:
    """Split a SharePoint list URL into (hostname, site_path, list_name).

    Recognises the canonical `/sites/<site>/Lists/<list-name>` shape
    used by SharePoint web URLs, with case-insensitive matching on the
    `Lists` segment to tolerate different SharePoint locales.

    Examples:
        https://contoso.sharepoint.com/sites/foo/Lists/Issues
            -> ("contoso.sharepoint.com", "/sites/foo", "Issues")

        https://contoso.sharepoint.com/sites/foo/Lists/Issue%20Tracker
            -> ("contoso.sharepoint.com", "/sites/foo", "Issue Tracker")

    Raises:
        ValueError: URL is not absolute, doesn't have a /Lists/<name>
            segment, or the list-name segment is empty.
    """
    if not url or not url.strip():
        raise ValueError("list URL must be non-empty")
    parsed = urlparse(url)
    if not parsed.netloc:
        raise ValueError(f"expected an absolute URL, got {url!r}")
    parts = [p for p in unquote(parsed.path).split("/") if p]
    if len(parts) < 4 or parts[0] not in {"sites", "teams"}:
        raise ValueError(
            f"list URL must look like https://<host>/sites/<name>/Lists/<list>; got {url!r}",
        )
    site_path = f"/{parts[0]}/{parts[1]}"
    # Find the "Lists" segment (case-insensitive) in the rest of the path.
    rest = parts[2:]
    list_idx = next((i for i, p in enumerate(rest) if p.lower() == "lists"), -1)
    if list_idx < 0 or list_idx + 1 >= len(rest):
        raise ValueError(
            f"list URL must include a /Lists/<list-name> segment; got {url!r}",
        )
    list_name = rest[list_idx + 1]
    if not list_name:
        raise ValueError(f"list URL has empty list-name segment: {url!r}")
    return parsed.netloc, site_path, list_name


# ---------------------------------------------------------------------
# Read tools
# ---------------------------------------------------------------------


def lists(
    site_url: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> list[dict[str, Any]]:
    """List all lists on a SharePoint site.

    Returns each list with id, name, display_name, web_url,
    description, created_date_time, last_modified_date_time, and
    template (e.g. "documentLibrary", "genericList", "tasks").
    """
    if not site_url or not site_url.strip():
        raise ValueError("sp_list_list requires a non-empty site_url")
    hostname, site_path, item_path = parse_sharepoint_url(site_url)
    if item_path:
        raise ValueError(
            f"sp_list_list expects a site URL, not a file/folder URL "
            f"(got {site_url!r}; item path {item_path!r}).",
        )

    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        response = client.get(f"{GRAPH_BASE}/sites/{site_id}/lists", headers=headers)
        response.raise_for_status()
        return _extract_lists(response.json())
    finally:
        if http is None:
            client.close()


def list_columns(
    list_url: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> list[dict[str, Any]]:
    """Return the column definitions (schema) of a SharePoint list.

    Each column: id, display_name, name (internal), column_group,
    description, required, hidden, read_only, indexed, type
    (best-effort based on which type-specific facet is populated:
    "text", "choice", "number", "boolean", "datetime", "person",
    "lookup", "calculated", or "" if unknown).
    """
    hostname, site_path, list_name = parse_list_url(list_url)
    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        response = client.get(
            f"{GRAPH_BASE}/sites/{site_id}/lists/{list_name}/columns",
            headers=headers,
        )
        response.raise_for_status()
        return _extract_columns(response.json())
    finally:
        if http is None:
            client.close()


def list_items(
    list_url: str,
    *,
    filter: str | None = None,
    top: int = 100,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> list[dict[str, Any]]:
    """List items in a SharePoint list, with optional OData $filter.

    `filter` is passed through verbatim as the Graph `$filter` query
    parameter — applies to the items' `fields` facet, e.g.
    `"fields/Status eq 'Open'"`. Empty/None disables filtering.

    `top` caps results (default 100, Microsoft caps at 5000 per page).
    The first page is returned without following @odata.nextLink —
    callers wanting more should issue subsequent calls with stricter
    filters.
    """
    if top < 1:
        raise ValueError(f"top must be >= 1, got {top!r}")
    hostname, site_path, list_name = parse_list_url(list_url)
    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        params: dict[str, str] = {"$expand": "fields", "$top": str(top)}
        if filter and filter.strip():
            params["$filter"] = filter
        response = client.get(
            f"{GRAPH_BASE}/sites/{site_id}/lists/{list_name}/items",
            headers=headers,
            params=params,
        )
        response.raise_for_status()
        return _extract_items(response.json())
    finally:
        if http is None:
            client.close()


def get_item(
    list_url: str,
    item_id: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> dict[str, Any]:
    """Fetch a single list item with all expanded fields."""
    if not item_id or not str(item_id).strip():
        raise ValueError("sp_list_item_get requires a non-empty item_id")
    hostname, site_path, list_name = parse_list_url(list_url)
    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        response = client.get(
            f"{GRAPH_BASE}/sites/{site_id}/lists/{list_name}/items/{item_id}",
            headers=headers,
            params={"$expand": "fields"},
        )
        response.raise_for_status()
        return _one_item(response.json())
    finally:
        if http is None:
            client.close()


# ---------------------------------------------------------------------
# Write tools
# ---------------------------------------------------------------------


def create_item(
    list_url: str,
    fields: dict[str, Any],
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> dict[str, Any]:
    """Create a new list item with the given fields. Returns the new item."""
    if not isinstance(fields, dict) or not fields:
        raise ValueError("sp_list_item_create requires a non-empty fields dict")
    hostname, site_path, list_name = parse_list_url(list_url)
    token = get_token(profile)
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        response = client.post(
            f"{GRAPH_BASE}/sites/{site_id}/lists/{list_name}/items",
            headers=headers,
            json={"fields": fields},
        )
        response.raise_for_status()
        return _one_item(response.json())
    finally:
        if http is None:
            client.close()


def update_item(
    list_url: str,
    item_id: str,
    fields: dict[str, Any],
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> dict[str, Any]:
    """Patch an existing list item's fields. Returns the updated fields.

    Microsoft Graph's PATCH on `/items/{id}/fields` accepts a flat
    dict and merges into the existing fields facet — fields not in
    the dict are unchanged.
    """
    if not item_id or not str(item_id).strip():
        raise ValueError("sp_list_item_update requires a non-empty item_id")
    if not isinstance(fields, dict) or not fields:
        raise ValueError("sp_list_item_update requires a non-empty fields dict")
    hostname, site_path, list_name = parse_list_url(list_url)
    token = get_token(profile)
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        response = client.patch(
            f"{GRAPH_BASE}/sites/{site_id}/lists/{list_name}/items/{item_id}/fields",
            headers=headers,
            json=fields,
        )
        response.raise_for_status()
        # The PATCH response is the fields dict itself, not a full
        # listItem. Return it as-is for the caller.
        payload = response.json() if response.content else {}
        return dict(payload) if isinstance(payload, dict) else {}
    finally:
        if http is None:
            client.close()


def delete_item(
    list_url: str,
    item_id: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> None:
    """Delete a list item. Sends it to the site recycle bin (per
    SharePoint's default behaviour for DELETE)."""
    if not item_id or not str(item_id).strip():
        raise ValueError("sp_list_item_delete requires a non-empty item_id")
    hostname, site_path, list_name = parse_list_url(list_url)
    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        response = client.delete(
            f"{GRAPH_BASE}/sites/{site_id}/lists/{list_name}/items/{item_id}",
            headers=headers,
        )
        response.raise_for_status()
    finally:
        if http is None:
            client.close()


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------


def _extract_lists(payload: dict[str, Any]) -> list[dict[str, Any]]:
    raw = payload.get("value", []) or []
    if not isinstance(raw, list):
        return []
    return [_one_list(entry) for entry in raw if isinstance(entry, dict)]


def _one_list(entry: dict[str, Any]) -> dict[str, Any]:
    info = entry.get("list") or {}
    template = (info.get("template") or "") if isinstance(info, dict) else ""
    return {
        "id": str(entry.get("id") or ""),
        "name": str(entry.get("name") or ""),
        "display_name": str(entry.get("displayName") or ""),
        "web_url": str(entry.get("webUrl") or ""),
        "description": str(entry.get("description") or ""),
        "created_date_time": str(entry.get("createdDateTime") or ""),
        "last_modified_date_time": str(entry.get("lastModifiedDateTime") or ""),
        "template": str(template),
    }


_COLUMN_TYPE_FACETS = (
    ("text", "text"),
    ("choice", "choice"),
    ("number", "number"),
    ("boolean", "boolean"),
    ("dateTime", "datetime"),
    ("personOrGroup", "person"),
    ("lookup", "lookup"),
    ("calculated", "calculated"),
    ("hyperlinkOrPicture", "hyperlink"),
    ("currency", "currency"),
)


def _extract_columns(payload: dict[str, Any]) -> list[dict[str, Any]]:
    raw = payload.get("value", []) or []
    if not isinstance(raw, list):
        return []
    return [_one_column(entry) for entry in raw if isinstance(entry, dict)]


def _one_column(entry: dict[str, Any]) -> dict[str, Any]:
    column_type = ""
    for facet, label in _COLUMN_TYPE_FACETS:
        if facet in entry:
            column_type = label
            break
    return {
        "id": str(entry.get("id") or ""),
        "display_name": str(entry.get("displayName") or ""),
        "name": str(entry.get("name") or ""),
        "column_group": str(entry.get("columnGroup") or ""),
        "description": str(entry.get("description") or ""),
        "required": bool(entry.get("required") or False),
        "hidden": bool(entry.get("hidden") or False),
        "read_only": bool(entry.get("readOnly") or False),
        "indexed": bool(entry.get("indexed") or False),
        "type": column_type,
    }


def _extract_items(payload: dict[str, Any]) -> list[dict[str, Any]]:
    raw = payload.get("value", []) or []
    if not isinstance(raw, list):
        return []
    return [_one_item(entry) for entry in raw if isinstance(entry, dict)]


def _one_item(entry: dict[str, Any]) -> dict[str, Any]:
    fields = entry.get("fields")
    if not isinstance(fields, dict):
        fields = {}
    return {
        "id": str(entry.get("id") or ""),
        "created_date_time": str(entry.get("createdDateTime") or ""),
        "last_modified_date_time": str(entry.get("lastModifiedDateTime") or ""),
        "created_by": _identity_display_name(entry.get("createdBy")),
        "last_modified_by": _identity_display_name(entry.get("lastModifiedBy")),
        "web_url": str(entry.get("webUrl") or ""),
        "fields": dict(fields),
    }


def _identity_display_name(raw: Any) -> str:
    if not isinstance(raw, dict):
        return ""
    user = raw.get("user") if isinstance(raw, dict) else None
    if isinstance(user, dict):
        return str(user.get("displayName") or user.get("email") or "")
    return ""
