# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_search — search SharePoint document libraries via Microsoft Graph search.

Wraps `POST /search/query` for `entityTypes: ["driveItem"]`. Read-only,
idempotent. Filter parameters (site / folder / file_type /
modified_after) translate into KQL fragments appended to the query
string — Microsoft's search service does the heavy lifting, we just
shape the request and unpack the response.
"""

from __future__ import annotations

from typing import Any

import httpx

from sharepoint_mcp.auth import get_token

GRAPH_SEARCH_URL = "https://graph.microsoft.com/v1.0/search/query"


def search(
    query: str,
    *,
    site: str | None = None,
    folder: str | None = None,
    file_type: str | None = None,
    modified_after: str | None = None,
    limit: int = 25,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> list[dict[str, Any]]:
    """Search SharePoint document libraries the signed-in user can see.

    Returns at most `limit` hits, each a dict with `name`, `path`,
    `web_url`, `last_modified`, `author`, `size`. An empty list is a
    valid result (no matches).

    Filter args translate to KQL fragments appended to `query`:

    - `site="https://contoso.sharepoint.com/sites/foo"` → `site:"<url>"`
    - `folder="/Shared Documents/policies"` → `path:"<folder>"`
    - `file_type="docx"` → `fileExtension:docx`
    - `modified_after="2024-01-01"` → `lastModifiedDateTime>=2024-01-01`

    Raises `httpx.HTTPStatusError` on a non-2xx response from Graph,
    `sharepoint_mcp.auth.AuthRequiredError` if no usable cached token
    exists for `profile`.
    """
    if not query or not query.strip():
        raise ValueError("sp_search requires a non-empty query")
    if limit <= 0:
        raise ValueError(f"limit must be positive, got {limit}")

    parts = [query]
    if site:
        parts.append(f'site:"{site}"')
    if folder:
        parts.append(f'path:"{folder}"')
    if file_type:
        parts.append(f"fileExtension:{file_type}")
    if modified_after:
        parts.append(f"lastModifiedDateTime>={modified_after}")

    body = {
        "requests": [
            {
                "entityTypes": ["driveItem"],
                "query": {"queryString": " ".join(parts)},
                "from": 0,
                "size": limit,
            },
        ],
    }

    token = get_token(profile)
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        response = client.post(
            GRAPH_SEARCH_URL,
            json=body,
            headers={"Authorization": f"Bearer {token}"},
        )
        response.raise_for_status()
        return _extract_hits(response.json())
    finally:
        if http is None:
            client.close()


def _extract_hits(payload: dict[str, Any]) -> list[dict[str, Any]]:
    """Convert Graph's nested `/search/query` response into flat hit dicts."""
    containers = payload.get("value", [])
    if not containers:
        return []
    hits_containers = containers[0].get("hitsContainers", [])
    if not hits_containers:
        return []

    out: list[dict[str, Any]] = []
    for hit in hits_containers[0].get("hits", []):
        resource = hit.get("resource") or {}
        out.append(
            {
                "name": resource.get("name"),
                "path": _extract_path(resource),
                "web_url": resource.get("webUrl"),
                "last_modified": resource.get("lastModifiedDateTime"),
                "author": _extract_user(resource.get("lastModifiedBy")),
                "size": resource.get("size"),
            },
        )
    return out


def _extract_user(modified_by: dict[str, Any] | None) -> str | None:
    if not modified_by:
        return None
    user = modified_by.get("user")
    if not isinstance(user, dict):
        return None
    return user.get("displayName") or user.get("email")


def _extract_path(resource: dict[str, Any]) -> str | None:
    """Microsoft returns parentReference.path like '/sites/.../root:/folder/sub'.

    We strip everything up to and including 'root:' to get a
    site-relative path like '/folder/sub', then append the file name.
    """
    parent = resource.get("parentReference") or {}
    path = parent.get("path") or ""
    name = resource.get("name") or ""
    if "root:" in path:
        path = path.split("root:", 1)[1]
    elif ":" in path:
        path = path.split(":", 1)[1]
    if path:
        return f"{path.rstrip('/')}/{name}" if name else path
    return name or None
