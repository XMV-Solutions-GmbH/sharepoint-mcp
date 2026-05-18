# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_file_metadata — read/write custom SharePoint column values on a drive file.

Wraps `GET /drives/{id}/items/{id}/listItem/fields` (read path) and
`PATCH /drives/{id}/items/{id}/listItem/fields` (write path).

Every file in a SharePoint document library is backed by a list item
that can carry arbitrary custom columns defined on the library.
Standard system fields (Author, Modified, etc.) are also present in
the fields facet.  This tool surfaces that facet directly so agents
can read retention labels, department tags, classification values, or
any other metadata without having to go through the Lists API.

When `fields` is omitted (read mode): one Graph call per invocation
(site lookup + item lookup + fields GET = 3 total).
When `fields` is provided (write mode): one extra PATCH; returns the
server-confirmed field state after the update.
"""

from __future__ import annotations

from typing import Any

import httpx

from sharepoint_mcp.auth import get_token
from sharepoint_mcp.tools._common import (
    GRAPH_BASE,
    parse_sharepoint_url,
    resolve_drive_item,
    resolve_site_id,
)


def file_metadata(
    url: str,
    fields: dict[str, Any] | None = None,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> dict[str, Any]:
    """Read (or update) the custom SharePoint column values attached to a file.

    `url` is the human-readable URL of a file inside a SharePoint
    document library.

    When `fields` is ``None`` (the default): performs a
    ``GET .../listItem/fields`` and returns the full field dict.

    When `fields` is a non-empty dict: performs a
    ``PATCH .../listItem/fields`` with those field updates and returns
    the updated field state as confirmed by Graph.  Only the keys
    present in `fields` are updated; unmentioned columns are
    unchanged.  Use internal column names (e.g. ``"Department"``,
    ``"_Status"``) — the same keys you get back from the read path.

    Returns:
        A flat dict of internal column-name → value covering all
        columns Graph exposes on this library item (system fields
        included).

    Raises:
        ValueError: `url` is empty/relative, points at a site or
            folder, or `fields` is provided but is not a dict.
        httpx.HTTPStatusError: Graph returned non-2xx.
        sharepoint_mcp.auth.AuthRequiredError: no usable cached token.
    """
    if not url or not url.strip():
        raise ValueError("sp_file_metadata requires a non-empty url")

    if fields is not None and not isinstance(fields, dict):
        raise TypeError("sp_file_metadata: fields must be a dict or None")

    hostname, site_path, item_path = parse_sharepoint_url(url)
    if not item_path:
        raise ValueError(
            f"sp_file_metadata needs a file URL, got a site/folder URL: {url!r}",
        )

    token = get_token(profile)
    auth_header = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=auth_header)
        drive_id, item_id = resolve_drive_item(client, site_id, item_path, headers=auth_header)

        fields_url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/listItem/fields"

        if fields is not None:
            patch = client.patch(
                fields_url,
                headers={**auth_header, "Content-Type": "application/json"},
                json=fields,
            )
            patch.raise_for_status()
            return dict(patch.json())

        get = client.get(fields_url, headers=auth_header)
        get.raise_for_status()
        return dict(get.json())
    finally:
        if http is None:
            client.close()
