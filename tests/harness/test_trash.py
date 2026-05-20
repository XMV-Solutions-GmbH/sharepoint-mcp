# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_site_trash_list (#50, partial).

`sp_trash_restore` is not implemented because Microsoft Graph's
beta site recycle-bin endpoint doesn't expose a restore action
(only the SharePoint Embedded fileStorageContainer recycleBin does).
We list-only here; restore stays in the v0.4 backlog until either
Microsoft adds the action or we implement a SharePoint REST API
fallback.
"""

from __future__ import annotations

import uuid
from collections.abc import Iterator
from pathlib import Path

import httpx
import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.tools._common import (
    GRAPH_BASE,
    parse_sharepoint_url,
    resolve_drive_item,
    resolve_site_id,
)
from sharepoint_mcp.tools.publish import publish
from sharepoint_mcp.tools.trash import trash_list

HARNESS_PROFILE = "harness"
HARNESS_SITE_URL = "https://xmvsolutions.sharepoint.com/sites/sharepoint-mcp-harness"
HARNESS_FOLDER_URL = f"{HARNESS_SITE_URL}/Shared Documents"


def _skip_if_no_harness() -> None:
    try:
        get_token(HARNESS_PROFILE)
    except AuthRequiredError as exc:
        pytest.skip(
            f"Harness credentials not available: {exc}. "
            "Run `uv run mcp-server-sharepoint login --profile harness` to populate.",
        )


@pytest.fixture
def published_then_deleted(tmp_path: Path) -> Iterator[tuple[str, str]]:
    """Publish a unique file, delete it via raw Graph (which sends it to
    the recycle bin), yield (filename, url)."""
    _skip_if_no_harness()
    run_id = uuid.uuid4().hex[:8]
    name = f"trash-harness-{run_id}.txt"
    local = tmp_path / name
    local.write_text(f"trash harness {run_id}\n", encoding="utf-8")
    url = f"{HARNESS_FOLDER_URL}/{name}"

    publish(str(local), HARNESS_FOLDER_URL, name=name, profile=HARNESS_PROFILE)
    _delete_file_to_recycle_bin(url)
    yield name, url


def test_trash_list_finds_recently_deleted_file(
    published_then_deleted: tuple[str, str],
) -> None:
    """With $orderby=deletedDateTime+desc, our just-deleted file is at the top."""
    name, _ = published_then_deleted
    items = trash_list(HARNESS_SITE_URL, profile=HARNESS_PROFILE, limit=50)
    assert isinstance(items, list)
    matching = [it for it in items if it.get("name") == name]
    assert matching, f"deleted file {name!r} not found in top-50 of recycle bin"
    assert matching[0]["id"]
    assert matching[0]["deleted_date_time"]


def test_trash_list_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        trash_list("", profile=HARNESS_PROFILE)


# ---------------------------------------------------------------------
# Helper: delete a file to the recycle bin via raw Graph
# ---------------------------------------------------------------------


def _delete_file_to_recycle_bin(url: str) -> None:
    """DELETE /drives/{id}/items/{id} sends the item to the recycle bin
    (SharePoint's default behaviour for DELETE on driveItem)."""
    try:
        token = get_token(HARNESS_PROFILE)
    except AuthRequiredError:
        return
    headers = {"Authorization": f"Bearer {token}"}
    with httpx.Client(timeout=30.0) as client:
        hostname, site_path, item_path = parse_sharepoint_url(url)
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        drive_id, item_id = resolve_drive_item(client, site_id, item_path, headers=headers)
        client.delete(
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}",
            headers=headers,
        )
