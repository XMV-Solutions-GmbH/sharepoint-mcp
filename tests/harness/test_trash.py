# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_trash_list / sp_trash_restore (#50).

Lifecycle: publish a throwaway file -> delete it via raw Graph ->
list trash and assert it's there -> restore it -> read tool confirms
it's back -> clean up.
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
from sharepoint_mcp.tools.trash import trash_list, trash_restore

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
    """Publish a unique file, delete it via raw Graph, yield (filename, url).

    On teardown: best-effort delete again in case the test re-restored the
    file but didn't clean up.
    """
    _skip_if_no_harness()
    run_id = uuid.uuid4().hex[:8]
    name = f"trash-harness-{run_id}.txt"
    local = tmp_path / name
    local.write_text(f"trash harness {run_id}\n", encoding="utf-8")
    url = f"{HARNESS_FOLDER_URL}/{name}"

    publish(str(local), HARNESS_FOLDER_URL, name=name, profile=HARNESS_PROFILE)
    _delete_file_to_recycle_bin(url)

    try:
        yield name, url
    finally:
        # If the test left the file un-restored, nothing to clean up.
        # If the test restored the file, delete it back to recycle bin.
        try:
            _delete_file_to_recycle_bin(url)
        except Exception:
            pass


def test_trash_list_finds_recently_deleted_file(
    published_then_deleted: tuple[str, str],
) -> None:
    name, _ = published_then_deleted
    items = trash_list(HARNESS_SITE_URL, profile=HARNESS_PROFILE)
    assert isinstance(items, list)
    matching = [it for it in items if it.get("name") == name]
    assert matching, f"deleted file {name!r} not found in recycle bin: {items!r}"
    assert matching[0]["id"]
    assert matching[0]["deleted_date_time"]


def test_trash_restore_brings_file_back(
    published_then_deleted: tuple[str, str],
) -> None:
    name, _ = published_then_deleted
    items = trash_list(HARNESS_SITE_URL, profile=HARNESS_PROFILE)
    matching = [it for it in items if it.get("name") == name]
    if not matching:
        pytest.skip(f"file {name!r} disappeared from recycle bin between tests")
    item_id = matching[0]["id"]
    trash_restore(HARNESS_SITE_URL, item_id, profile=HARNESS_PROFILE)
    # Confirm it's no longer in trash
    items_after = trash_list(HARNESS_SITE_URL, profile=HARNESS_PROFILE)
    still_there = [it for it in items_after if it.get("id") == item_id]
    assert not still_there, "item still in recycle bin after restore"


def test_trash_list_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        trash_list("", profile=HARNESS_PROFILE)


def test_trash_restore_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        trash_restore("", "id", profile=HARNESS_PROFILE)
    with pytest.raises(ValueError, match="non-empty item_id"):
        trash_restore(HARNESS_SITE_URL, "", profile=HARNESS_PROFILE)


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
