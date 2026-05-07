# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sharing-link tools (#47).

Lifecycle: publish a throwaway file -> create org-scoped view link ->
list links (assert ours is there) -> revoke -> list (assert gone) ->
delete the seed file at teardown.

We deliberately use scope='organization' (not 'anonymous') in the
harness even though scope='anonymous' is the more dangerous code
path, because:
- Tenants commonly disable anonymous sharing tenant-wide; the test
  would 403 in those tenants and become noise.
- The unit tests cover the body construction for both scopes; what
  the harness verifies is the create/list/revoke loop.
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
from sharepoint_mcp.tools.sharing import share_create, share_list, share_revoke

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
def published_file(tmp_path: Path) -> Iterator[str]:
    """Publish a throwaway file; yield its URL; delete on teardown."""
    _skip_if_no_harness()
    name = f"share-harness-{uuid.uuid4().hex[:8]}.txt"
    local = tmp_path / name
    local.write_text("share harness seed\n", encoding="utf-8")
    url = f"{HARNESS_FOLDER_URL}/{name}"
    publish(str(local), HARNESS_FOLDER_URL, name=name, profile=HARNESS_PROFILE)
    try:
        yield url
    finally:
        _delete_via_graph(url)


def test_share_create_then_list_then_revoke(published_file: str) -> None:
    url = published_file

    # CREATE — conservative defaults
    created = share_create(url, profile=HARNESS_PROFILE)
    assert created["id"]
    assert created["web_url"]
    assert created["type"] == "view"
    assert created["scope"] == "organization"
    link_id = created["id"]

    # LIST — must include the link we just created
    listed = share_list(url, profile=HARNESS_PROFILE)
    matching = [it for it in listed if it["id"] == link_id]
    assert matching, f"created link {link_id!r} missing from list: {listed!r}"
    assert matching[0]["web_url"] == created["web_url"]

    # REVOKE
    share_revoke(url, link_id, profile=HARNESS_PROFILE)

    # LIST again — must be gone
    after = share_list(url, profile=HARNESS_PROFILE)
    assert not [it for it in after if it["id"] == link_id], (
        f"link {link_id!r} still present after revoke"
    )


def test_share_create_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        share_create("", profile=HARNESS_PROFILE)
    with pytest.raises(ValueError, match="type must be one of"):
        share_create(HARNESS_FOLDER_URL + "/x.txt", type="bogus", profile=HARNESS_PROFILE)
    with pytest.raises(ValueError, match="scope must be one of"):
        share_create(HARNESS_FOLDER_URL + "/x.txt", scope="public", profile=HARNESS_PROFILE)


def test_share_revoke_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        share_revoke("", "link-id", profile=HARNESS_PROFILE)
    with pytest.raises(ValueError, match="non-empty link_id"):
        share_revoke(HARNESS_FOLDER_URL + "/x.txt", "", profile=HARNESS_PROFILE)


def test_share_list_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        share_list("", profile=HARNESS_PROFILE)
    with pytest.raises(ValueError, match="file/folder URL"):
        share_list(HARNESS_SITE_URL, profile=HARNESS_PROFILE)


# ---------------------------------------------------------------------
# Helper: delete a file via raw Graph
# ---------------------------------------------------------------------


def _delete_via_graph(url: str) -> None:
    try:
        token = get_token(HARNESS_PROFILE)
    except AuthRequiredError:
        return
    headers = {"Authorization": f"Bearer {token}"}
    try:
        with httpx.Client(timeout=30.0) as client:
            hostname, site_path, item_path = parse_sharepoint_url(url)
            site_id = resolve_site_id(client, hostname, site_path, headers=headers)
            drive_id, item_id = resolve_drive_item(client, site_id, item_path, headers=headers)
            client.delete(
                f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}",
                headers=headers,
            )
    except (httpx.HTTPError, KeyError):
        pass
