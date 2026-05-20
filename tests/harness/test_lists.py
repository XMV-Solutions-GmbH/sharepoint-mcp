# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for the SharePoint Lists CRUD tools (#44).

For listing tools (sp_list_list, sp_list_column_list, sp_list_item_list): we
exercise against the harness site directly. The default Documents
library is exposed via `/sites/{id}/lists` (Graph treats it as a
list with template="documentLibrary"), so every site has at least
one entry.

For mutating tools (create / update / delete): we create a fresh
custom list via raw Graph at fixture setup, run the lifecycle, then
delete the list at teardown. This avoids polluting the harness
sandbox between runs.
"""

from __future__ import annotations

import uuid
from collections.abc import Iterator

import httpx
import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.tools._common import GRAPH_BASE, parse_sharepoint_url, resolve_site_id
from sharepoint_mcp.tools.lists import (
    create_item,
    delete_item,
    get_item,
    list_columns,
    list_items,
    lists,
    update_item,
)

HARNESS_PROFILE = "harness"
HARNESS_SITE_URL = "https://xmvsolutions.sharepoint.com/sites/sharepoint-mcp-harness"


def _skip_if_no_harness() -> None:
    try:
        get_token(HARNESS_PROFILE)
    except AuthRequiredError as exc:
        pytest.skip(
            f"Harness credentials not available: {exc}. "
            "Run `uv run mcp-server-sharepoint login --profile harness` to populate.",
        )


# ---------------------------------------------------------------------
# Read-only against the existing harness site
# ---------------------------------------------------------------------


def test_lists_returns_at_least_one_list() -> None:
    _skip_if_no_harness()
    out = lists(HARNESS_SITE_URL, profile=HARNESS_PROFILE)
    assert isinstance(out, list)
    assert len(out) >= 1, "every site has at least one list (Documents)"
    for entry in out:
        assert "id" in entry
        assert "name" in entry
        assert "web_url" in entry


def test_lists_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        lists("", profile=HARNESS_PROFILE)


# ---------------------------------------------------------------------
# Lifecycle test: create a fresh custom list, exercise CRUD, delete
# ---------------------------------------------------------------------


@pytest.fixture
def temporary_list() -> Iterator[str]:
    """Create a fresh custom list via raw Graph; yield its URL; delete it on teardown."""
    _skip_if_no_harness()
    list_name = f"harness-crud-{uuid.uuid4().hex[:8]}"
    token = get_token(HARNESS_PROFILE)
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    list_id: str | None = None
    site_id: str | None = None
    with httpx.Client(timeout=30.0) as client:
        hostname, site_path, _ = parse_sharepoint_url(HARNESS_SITE_URL)
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)
        # Create the list with a couple of columns
        create_response = client.post(
            f"{GRAPH_BASE}/sites/{site_id}/lists",
            headers=headers,
            json={
                "displayName": list_name,
                "columns": [
                    {"name": "Title", "text": {}},
                    {"name": "Note", "text": {}},
                ],
                "list": {"template": "genericList"},
            },
        )
        if create_response.status_code != 201:
            pytest.skip(
                f"Could not create harness list (HTTP {create_response.status_code}): "
                f"{create_response.text[:200]}",
            )
        list_id = create_response.json()["id"]

    # Provide the list URL the tools accept
    list_url = f"{HARNESS_SITE_URL}/Lists/{list_name}"
    try:
        yield list_url
    finally:
        if site_id and list_id:
            with httpx.Client(timeout=30.0) as client:
                try:
                    client.delete(
                        f"{GRAPH_BASE}/sites/{site_id}/lists/{list_id}",
                        headers={"Authorization": f"Bearer {token}"},
                    )
                except httpx.HTTPError:
                    pass


def test_list_lifecycle_create_get_update_delete(temporary_list: str) -> None:
    list_url = temporary_list

    # Schema reflects the columns we just created (Title is built-in)
    columns = list_columns(list_url, profile=HARNESS_PROFILE)
    column_names = {c["name"] for c in columns}
    assert "Title" in column_names

    # CREATE
    created = create_item(
        list_url,
        {"Title": "First entry", "Note": "harness-test"},
        profile=HARNESS_PROFILE,
    )
    item_id = created["id"]
    assert item_id

    # GET
    fetched = get_item(list_url, item_id, profile=HARNESS_PROFILE)
    assert fetched["id"] == item_id
    assert fetched["fields"].get("Title") == "First entry"

    # LIST — newly-created item should be there
    all_items = list_items(list_url, profile=HARNESS_PROFILE, top=100)
    assert any(it["id"] == item_id for it in all_items)

    # UPDATE
    updated_fields = update_item(
        list_url,
        item_id,
        {"Title": "Updated"},
        profile=HARNESS_PROFILE,
    )
    assert updated_fields.get("Title") == "Updated"

    # GET again — confirms the patch landed
    fetched_again = get_item(list_url, item_id, profile=HARNESS_PROFILE)
    assert fetched_again["fields"].get("Title") == "Updated"

    # DELETE
    delete_item(list_url, item_id, profile=HARNESS_PROFILE)
    after_delete = list_items(list_url, profile=HARNESS_PROFILE, top=100)
    assert not any(it["id"] == item_id for it in after_delete)
