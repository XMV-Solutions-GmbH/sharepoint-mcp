# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for the SharePoint Lists CRUD tools (#44)."""

from __future__ import annotations

import time
from collections.abc import Iterator

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.lists import (
    _identity_display_name,
    _one_column,
    _one_item,
    _one_list,
    create_item,
    delete_item,
    get_item,
    list_columns,
    list_items,
    lists,
    parse_list_url,
    update_item,
)


class _MemStore:
    def __init__(self, value: bytes | None) -> None:
        self._v = value

    def get(self, profile: str) -> bytes | None:
        return self._v

    def set(self, profile: str, value: bytes) -> None:
        self._v = value

    def delete(self, profile: str) -> None:
        self._v = None


@pytest.fixture
def store_with_fresh_token(monkeypatch: pytest.MonkeyPatch) -> Iterator[None]:
    cached = CachedToken(
        access_token="AT-test",
        refresh_token="RT-test",
        expires_at=time.time() + 3600,
        scope="",
    )
    monkeypatch.setattr(
        "sharepoint_mcp.auth.get_token_store",
        lambda: _MemStore(cached.to_json().encode()),
    )
    monkeypatch.delenv("SP_AUTH_MODE", raising=False)
    monkeypatch.delenv("SP_CLIENT_SECRET", raising=False)
    yield


SITE_HOST = "contoso.sharepoint.com"
SITE_PATH = "/sites/foo"
SITE_ID = "contoso.sharepoint.com,site,web"
LIST_NAME = "Issues"
LIST_URL = f"https://{SITE_HOST}{SITE_PATH}/Lists/{LIST_NAME}"
SITE_URL = f"https://{SITE_HOST}{SITE_PATH}"


def _mock_site_lookup() -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(
        json={"id": SITE_ID},
    )


# ---------------------------------------------------------------------
# parse_list_url
# ---------------------------------------------------------------------


def test_parse_list_url_basic() -> None:
    assert parse_list_url(LIST_URL) == (SITE_HOST, SITE_PATH, "Issues")


def test_parse_list_url_url_decodes_name() -> None:
    url = f"https://{SITE_HOST}{SITE_PATH}/Lists/Issue%20Tracker"
    assert parse_list_url(url) == (SITE_HOST, SITE_PATH, "Issue Tracker")


def test_parse_list_url_case_insensitive_lists_segment() -> None:
    url = f"https://{SITE_HOST}{SITE_PATH}/lists/Issues"
    assert parse_list_url(url) == (SITE_HOST, SITE_PATH, "Issues")


def test_parse_list_url_works_with_teams_root() -> None:
    url = f"https://{SITE_HOST}/teams/sales/Lists/Leads"
    assert parse_list_url(url) == (SITE_HOST, "/teams/sales", "Leads")


def test_parse_list_url_rejects_empty() -> None:
    with pytest.raises(ValueError, match="non-empty"):
        parse_list_url("")


def test_parse_list_url_rejects_relative() -> None:
    with pytest.raises(ValueError, match="absolute URL"):
        parse_list_url("/sites/foo/Lists/Issues")


def test_parse_list_url_rejects_no_lists_segment() -> None:
    with pytest.raises(ValueError, match="/Lists/"):
        parse_list_url(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents")


def test_parse_list_url_rejects_lists_segment_without_name() -> None:
    with pytest.raises(ValueError, match="/Lists/"):
        parse_list_url(f"https://{SITE_HOST}{SITE_PATH}/Lists")


def test_parse_list_url_rejects_non_site_root() -> None:
    with pytest.raises(ValueError, match="must look like"):
        parse_list_url(f"https://{SITE_HOST}/personal/me/Lists/X")


# ---------------------------------------------------------------------
# lists()
# ---------------------------------------------------------------------


@respx.mock
def test_lists_returns_normalised(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/lists").respond(
        json={
            "value": [
                {
                    "id": "L1",
                    "name": "Issues",
                    "displayName": "Issue Tracker",
                    "webUrl": "https://x/sites/foo/Lists/Issues",
                    "description": "bug list",
                    "createdDateTime": "2026-01-01T00:00:00Z",
                    "lastModifiedDateTime": "2026-04-01T00:00:00Z",
                    "list": {"template": "genericList"},
                },
            ],
        },
    )
    [out] = lists(SITE_URL)
    assert out["id"] == "L1"
    assert out["display_name"] == "Issue Tracker"
    assert out["template"] == "genericList"


def test_lists_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        lists("")


def test_lists_rejects_file_url() -> None:
    with pytest.raises(ValueError, match="expects a site URL"):
        lists(LIST_URL)  # has /Lists/ — not a site


# ---------------------------------------------------------------------
# list_columns()
# ---------------------------------------------------------------------


@respx.mock
def test_list_columns_returns_typed_columns(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/lists/{LIST_NAME}/columns").respond(
        json={
            "value": [
                {
                    "id": "c1",
                    "displayName": "Title",
                    "name": "Title",
                    "required": True,
                    "text": {},
                },
                {
                    "id": "c2",
                    "displayName": "Status",
                    "name": "Status",
                    "choice": {"choices": ["Open", "Closed"]},
                },
                {
                    "id": "c3",
                    "displayName": "Priority",
                    "name": "Priority",
                    "number": {},
                },
                {
                    "id": "c4",
                    "displayName": "Assigned To",
                    "name": "AssignedTo",
                    "personOrGroup": {},
                },
                {
                    "id": "c5",
                    "displayName": "Mystery",
                    "name": "Mystery",
                },
            ],
        },
    )
    [title, status, priority, assigned, mystery] = list_columns(LIST_URL)
    assert title["type"] == "text"
    assert title["required"] is True
    assert status["type"] == "choice"
    assert priority["type"] == "number"
    assert assigned["type"] == "person"
    assert mystery["type"] == ""  # unknown facet


# ---------------------------------------------------------------------
# list_items()
# ---------------------------------------------------------------------


@respx.mock
def test_list_items_passes_filter_and_top(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    route = respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/lists/{LIST_NAME}/items").respond(
        json={"value": []},
    )
    list_items(LIST_URL, filter="fields/Status eq 'Open'", top=42)
    url = str(route.calls.last.request.url)
    assert "Status+eq+%27Open%27" in url or "Status eq 'Open'" in url
    assert "$top=42" in url or "%24top=42" in url
    assert "$expand=fields" in url or "%24expand=fields" in url


@respx.mock
def test_list_items_returns_normalised_items(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/lists/{LIST_NAME}/items").respond(
        json={
            "value": [
                {
                    "id": "1",
                    "createdDateTime": "2026-04-01T08:00:00Z",
                    "lastModifiedDateTime": "2026-04-02T09:00:00Z",
                    "createdBy": {"user": {"displayName": "Alice"}},
                    "lastModifiedBy": {"user": {"displayName": "Bob"}},
                    "webUrl": "https://x/sites/foo/Lists/Issues/EditForm.aspx?ID=1",
                    "fields": {"Title": "First bug", "Status": "Open"},
                },
            ],
        },
    )
    [out] = list_items(LIST_URL)
    assert out["id"] == "1"
    assert out["fields"]["Title"] == "First bug"
    assert out["created_by"] == "Alice"
    assert out["last_modified_by"] == "Bob"


def test_list_items_rejects_zero_top() -> None:
    with pytest.raises(ValueError, match="top must be"):
        list_items(LIST_URL, top=0)


# ---------------------------------------------------------------------
# get_item / create / update / delete
# ---------------------------------------------------------------------


@respx.mock
def test_get_item_fetches_with_expand(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    route = respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/lists/{LIST_NAME}/items/42").respond(
        json={"id": "42", "fields": {"Title": "T"}},
    )
    out = get_item(LIST_URL, "42")
    assert out["id"] == "42"
    assert out["fields"]["Title"] == "T"
    assert "$expand=fields" in str(route.calls.last.request.url) or "%24expand=fields" in str(
        route.calls.last.request.url
    )


def test_get_item_rejects_empty_item_id() -> None:
    with pytest.raises(ValueError, match="non-empty item_id"):
        get_item(LIST_URL, "")


@respx.mock
def test_create_item_posts_fields(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    route = respx.post(f"{GRAPH_BASE}/sites/{SITE_ID}/lists/{LIST_NAME}/items").respond(
        json={"id": "99", "fields": {"Title": "New"}},
    )
    out = create_item(LIST_URL, {"Title": "New"})
    assert out["id"] == "99"
    body = route.calls.last.request.read().decode()
    assert "Title" in body and "New" in body


def test_create_item_rejects_empty_fields() -> None:
    with pytest.raises(ValueError, match="non-empty fields"):
        create_item(LIST_URL, {})


@respx.mock
def test_update_item_patches_fields_endpoint(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    route = respx.patch(f"{GRAPH_BASE}/sites/{SITE_ID}/lists/{LIST_NAME}/items/42/fields").respond(
        json={"Title": "Updated", "Status": "Closed"}
    )
    out = update_item(LIST_URL, "42", {"Status": "Closed"})
    assert out["Status"] == "Closed"
    body = route.calls.last.request.read().decode()
    assert "Status" in body and "Closed" in body


def test_update_item_rejects_empty_item_id() -> None:
    with pytest.raises(ValueError, match="non-empty item_id"):
        update_item(LIST_URL, "", {"X": "Y"})


def test_update_item_rejects_empty_fields() -> None:
    with pytest.raises(ValueError, match="non-empty fields"):
        update_item(LIST_URL, "42", {})


@respx.mock
def test_delete_item_does_not_raise(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    route = respx.delete(f"{GRAPH_BASE}/sites/{SITE_ID}/lists/{LIST_NAME}/items/42").respond(204)
    delete_item(LIST_URL, "42")
    assert route.call_count == 1


def test_delete_item_rejects_empty_item_id() -> None:
    with pytest.raises(ValueError, match="non-empty item_id"):
        delete_item(LIST_URL, "")


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------


def test_one_list_normalises_missing_fields() -> None:
    out = _one_list({})
    assert out["id"] == ""
    assert out["template"] == ""


def test_one_list_handles_non_dict_list_facet() -> None:
    out = _one_list({"id": "X", "list": "not-a-dict"})
    assert out["template"] == ""


def test_one_column_picks_first_matching_facet() -> None:
    out = _one_column({"id": "x", "text": {}, "choice": {}})
    assert out["type"] == "text"  # text comes first in our priority list


def test_one_item_handles_missing_fields_facet() -> None:
    out = _one_item({"id": "X"})
    assert out["fields"] == {}


def test_identity_display_name_uses_email_fallback() -> None:
    assert _identity_display_name({"user": {"email": "alice@x"}}) == "alice@x"


def test_identity_display_name_handles_empty() -> None:
    assert _identity_display_name(None) == ""
    assert _identity_display_name({}) == ""
    assert _identity_display_name({"user": "not-a-dict"}) == ""


@respx.mock
def test_list_items_skips_filter_when_blank(store_with_fresh_token: None) -> None:
    """Empty/None filter must NOT add a $filter param (Graph rejects empty)."""
    del store_with_fresh_token
    _mock_site_lookup()
    route = respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/lists/{LIST_NAME}/items").respond(
        json={"value": []}
    )
    list_items(LIST_URL, filter="")
    url = str(route.calls.last.request.url)
    assert "$filter" not in url and "%24filter" not in url
