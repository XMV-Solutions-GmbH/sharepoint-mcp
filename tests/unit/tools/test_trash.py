# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_trash_list / sp_trash_restore (#50)."""

from __future__ import annotations

import time
from collections.abc import Iterator

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.trash import (
    GRAPH_BETA_BASE,
    _extract_trash_items,
    _one_trash_item,
    trash_list,
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
ITEM_ID = "rb-item-1"

SITE_URL = f"https://{SITE_HOST}{SITE_PATH}"


def _mock_site_lookup() -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(
        json={"id": SITE_ID},
    )


# ---------------------------------------------------------------------
# trash_list()
# ---------------------------------------------------------------------


@respx.mock
def test_trash_list_returns_normalised_items(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BETA_BASE}/sites/{SITE_ID}/recycleBin/items").respond(
        json={
            "value": [
                {
                    "id": "rb1",
                    "name": "lost.docx",
                    "size": 4096,
                    "deletedDateTime": "2026-05-01T08:00:00Z",
                    "deletedFromLocation": "Shared Documents/policies",
                    "deletedBy": {"user": {"displayName": "Alice", "email": "a@x"}},
                },
            ],
        },
    )
    [out] = trash_list(SITE_URL)
    assert out["id"] == "rb1"
    assert out["name"] == "lost.docx"
    assert out["size"] == 4096
    assert out["deleted_from_location"] == "Shared Documents/policies"
    assert out["deleted_by"] == "Alice"


@respx.mock
def test_trash_list_paginates_via_nextLink(store_with_fresh_token: None) -> None:
    """Multi-page recycle bin: follow @odata.nextLink until limit or done."""
    del store_with_fresh_token
    _mock_site_lookup()
    next_link = f"{GRAPH_BETA_BASE}/sites/{SITE_ID}/recycleBin-page2"
    page1 = respx.get(url__startswith=f"{GRAPH_BETA_BASE}/sites/{SITE_ID}/recycleBin/items")
    page1.respond(
        json={
            "value": [{"id": "p1-a", "name": "a"}],
            "@odata.nextLink": next_link,
        },
    )
    respx.get(next_link).respond(json={"value": [{"id": "p2-a", "name": "b"}]})
    out = trash_list(SITE_URL)
    assert [r["id"] for r in out] == ["p1-a", "p2-a"]


@respx.mock
def test_trash_list_respects_limit(store_with_fresh_token: None) -> None:
    """If the first page already has `limit` items, don't fetch the next page."""
    del store_with_fresh_token
    _mock_site_lookup()
    next_link = f"{GRAPH_BETA_BASE}/sites/{SITE_ID}/recycleBin-page2"
    page1 = respx.get(
        url__startswith=f"{GRAPH_BETA_BASE}/sites/{SITE_ID}/recycleBin/items",
    ).respond(
        json={
            "value": [{"id": "1"}, {"id": "2"}, {"id": "3"}],
            "@odata.nextLink": next_link,
        },
    )
    page2 = respx.get(next_link).respond(json={"value": [{"id": "4"}]})
    out = trash_list(SITE_URL, limit=2)
    assert [r["id"] for r in out] == ["1", "2"]
    assert page1.call_count == 1
    assert page2.call_count == 0


@respx.mock
def test_trash_list_orderby_desc_and_top_in_first_call(store_with_fresh_token: None) -> None:
    """First-call URL includes the sort + paging hint."""
    del store_with_fresh_token
    _mock_site_lookup()
    route = respx.get(f"{GRAPH_BETA_BASE}/sites/{SITE_ID}/recycleBin/items").respond(
        json={"value": []}
    )
    trash_list(SITE_URL)
    url = str(route.calls.last.request.url)
    assert "deletedDateTime" in url and "desc" in url
    assert "$top=200" in url or "%24top=200" in url


@respx.mock
def test_trash_list_empty(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BETA_BASE}/sites/{SITE_ID}/recycleBin/items").respond(
        json={"value": []},
    )
    assert trash_list(SITE_URL) == []


@respx.mock
def test_trash_list_propagates_403(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BETA_BASE}/sites/{SITE_ID}/recycleBin/items").respond(
        403, json={"error": {"code": "accessDenied"}}
    )
    with pytest.raises(httpx.HTTPStatusError):
        trash_list(SITE_URL)


def test_trash_list_rejects_zero_limit() -> None:
    with pytest.raises(ValueError, match="limit must be"):
        trash_list(SITE_URL, limit=0)


def test_trash_list_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        trash_list("")


def test_trash_list_rejects_file_url() -> None:
    with pytest.raises(ValueError, match="expects a site URL"):
        trash_list(f"{SITE_URL}/Shared Documents/x.docx")


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------


def test_extract_trash_items_handles_missing_value() -> None:
    assert _extract_trash_items({}) == []


def test_extract_trash_items_filters_non_dict_entries() -> None:
    [out] = _extract_trash_items({"value": [{"id": "a"}, "not-a-dict", None]})
    assert out["id"] == "a"


def test_one_trash_item_normalises_missing_fields() -> None:
    out = _one_trash_item({})
    assert out == {
        "id": "",
        "name": "",
        "size": 0,
        "deleted_date_time": "",
        "deleted_from_location": "",
        "deleted_by": "",
    }


def test_one_trash_item_handles_email_only_user() -> None:
    """deletedBy.user with email but no displayName."""
    out = _one_trash_item({"deletedBy": {"user": {"email": "alice@x.com"}}})
    assert out["deleted_by"] == "alice@x.com"


def test_one_trash_item_handles_non_dict_deletedBy() -> None:
    """Defensive: malformed payload doesn't crash."""
    out = _one_trash_item({"deletedBy": "alice"})
    assert out["deleted_by"] == ""
