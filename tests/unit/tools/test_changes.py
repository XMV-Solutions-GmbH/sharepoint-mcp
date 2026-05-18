# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_file_changes (#51) — delta-query change tracking.

Edge cases covered:
- First-call (no cursor) and subsequent-call (with cursor) paths
- Pagination via @odata.nextLink chained until @odata.deltaLink
- 410 Gone (stale cursor) propagates clearly
- Item type discrimination: file vs folder vs deleted
- Deleted items have name=='' and deleted=True; size 0; we don't crash
- parentReference missing / non-dict / present-with-path
- Empty result + non-list value (defensive)
- Cursor is preserved verbatim across calls (treated as opaque)
"""

from __future__ import annotations

import time
from collections.abc import Iterator

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.changes import _extract_items, _one_item, changes


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
SITE_URL = f"https://{SITE_HOST}{SITE_PATH}"


def _mock_site_lookup() -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(
        json={"id": SITE_ID},
    )


# ---------------------------------------------------------------------
# First-call (no cursor)
# ---------------------------------------------------------------------


@respx.mock
def test_first_call_uses_root_delta_endpoint(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    delta_link = "https://graph.microsoft.com/v1.0/.../delta?token=NEW"
    route = respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/delta").respond(
        json={
            "value": [
                {
                    "id": "i1",
                    "name": "policy.docx",
                    "webUrl": "https://x/foo/policy.docx",
                    "size": 1024,
                    "lastModifiedDateTime": "2026-04-01T00:00:00Z",
                    "parentReference": {"path": "/drive/root:/policies"},
                },
                {
                    "id": "i2",
                    "name": "policies",
                    "folder": {"childCount": 3},
                    "webUrl": "https://x/foo/policies",
                    "parentReference": {"path": "/drive/root:"},
                },
            ],
            "@odata.deltaLink": delta_link,
        },
    )
    out = changes(SITE_URL)
    assert route.called
    assert out["cursor"] == delta_link
    assert len(out["items"]) == 2
    assert out["items"][0]["type"] == "file"
    assert out["items"][1]["type"] == "folder"
    assert out["items"][1]["size"] == 0  # folder has no size in our shape


# ---------------------------------------------------------------------
# Subsequent call (with cursor)
# ---------------------------------------------------------------------


@respx.mock
def test_subsequent_call_uses_cursor_url_directly(store_with_fresh_token: None) -> None:
    """The cursor IS the URL Graph wants for the next call — we pass it
    verbatim. Site lookup must NOT happen on the cursor path."""
    del store_with_fresh_token
    cursor = "https://graph.microsoft.com/v1.0/.../delta?token=PREVIOUS"
    new_cursor = "https://graph.microsoft.com/v1.0/.../delta?token=NEXT"
    cursor_route = respx.get(cursor).respond(
        json={
            "value": [
                {"id": "i3", "name": "new.txt", "webUrl": "https://x/new.txt"},
            ],
            "@odata.deltaLink": new_cursor,
        },
    )
    site_route = respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(
        json={"id": SITE_ID},
    )
    out = changes(SITE_URL, since=cursor)
    assert cursor_route.called
    assert not site_route.called  # cursor path skips site lookup
    assert out["cursor"] == new_cursor
    assert out["items"][0]["name"] == "new.txt"


# ---------------------------------------------------------------------
# Pagination via @odata.nextLink
# ---------------------------------------------------------------------


@respx.mock
def test_paginates_via_nextLink_until_deltaLink(store_with_fresh_token: None) -> None:
    """Multi-page delta: follow nextLink, accumulate items, stop on deltaLink."""
    del store_with_fresh_token
    _mock_site_lookup()
    page2 = "https://graph.microsoft.com/v1.0/.../delta-page2"
    final_cursor = "https://graph.microsoft.com/v1.0/.../delta?token=FINAL"
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/delta").respond(
        json={"value": [{"id": "p1-a"}], "@odata.nextLink": page2},
    )
    respx.get(page2).respond(
        json={"value": [{"id": "p2-a"}], "@odata.deltaLink": final_cursor},
    )
    out = changes(SITE_URL)
    assert [i["id"] for i in out["items"]] == ["p1-a", "p2-a"]
    assert out["cursor"] == final_cursor


@respx.mock
def test_paginates_three_pages(store_with_fresh_token: None) -> None:
    """Sanity-check that we don't stop after the first nextLink follow."""
    del store_with_fresh_token
    _mock_site_lookup()
    page2 = "https://x/page2"
    page3 = "https://x/page3"
    final = "https://x/delta?token=END"
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/delta").respond(
        json={"value": [{"id": "1"}], "@odata.nextLink": page2},
    )
    respx.get(page2).respond(json={"value": [{"id": "2"}], "@odata.nextLink": page3})
    respx.get(page3).respond(json={"value": [{"id": "3"}], "@odata.deltaLink": final})
    out = changes(SITE_URL)
    assert [i["id"] for i in out["items"]] == ["1", "2", "3"]
    assert out["cursor"] == final


# ---------------------------------------------------------------------
# Deleted items
# ---------------------------------------------------------------------


@respx.mock
def test_deleted_items_are_marked_with_type_deleted(store_with_fresh_token: None) -> None:
    """Graph emits a `deleted` facet for items removed since cursor; name is empty."""
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/delta").respond(
        json={
            "value": [
                {
                    "id": "i-gone",
                    "deleted": {"state": "deleted"},
                },
            ],
            "@odata.deltaLink": "https://x/cursor",
        },
    )
    [out_item] = changes(SITE_URL)["items"]
    assert out_item["type"] == "deleted"
    assert out_item["deleted"] is True
    assert out_item["name"] == ""
    assert out_item["size"] == 0


# ---------------------------------------------------------------------
# Error propagation
# ---------------------------------------------------------------------


@respx.mock
def test_410_gone_propagates_for_stale_cursor(store_with_fresh_token: None) -> None:
    """Graph returns 410 when a cursor has been rolled past its retention.
    We propagate so the caller can surface 'reset cursor' to the user."""
    del store_with_fresh_token
    cursor = "https://graph.microsoft.com/v1.0/.../delta?token=ANCIENT"
    respx.get(cursor).respond(410, json={"error": {"code": "resyncRequired"}})
    with pytest.raises(httpx.HTTPStatusError) as excinfo:
        changes(SITE_URL, since=cursor)
    assert excinfo.value.response.status_code == 410


@respx.mock
def test_403_propagates_on_first_call(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/delta").respond(
        403, json={"error": {"code": "accessDenied"}}
    )
    with pytest.raises(httpx.HTTPStatusError):
        changes(SITE_URL)


# ---------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------


def test_changes_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty scope_url"):
        changes("")


def test_changes_rejects_blank_url() -> None:
    with pytest.raises(ValueError, match="non-empty scope_url"):
        changes("   ")


def test_changes_rejects_file_url() -> None:
    """File/folder URLs aren't site-scoped delta — reject (folder-scoped
    delta is a deferred follow-up)."""
    with pytest.raises(ValueError, match="site URL"):
        changes(f"{SITE_URL}/Shared Documents/policy.docx")


# ---------------------------------------------------------------------
# Edge cases — empty / malformed payloads
# ---------------------------------------------------------------------


@respx.mock
def test_empty_value_list_returns_empty_items_with_cursor(
    store_with_fresh_token: None,
) -> None:
    """Newly-created site or 'no changes since cursor' returns no items
    but still a fresh cursor."""
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/delta").respond(
        json={"value": [], "@odata.deltaLink": "https://x/cursor"},
    )
    out = changes(SITE_URL)
    assert out["items"] == []
    assert out["cursor"] == "https://x/cursor"


@respx.mock
def test_response_with_no_links_at_all_still_returns_cursor_unchanged(
    store_with_fresh_token: None,
) -> None:
    """Defensive: Graph response missing both nextLink and deltaLink shouldn't
    loop forever — we treat missing-deltaLink-on-final-page as 'preserve any
    incoming cursor', and the loop terminates."""
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/delta").respond(
        json={"value": [{"id": "x"}]},  # no nextLink, no deltaLink
    )
    out = changes(SITE_URL)
    # No cursor provided initially, no deltaLink in response → empty cursor
    assert out["cursor"] == ""
    assert len(out["items"]) == 1


@respx.mock
def test_cursor_preserved_when_response_omits_deltaLink(
    store_with_fresh_token: None,
) -> None:
    """If we passed a cursor in and Graph's response is missing deltaLink
    (defensive), keep the original cursor so the next call retries from
    the same point rather than full-resync."""
    del store_with_fresh_token
    cursor = "https://graph.microsoft.com/.../old"
    respx.get(cursor).respond(json={"value": [{"id": "x"}]})  # no links
    out = changes(SITE_URL, since=cursor)
    assert out["cursor"] == cursor


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------


def test_extract_items_handles_missing_value() -> None:
    assert _extract_items({}) == []


def test_extract_items_handles_non_list_value() -> None:
    assert _extract_items({"value": "not-a-list"}) == []


def test_extract_items_filters_non_dict_entries() -> None:
    [out] = _extract_items({"value": [{"id": "a"}, "junk", None, 7]})
    assert out["id"] == "a"


def test_one_item_classifies_folder() -> None:
    out = _one_item({"id": "f", "name": "folder", "folder": {"childCount": 0}})
    assert out["type"] == "folder"


def test_one_item_classifies_file_when_no_folder_no_deleted() -> None:
    out = _one_item({"id": "x", "name": "x.txt", "size": 5})
    assert out["type"] == "file"


def test_one_item_classifies_deleted_takes_precedence_over_folder() -> None:
    """A folder that's been deleted: type='deleted' wins."""
    out = _one_item({"id": "x", "deleted": {"state": "deleted"}, "folder": {}})
    assert out["type"] == "deleted"
    assert out["deleted"] is True


def test_one_item_handles_non_dict_deleted_facet() -> None:
    """Defensive: deleted not as a dict — don't crash, treat as not-deleted."""
    out = _one_item({"id": "x", "deleted": "yes", "name": "x.txt"})
    assert out["type"] == "file"
    assert out["deleted"] is False


def test_one_item_handles_missing_parentReference() -> None:
    out = _one_item({"id": "x"})
    assert out["parent_path"] == ""


def test_one_item_handles_non_dict_parentReference() -> None:
    out = _one_item({"id": "x", "parentReference": "not-a-dict"})
    assert out["parent_path"] == ""


def test_one_item_extracts_parent_path() -> None:
    out = _one_item({"id": "x", "parentReference": {"path": "/drive/root:/policies"}})
    assert out["parent_path"] == "/drive/root:/policies"


def test_one_item_normalises_missing_size_to_zero() -> None:
    out = _one_item({"id": "x", "name": "x"})
    assert out["size"] == 0


def test_one_item_handles_non_int_size() -> None:
    """Microsoft sometimes returns size as a string in some endpoints."""
    out = _one_item({"id": "x", "size": "1024"})
    assert out["size"] == 1024
