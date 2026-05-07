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
    _restored_item_summary,
    trash_list,
    trash_restore,
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


def test_trash_list_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        trash_list("")


def test_trash_list_rejects_file_url() -> None:
    with pytest.raises(ValueError, match="expects a site URL"):
        trash_list(f"{SITE_URL}/Shared Documents/x.docx")


# ---------------------------------------------------------------------
# trash_restore()
# ---------------------------------------------------------------------


@respx.mock
def test_trash_restore_returns_restored_item_summary_on_200(
    store_with_fresh_token: None,
) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.post(f"{GRAPH_BETA_BASE}/sites/{SITE_ID}/recycleBin/items/{ITEM_ID}/restore").respond(
        json={"id": "back-1", "name": "lost.docx", "webUrl": "https://x/lost.docx"},
    )
    result = trash_restore(SITE_URL, ITEM_ID)
    assert result == {
        "id": "back-1",
        "name": "lost.docx",
        "web_url": "https://x/lost.docx",
    }


@respx.mock
def test_trash_restore_returns_empty_on_204(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.post(f"{GRAPH_BETA_BASE}/sites/{SITE_ID}/recycleBin/items/{ITEM_ID}/restore").respond(204)
    assert trash_restore(SITE_URL, ITEM_ID) == {}


@respx.mock
def test_trash_restore_propagates_404(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.post(f"{GRAPH_BETA_BASE}/sites/{SITE_ID}/recycleBin/items/{ITEM_ID}/restore").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    with pytest.raises(httpx.HTTPStatusError):
        trash_restore(SITE_URL, ITEM_ID)


def test_trash_restore_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        trash_restore("", ITEM_ID)


def test_trash_restore_rejects_empty_item_id() -> None:
    with pytest.raises(ValueError, match="non-empty item_id"):
        trash_restore(SITE_URL, "")


def test_trash_restore_rejects_file_url() -> None:
    with pytest.raises(ValueError, match="expects a site URL"):
        trash_restore(f"{SITE_URL}/Shared Documents/x.docx", ITEM_ID)


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


def test_restored_item_summary_normalises_missing_fields() -> None:
    assert _restored_item_summary({}) == {"id": "", "name": "", "web_url": ""}
