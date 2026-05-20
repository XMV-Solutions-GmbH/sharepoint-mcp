# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_drive_file_delete (issue #92)."""

from __future__ import annotations

import time
from collections.abc import Iterator

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.delete_file import delete_file


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
DRIVE_ID = "drive-abc"
ITEM_ID = "item-xyz"
SITE_URL = f"https://{SITE_HOST}{SITE_PATH}"
FILE_PATH = "Documents/report.md"


def _mock_site() -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(json={"id": SITE_ID})


def _mock_resolve_item(path: str = FILE_PATH) -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/{path}").respond(
        json={
            "id": ITEM_ID,
            "parentReference": {"driveId": DRIVE_ID},
        }
    )


def _mock_delete() -> respx.Route:
    return respx.delete(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}").respond(204)


# ------------------------------------------------------------------
# Happy path
# ------------------------------------------------------------------


@respx.mock
def test_delete_file_returns_deleted_true(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_item()
    _mock_delete()

    result = delete_file(SITE_URL, FILE_PATH)

    assert result["deleted"] is True
    assert result["path"] == FILE_PATH


@respx.mock
def test_delete_file_sends_delete_to_correct_url(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_item()
    route = _mock_delete()

    delete_file(SITE_URL, FILE_PATH)

    assert route.call_count == 1
    assert f"/drives/{DRIVE_ID}/items/{ITEM_ID}" in str(route.calls.last.request.url)


@respx.mock
def test_delete_file_strips_leading_slash_from_path(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_item()
    _mock_delete()

    result = delete_file(SITE_URL, "/" + FILE_PATH)

    assert result["deleted"] is True
    assert result["path"] == FILE_PATH


@respx.mock
def test_delete_file_path_at_root_level(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    path = "root-file.txt"
    _mock_site()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/{path}").respond(
        json={"id": ITEM_ID, "parentReference": {"driveId": DRIVE_ID}}
    )
    _mock_delete()

    result = delete_file(SITE_URL, path)

    assert result["path"] == path


@respx.mock
def test_delete_file_accepts_bearer_token(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_item()
    route = _mock_delete()

    delete_file(SITE_URL, FILE_PATH)

    assert route.calls.last.request.headers["Authorization"] == "Bearer AT-test"


# ------------------------------------------------------------------
# Error paths
# ------------------------------------------------------------------


@respx.mock
def test_delete_file_propagates_404(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/{FILE_PATH}").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    # resolve_drive_item_full fallback 1: strips the "Documents" prefix and retries.
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/report.md").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    # resolve_drive_item_full fallback 2: looks up drives by name; empty list → re-raise.
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drives").respond(json={"value": []})

    with pytest.raises(httpx.HTTPStatusError) as exc_info:
        delete_file(SITE_URL, FILE_PATH)

    assert exc_info.value.response.status_code == 404


@respx.mock
def test_delete_file_propagates_403(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_item()
    respx.delete(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}").respond(
        403, json={"error": {"code": "accessDenied"}}
    )

    with pytest.raises(httpx.HTTPStatusError) as exc_info:
        delete_file(SITE_URL, FILE_PATH)

    assert exc_info.value.response.status_code == 403


def test_delete_file_rejects_empty_site_url() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        delete_file("", FILE_PATH)


def test_delete_file_rejects_empty_path() -> None:
    with pytest.raises(ValueError, match="non-empty path"):
        delete_file(SITE_URL, "")


def test_delete_file_rejects_whitespace_site_url() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        delete_file("   ", FILE_PATH)


def test_delete_file_rejects_whitespace_path() -> None:
    with pytest.raises(ValueError, match="non-empty path"):
        delete_file(SITE_URL, "  ")


# ------------------------------------------------------------------
# http injection
# ------------------------------------------------------------------


@respx.mock
def test_delete_file_reuses_injected_client(store_with_fresh_token: None) -> None:
    """Injected httpx.Client is not closed by delete_file."""
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_item()
    _mock_delete()

    with httpx.Client() as client:
        delete_file(SITE_URL, FILE_PATH, http=client)
        assert not client.is_closed
