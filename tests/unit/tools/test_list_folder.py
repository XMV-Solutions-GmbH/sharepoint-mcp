# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_list."""

from __future__ import annotations

import time
from collections.abc import Iterator

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE, parse_sharepoint_url
from sharepoint_mcp.tools.list_folder import _extract_items, list_folder


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
    fake = _MemStore(cached.to_json().encode())
    monkeypatch.setattr("sharepoint_mcp.auth.get_token_store", lambda: fake)
    yield


# ---------------------------------------------------------------------
# parse_sharepoint_url
# ---------------------------------------------------------------------


def test_parse_site_url_only() -> None:
    assert parse_sharepoint_url("https://contoso.sharepoint.com/sites/foo") == (
        "contoso.sharepoint.com",
        "/sites/foo",
        "",
    )


def test_parse_site_url_with_default_library() -> None:
    """Library segment 'Shared Documents' is stripped."""
    assert parse_sharepoint_url(
        "https://contoso.sharepoint.com/sites/foo/Shared Documents",
    ) == ("contoso.sharepoint.com", "/sites/foo", "")


def test_parse_site_url_with_default_library_url_encoded() -> None:
    """%20-encoded 'Shared Documents' is also stripped."""
    assert parse_sharepoint_url(
        "https://contoso.sharepoint.com/sites/foo/Shared%20Documents",
    ) == ("contoso.sharepoint.com", "/sites/foo", "")


def test_parse_folder_under_default_library() -> None:
    assert parse_sharepoint_url(
        "https://contoso.sharepoint.com/sites/foo/Shared Documents/policies",
    ) == ("contoso.sharepoint.com", "/sites/foo", "policies")


def test_parse_nested_folder() -> None:
    assert parse_sharepoint_url(
        "https://contoso.sharepoint.com/sites/foo/Shared Documents/policies/iso27001",
    ) == ("contoso.sharepoint.com", "/sites/foo", "policies/iso27001")


def test_parse_teams_url() -> None:
    """Teams sites use /teams/<name> instead of /sites/<name>."""
    assert parse_sharepoint_url(
        "https://contoso.sharepoint.com/teams/eng",
    ) == ("contoso.sharepoint.com", "/teams/eng", "")


def test_parse_rejects_relative_url() -> None:
    with pytest.raises(ValueError, match="absolute URL"):
        parse_sharepoint_url("/sites/foo")  # raises from _common, not list_folder


# ---------------------------------------------------------------------
# list_folder — happy path against mocked Graph (2 calls: site + children)
# ---------------------------------------------------------------------


SITE_HOST = "contoso.sharepoint.com"
SITE_PATH = "/sites/foo"
SITE_ID = "contoso.sharepoint.com,site-guid,web-guid"


def _mock_site_lookup() -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(
        json={"id": SITE_ID, "displayName": "foo"}
    )


def _mock_share_lookup(
    name: str = "policies",
    drive_id: str = "DID",
    item_id: str = "FID",
) -> respx.Route:
    """Primary resolver for folder/file URLs: /shares/{u!base64}/driveItem."""
    import re

    return respx.get(re.compile(rf"{re.escape(GRAPH_BASE)}/shares/u!.*?/driveItem")).respond(
        json={
            "id": item_id,
            "name": name,
            "parentReference": {"driveId": drive_id},
        },
    )


@respx.mock
def test_list_folder_site_root(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/children").respond(
        json={
            "value": [
                {
                    "name": "policies",
                    "folder": {"childCount": 0},
                    "size": 0,
                    "lastModifiedDateTime": "2026-04-15T10:00:00Z",
                    "webUrl": f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/policies",
                },
                {
                    "name": "readme.md",
                    "file": {"mimeType": "text/markdown"},
                    "size": 256,
                    "lastModifiedDateTime": "2026-04-16T11:00:00Z",
                    "webUrl": f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/readme.md",
                },
            ],
        },
    )
    items = list_folder(f"https://{SITE_HOST}{SITE_PATH}")
    assert [i["name"] for i in items] == ["policies", "readme.md"]
    assert items[0]["type"] == "folder"
    assert items[1]["type"] == "file"


@respx.mock
def test_list_folder_subfolder(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    # Resolve subfolder driveItem first, then list /drives/.../children
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/policies").respond(
        json={"id": "FID", "name": "policies", "parentReference": {"driveId": "DID"}},
    )
    respx.get(f"{GRAPH_BASE}/drives/DID/items/FID/children").respond(
        json={
            "value": [
                {
                    "name": "iso27001.docx",
                    "file": {},
                    "size": 99,
                    "lastModifiedDateTime": "2026-04-15T10:00:00Z",
                    "webUrl": "https://x/policies/iso27001.docx",
                },
            ],
        },
    )
    items = list_folder(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/policies")
    assert items[0]["name"] == "iso27001.docx"


@respx.mock
def test_list_folder_localized_library_name(store_with_fresh_token: None) -> None:
    """Regression test for #79: German-tenant URL with 'Freigegebene
    Dokumente'. The resolver's first fallback strips the localized
    library segment and retries against the default drive."""
    del store_with_fresh_token
    _mock_site_lookup()
    # Primary 404 with the German library in the path.
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/Freigegebene Dokumente/Finanzen").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    # First-fallback retry: strip "Freigegebene Dokumente", try default drive.
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/Finanzen").respond(
        json={"id": "FID-de", "name": "Finanzen", "parentReference": {"driveId": "DID"}},
    )
    respx.get(f"{GRAPH_BASE}/drives/DID/items/FID-de/children").respond(
        json={
            "value": [
                {
                    "name": "steuer.pdf",
                    "file": {},
                    "size": 12345,
                    "lastModifiedDateTime": "2026-04-15T10:00:00Z",
                    "webUrl": "https://x/Freigegebene Dokumente/Finanzen/steuer.pdf",
                },
            ],
        },
    )
    items = list_folder(
        f"https://{SITE_HOST}{SITE_PATH}/Freigegebene Dokumente/Finanzen",
    )
    assert items[0]["name"] == "steuer.pdf"


@respx.mock
def test_list_folder_empty(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/children").respond(
        json={"value": []},
    )
    assert list_folder(f"https://{SITE_HOST}{SITE_PATH}") == []


@respx.mock
def test_list_folder_sends_bearer_and_top(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    site_route = _mock_site_lookup()
    children_route = respx.get(
        f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/children",
    ).respond(json={"value": []})

    list_folder(f"https://{SITE_HOST}{SITE_PATH}", limit=42)

    for call in (site_route.calls.last, children_route.calls.last):
        assert call.request.headers.get("authorization") == "Bearer AT-test"
    children_url = str(children_route.calls.last.request.url)
    assert "%24top=42" in children_url or "$top=42" in children_url


@respx.mock
def test_list_folder_propagates_404_on_site_lookup(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    with pytest.raises(httpx.HTTPStatusError):
        list_folder(f"https://{SITE_HOST}{SITE_PATH}")


@respx.mock
def test_list_folder_propagates_404_on_folder(store_with_fresh_token: None) -> None:
    """Primary 404 → strip-first-segment retry 404 → library-search empty
    → original 404 propagates."""
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/missing").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drives").respond(json={"value": []})
    with pytest.raises(httpx.HTTPStatusError):
        list_folder(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/missing")


# ---------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------


def test_list_folder_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        list_folder("")


def test_list_folder_rejects_blank_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        list_folder("   ")


def test_list_folder_rejects_non_positive_limit() -> None:
    with pytest.raises(ValueError, match="limit must be positive"):
        list_folder("https://example.com/sites/foo", limit=0)


# ---------------------------------------------------------------------
# _extract_items
# ---------------------------------------------------------------------


def test_extract_items_folder_marker() -> None:
    payload = {"value": [{"name": "x", "folder": {}}]}
    assert _extract_items(payload)[0]["type"] == "folder"


def test_extract_items_file_marker() -> None:
    payload = {"value": [{"name": "x", "file": {}}]}
    assert _extract_items(payload)[0]["type"] == "file"


def test_extract_items_no_marker_defaults_file() -> None:
    payload = {"value": [{"name": "x"}]}
    assert _extract_items(payload)[0]["type"] == "file"


def test_extract_items_handles_missing_value() -> None:
    assert _extract_items({}) == []
