# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_sites / sp_followed_sites (#49)."""

from __future__ import annotations

import time
from collections.abc import Iterator

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.sites import (
    _extract_sites,
    _one_drive,
    _one_site,
    drives,
    followed_sites,
    sites,
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
SITE_ID = "contoso.sharepoint.com,site-guid,web-guid"


# ---------------------------------------------------------------------
# sites()
# ---------------------------------------------------------------------


@respx.mock
def test_sites_returns_parsed_results(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    respx.get(f"{GRAPH_BASE}/sites").respond(
        json={
            "value": [
                {
                    "id": "id-1",
                    "displayName": "Finance",
                    "webUrl": "https://x/sites/finance",
                    "description": "Money things",
                    "lastModifiedDateTime": "2026-04-01T00:00:00Z",
                },
                {
                    "id": "id-2",
                    "name": "hr",
                    "webUrl": "https://x/sites/hr",
                },
            ]
        },
    )
    result = sites("budget")
    assert len(result) == 2
    assert result[0]["name"] == "Finance"
    assert result[0]["description"] == "Money things"
    assert result[0]["last_modified"] == "2026-04-01T00:00:00Z"
    assert result[1]["name"] == "hr"
    assert result[1]["description"] == ""


@respx.mock
def test_sites_default_query_uses_wildcard(store_with_fresh_token: None) -> None:
    """Empty/None query becomes search=* to satisfy Microsoft's required param."""
    del store_with_fresh_token
    route = respx.get(f"{GRAPH_BASE}/sites").respond(json={"value": []})
    sites()
    assert "search=%2A" in str(route.calls.last.request.url) or "search=*" in str(
        route.calls.last.request.url
    )


@respx.mock
def test_sites_passes_query_verbatim(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    route = respx.get(f"{GRAPH_BASE}/sites").respond(json={"value": []})
    sites("policy")
    url = str(route.calls.last.request.url)
    assert "search=policy" in url


@respx.mock
def test_sites_empty_value_returns_empty(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    respx.get(f"{GRAPH_BASE}/sites").respond(json={"value": []})
    assert sites() == []


@respx.mock
def test_sites_propagates_4xx(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    respx.get(f"{GRAPH_BASE}/sites").respond(403, json={"error": {"code": "accessDenied"}})
    with pytest.raises(httpx.HTTPStatusError):
        sites()


# ---------------------------------------------------------------------
# followed_sites()
# ---------------------------------------------------------------------


@respx.mock
def test_followed_sites_lists_user_followed(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    respx.get(f"{GRAPH_BASE}/me/followedSites").respond(
        json={
            "value": [
                {
                    "id": "id-followed-1",
                    "displayName": "Engineering",
                    "webUrl": "https://x/sites/engineering",
                },
            ]
        },
    )
    [out] = followed_sites()
    assert out["name"] == "Engineering"


@respx.mock
def test_followed_sites_400_translates_to_helpful_error(store_with_fresh_token: None) -> None:
    """Graph returns 400 for /me in app-only mode; surface a clear message."""
    del store_with_fresh_token
    respx.get(f"{GRAPH_BASE}/me/followedSites").respond(
        400, json={"error": {"code": "Request_BadRequest"}}
    )
    with pytest.raises(RuntimeError, match="service-principal"):
        followed_sites()


@respx.mock
def test_followed_sites_propagates_other_4xx(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    respx.get(f"{GRAPH_BASE}/me/followedSites").respond(
        403, json={"error": {"code": "accessDenied"}}
    )
    with pytest.raises(httpx.HTTPStatusError):
        followed_sites()


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------


def test_extract_sites_handles_missing_value() -> None:
    assert _extract_sites({}) == []


def test_extract_sites_filters_non_dict_entries() -> None:
    """Defensive: shouldn't crash on a malformed payload."""
    payload = {"value": [{"id": "a", "displayName": "A"}, "not-a-dict", None, 42]}
    [out] = _extract_sites(payload)
    assert out["id"] == "a"


def test_one_site_normalises_missing_fields() -> None:
    assert _one_site({}) == {
        "id": "",
        "name": "",
        "web_url": "",
        "description": "",
        "last_modified": "",
    }


def test_one_site_prefers_displayName_over_name() -> None:
    out = _one_site({"displayName": "Display", "name": "Internal"})
    assert out["name"] == "Display"


def test_one_site_falls_back_to_name_when_no_displayName() -> None:
    out = _one_site({"name": "Internal Only"})
    assert out["name"] == "Internal Only"


# ---------------------------------------------------------------------
# drives()
# ---------------------------------------------------------------------


@respx.mock
def test_drives_lists_libraries_on_site(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:/sites/foo").respond(json={"id": SITE_ID})
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drives").respond(
        json={
            "value": [
                {
                    "id": "drive-1",
                    "name": "Documents",
                    "webUrl": "https://x/sites/foo/Shared Documents",
                    "driveType": "documentLibrary",
                    "quota": {"total": 1000, "used": 50},
                },
                {
                    "id": "drive-2",
                    "name": "Site Assets",
                    "webUrl": "https://x/sites/foo/SiteAssets",
                    "driveType": "documentLibrary",
                },
            ],
        },
    )
    result = drives(f"https://{SITE_HOST}/sites/foo")
    assert len(result) == 2
    assert result[0]["name"] == "Documents"
    assert result[0]["drive_type"] == "documentLibrary"
    assert result[0]["quota_total"] == 1000
    assert result[1]["name"] == "Site Assets"
    assert result[1]["quota_total"] == 0  # missing quota ok


def test_drives_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        drives("")


def test_drives_rejects_file_url() -> None:
    with pytest.raises(ValueError, match="expects a site URL"):
        drives(f"https://{SITE_HOST}/sites/foo/Shared Documents/policy.docx")


def test_one_drive_normalises_missing_fields() -> None:
    out = _one_drive({})
    assert out["id"] == ""
    assert out["name"] == ""
    assert out["drive_type"] == ""
    assert out["quota_total"] == 0
    assert out["quota_used"] == 0


def test_one_drive_handles_non_dict_quota() -> None:
    """If Graph returns a non-dict quota field, don't crash."""
    out = _one_drive({"id": "x", "name": "y", "quota": None})
    assert out["quota_total"] == 0
    assert out["quota_used"] == 0
