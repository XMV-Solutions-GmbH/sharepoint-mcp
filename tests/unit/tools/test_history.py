# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_drive_file_history."""

from __future__ import annotations

import time
from collections.abc import Iterator

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.history import _extract_user, _extract_versions, history


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
    yield


SITE_HOST = "contoso.sharepoint.com"
SITE_PATH = "/sites/foo"
SITE_ID = "contoso.sharepoint.com,site-guid,web-guid"
DRIVE_ID = "b!drive"
ITEM_ID = "01ITEM"


def _mock_lookups() -> None:
    respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(json={"id": SITE_ID})
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/policy.docx").respond(
        json={
            "id": ITEM_ID,
            "name": "policy.docx",
            "parentReference": {"driveId": DRIVE_ID},
        }
    )


# ---------------------------------------------------------------------
# Happy path
# ---------------------------------------------------------------------


@respx.mock
def test_history_returns_parsed_versions(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_lookups()
    respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/versions").respond(
        json={
            "value": [
                {
                    "id": "3.0",
                    "lastModifiedDateTime": "2026-05-07T12:00:00Z",
                    "lastModifiedBy": {"user": {"displayName": "Alice", "email": "a@x"}},
                    "size": 4567,
                },
                {
                    "id": "2.0",
                    "lastModifiedDateTime": "2026-05-01T09:30:00Z",
                    "lastModifiedBy": {"user": {"displayName": "Bob"}},
                    "size": 4500,
                },
            ]
        }
    )
    result = history(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/policy.docx", limit=10)
    assert len(result) == 2
    assert result[0]["id"] == "3.0"
    assert result[0]["last_modified_by"] == "Alice"
    assert result[0]["size"] == 4567
    assert result[1]["last_modified_by"] == "Bob"


@respx.mock
def test_history_empty(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_lookups()
    respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/versions").respond(
        json={"value": []}
    )
    assert history(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/policy.docx") == []


@respx.mock
def test_history_sends_top_and_orderby(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_lookups()
    route = respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/versions").respond(
        json={"value": []}
    )
    history(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/policy.docx", limit=42)
    url_str = str(route.calls.last.request.url)
    assert "%24top=42" in url_str or "$top=42" in url_str
    assert "lastModifiedDateTime+desc" in url_str or "lastModifiedDateTime%20desc" in url_str


@respx.mock
def test_history_propagates_404(store_with_fresh_token: None) -> None:
    """When the default-drive lookup 404s and no library matches the path,
    the original 404 is surfaced. The library-fallback in resolve_drive_item
    adds one /drives lookup per 404 — we mock it to return an empty list."""
    del store_with_fresh_token
    respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(json={"id": SITE_ID})
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/missing.docx").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drives").respond(json={"value": []})
    with pytest.raises(httpx.HTTPStatusError):
        history(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/missing.docx")


# ---------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------


def test_history_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        history("")


def test_history_rejects_blank_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        history("   ")


def test_history_rejects_site_only_url() -> None:
    with pytest.raises(ValueError, match="needs a file URL"):
        history(f"https://{SITE_HOST}{SITE_PATH}")


def test_history_rejects_zero_limit() -> None:
    with pytest.raises(ValueError, match="limit must be positive"):
        history(f"https://{SITE_HOST}{SITE_PATH}/foo.txt", limit=0)


# ---------------------------------------------------------------------
# Helper extraction
# ---------------------------------------------------------------------


def test_extract_versions_handles_missing_value() -> None:
    assert _extract_versions({}) == []


def test_extract_versions_minimal_entry() -> None:
    payload = {"value": [{"id": "1.0"}]}
    [v] = _extract_versions(payload)
    assert v["id"] == "1.0"
    assert v["last_modified"] == ""
    assert v["last_modified_by"] is None
    assert v["size"] == 0


def test_extract_user_displayName_wins() -> None:
    assert _extract_user({"user": {"displayName": "Alice", "email": "a@x"}}) == "Alice"


def test_extract_user_email_fallback() -> None:
    assert _extract_user({"user": {"email": "a@x"}}) == "a@x"


def test_extract_user_none_when_missing() -> None:
    assert _extract_user(None) is None
    assert _extract_user({}) is None
    assert _extract_user({"user": None}) is None
