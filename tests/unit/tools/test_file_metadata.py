# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_file_metadata."""

from __future__ import annotations

import time
from collections.abc import Iterator
from typing import Any

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.file_metadata import file_metadata


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


def _mock_lookups(item_name: str = "report.docx") -> None:
    respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(json={"id": SITE_ID})
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/{item_name}").respond(
        json={
            "id": ITEM_ID,
            "name": item_name,
            "parentReference": {"driveId": DRIVE_ID},
        }
    )


_SAMPLE_FIELDS: dict[str, Any] = {
    "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#...",
    "id": "42",
    "Department": "Finance",
    "RetentionLabel": "5-year",
    "Modified": "2026-03-01T10:00:00Z",
}


# ---------------------------------------------------------------------------
# Read (GET) path
# ---------------------------------------------------------------------------


@respx.mock
def test_file_metadata_read_returns_fields(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_lookups()
    respx.get(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/listItem/fields",
    ).respond(json=_SAMPLE_FIELDS)

    result = file_metadata(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/report.docx")
    assert result["Department"] == "Finance"
    assert result["RetentionLabel"] == "5-year"


@respx.mock
def test_file_metadata_read_none_fields_goes_to_get(store_with_fresh_token: None) -> None:
    """Explicitly passing fields=None must use GET, not PATCH."""
    del store_with_fresh_token
    _mock_lookups()
    get_route = respx.get(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/listItem/fields",
    ).respond(json=_SAMPLE_FIELDS)
    respx.patch(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/listItem/fields",
    ).respond(status_code=500)  # must not be called

    file_metadata(
        f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/report.docx",
        fields=None,
    )
    assert get_route.called


@respx.mock
def test_file_metadata_read_sends_bearer(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_lookups()
    fields_route = respx.get(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/listItem/fields",
    ).respond(json=_SAMPLE_FIELDS)

    file_metadata(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/report.docx")
    assert fields_route.calls.last.request.headers.get("authorization") == "Bearer AT-test"


@respx.mock
def test_file_metadata_read_propagates_404(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_lookups()
    respx.get(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/listItem/fields",
    ).respond(404, json={"error": {"code": "itemNotFound"}})

    with pytest.raises(httpx.HTTPStatusError):
        file_metadata(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/report.docx")


# ---------------------------------------------------------------------------
# Write (PATCH) path
# ---------------------------------------------------------------------------


@respx.mock
def test_file_metadata_write_patches_and_returns_updated(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_lookups()
    updated = {**_SAMPLE_FIELDS, "Department": "Legal"}
    patch_route = respx.patch(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/listItem/fields",
    ).respond(json=updated)

    result = file_metadata(
        f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/report.docx",
        fields={"Department": "Legal"},
    )
    assert result["Department"] == "Legal"
    assert patch_route.called


@respx.mock
def test_file_metadata_write_sends_correct_json_body(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_lookups()
    patch_route = respx.patch(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/listItem/fields",
    ).respond(json={"Department": "HR"})

    file_metadata(
        f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/report.docx",
        fields={"Department": "HR"},
    )
    import json

    body = json.loads(patch_route.calls.last.request.content)
    assert body == {"Department": "HR"}


@respx.mock
def test_file_metadata_write_sets_content_type(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_lookups()
    patch_route = respx.patch(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/listItem/fields",
    ).respond(json={})

    file_metadata(
        f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/report.docx",
        fields={"Status": "Approved"},
    )
    ct = patch_route.calls.last.request.headers.get("content-type", "")
    assert "application/json" in ct


@respx.mock
def test_file_metadata_write_propagates_403(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_lookups()
    respx.patch(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/listItem/fields",
    ).respond(403, json={"error": {"code": "accessDenied"}})

    with pytest.raises(httpx.HTTPStatusError):
        file_metadata(
            f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/report.docx",
            fields={"Department": "IT"},
        )


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------


def test_file_metadata_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        file_metadata("")


def test_file_metadata_rejects_blank_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        file_metadata("   ")


def test_file_metadata_rejects_site_url() -> None:
    with pytest.raises(ValueError, match="site/folder URL"):
        file_metadata(f"https://{SITE_HOST}{SITE_PATH}")


def test_file_metadata_rejects_non_dict_fields() -> None:
    with pytest.raises(TypeError, match="fields must be a dict"):
        file_metadata(
            f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/x.docx",
            fields="not-a-dict",  # type: ignore[arg-type]
        )
