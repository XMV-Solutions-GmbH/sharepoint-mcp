# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_create_folder (issue #86)."""

from __future__ import annotations

import time
from collections.abc import Iterator
from typing import Any

import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.create_folder import create_folder


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
SITE_URL = f"https://{SITE_HOST}{SITE_PATH}"
FOLDER_WEB_URL = f"https://{SITE_HOST}{SITE_PATH}/Shared%20Documents/2026"


def _mock_site() -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(
        json={"id": SITE_ID},
    )


def _folder_created_response(name: str, web_url: str = FOLDER_WEB_URL) -> dict[str, Any]:
    return {"id": "FID", "name": name, "webUrl": web_url, "folder": {}}


def _name_already_exists_409() -> dict[str, Any]:
    return {"error": {"code": "nameAlreadyExists", "message": "Name already exists"}}


# ---------------------------------------------------------------------
# Happy path — single segment
# ---------------------------------------------------------------------


@respx.mock
def test_create_folder_single_new_segment(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    post = respx.post(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/children").respond(
        201,
        json=_folder_created_response("2026"),
    )

    result = create_folder(SITE_URL, "2026")

    assert result["created"] == ["2026"]
    assert result["already_existed"] == []
    assert result["web_url"] == FOLDER_WEB_URL
    assert post.called


@respx.mock
def test_create_folder_deep_path_all_new(store_with_fresh_token: None) -> None:
    """Three nested segments → three POST calls, all return 201."""
    del store_with_fresh_token
    _mock_site()
    respx.post(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/children").respond(
        201,
        json=_folder_created_response("2026"),
    )
    respx.post(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/2026/children").respond(
        201,
        json=_folder_created_response("Q2"),
    )
    q2_web_url = f"https://{SITE_HOST}{SITE_PATH}/Shared%20Documents/2026/Q2/Reports"
    respx.post(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/2026/Q2/children").respond(
        201,
        json=_folder_created_response("Reports", web_url=q2_web_url),
    )

    result = create_folder(SITE_URL, "2026/Q2/Reports")

    assert result["created"] == ["2026", "2026/Q2", "2026/Q2/Reports"]
    assert result["already_existed"] == []
    assert result["web_url"] == q2_web_url


# ---------------------------------------------------------------------
# Idempotency — partial and full pre-existence
# ---------------------------------------------------------------------


@respx.mock
def test_create_folder_partial_path_exists(store_with_fresh_token: None) -> None:
    """First segment already exists (409) → remaining segments created."""
    del store_with_fresh_token
    _mock_site()
    # "2026" already exists
    respx.post(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/children").respond(
        409,
        json=_name_already_exists_409(),
    )
    # "2026/Q2" is new
    q2_url = f"https://{SITE_HOST}{SITE_PATH}/Shared%20Documents/2026/Q2"
    respx.post(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/2026/children").respond(
        201,
        json=_folder_created_response("Q2", web_url=q2_url),
    )

    result = create_folder(SITE_URL, "2026/Q2")

    assert result["created"] == ["2026/Q2"]
    assert result["already_existed"] == ["2026"]
    assert result["web_url"] == q2_url


@respx.mock
def test_create_folder_all_segments_exist(store_with_fresh_token: None) -> None:
    """All segments already exist → created=[], already_existed=[all], web_url fetched via GET."""
    del store_with_fresh_token
    _mock_site()
    respx.post(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/children").respond(
        409,
        json=_name_already_exists_409(),
    )
    respx.post(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/2026/children").respond(
        409,
        json=_name_already_exists_409(),
    )
    # Extra GET to fetch web_url for deepest folder
    q2_url = f"https://{SITE_HOST}{SITE_PATH}/Shared%20Documents/2026/Q2"
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/2026/Q2").respond(
        200,
        json={"id": "FID", "webUrl": q2_url},
    )

    result = create_folder(SITE_URL, "2026/Q2")

    assert result["created"] == []
    assert result["already_existed"] == ["2026", "2026/Q2"]
    assert result["web_url"] == q2_url


# ---------------------------------------------------------------------
# Collision with a FILE at the same name
# ---------------------------------------------------------------------


@respx.mock
def test_create_folder_name_collision_with_file_raises(store_with_fresh_token: None) -> None:
    """409 with a code OTHER than nameAlreadyExists is propagated as HTTPStatusError."""
    del store_with_fresh_token
    _mock_site()
    respx.post(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/children").respond(
        409,
        json={"error": {"code": "someOtherConflict", "message": "not a folder"}},
    )

    import httpx

    with pytest.raises(httpx.HTTPStatusError):
        create_folder(SITE_URL, "documents")


# ---------------------------------------------------------------------
# Path normalisation
# ---------------------------------------------------------------------


@respx.mock
def test_create_folder_strips_shared_documents_prefix(store_with_fresh_token: None) -> None:
    """'Shared Documents/2026' and '2026' should hit the same Graph endpoint."""
    del store_with_fresh_token
    _mock_site()
    post = respx.post(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/children").respond(
        201,
        json=_folder_created_response("2026"),
    )

    result = create_folder(SITE_URL, "Shared Documents/2026")

    assert result["created"] == ["2026"]
    assert post.called


@respx.mock
def test_create_folder_strips_leading_trailing_slashes(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    post = respx.post(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/children").respond(
        201,
        json=_folder_created_response("2026"),
    )

    result = create_folder(SITE_URL, "/2026/")

    assert result["created"] == ["2026"]
    assert post.called


@respx.mock
def test_create_folder_returns_web_url_from_creation_response(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    expected_url = "https://contoso.sharepoint.com/sites/foo/Shared%20Documents/reports"
    respx.post(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/children").respond(
        201,
        json=_folder_created_response("reports", web_url=expected_url),
    )

    result = create_folder(SITE_URL, "reports")

    assert result["web_url"] == expected_url


# ---------------------------------------------------------------------
# Error propagation
# ---------------------------------------------------------------------


@respx.mock
def test_create_folder_site_404_raises(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )

    import httpx

    with pytest.raises(httpx.HTTPStatusError):
        create_folder(SITE_URL, "2026")


@respx.mock
def test_create_folder_post_500_raises(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    respx.post(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root/children").respond(500)

    import httpx

    with pytest.raises(httpx.HTTPStatusError):
        create_folder(SITE_URL, "2026")


# ---------------------------------------------------------------------
# Input validation (no Graph calls needed)
# ---------------------------------------------------------------------


def test_create_folder_empty_site_url_raises() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        create_folder("", "2026")


def test_create_folder_blank_site_url_raises() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        create_folder("   ", "2026")


def test_create_folder_empty_path_raises() -> None:
    with pytest.raises(ValueError, match="non-empty path"):
        create_folder(SITE_URL, "")


def test_create_folder_slash_only_path_raises() -> None:
    with pytest.raises(ValueError, match="no valid folder segments"):
        create_folder(SITE_URL, "/")


def test_create_folder_library_prefix_only_path_raises() -> None:
    with pytest.raises(ValueError, match="no valid folder segments"):
        create_folder(SITE_URL, "Shared Documents/")
