# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_drive_file_copy (issue #96)."""

from __future__ import annotations

import time
from collections.abc import Iterator

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.copy_file import _poll_copy_operation, copy_file


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
SRC_ITEM_ID = "item-src"
DEST_FOLDER_ITEM_ID = "item-dest-folder"
SITE_URL = f"https://{SITE_HOST}{SITE_PATH}"
OPERATION_URL = "https://graph.microsoft.com/v1.0/drives/drive-abc/operations/op-123"
NEW_ITEM_WEB_URL = f"https://{SITE_HOST}{SITE_PATH}/Shared%20Documents/Archive/copy.md"


def _mock_site() -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(json={"id": SITE_ID})


def _mock_resolve_source(path: str) -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/{path}").respond(
        json={
            "id": SRC_ITEM_ID,
            "parentReference": {"driveId": DRIVE_ID},
        }
    )


def _mock_resolve_dest_folder(folder_path: str) -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/{folder_path}").respond(
        json={
            "id": DEST_FOLDER_ITEM_ID,
            "parentReference": {"driveId": DRIVE_ID},
        }
    )


def _mock_copy_202() -> respx.Route:
    return respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{SRC_ITEM_ID}/copy").respond(
        202, headers={"Location": OPERATION_URL}
    )


def _mock_operation_completed() -> respx.Route:
    return respx.get(OPERATION_URL).respond(
        json={"status": "completed", "resourceLink": NEW_ITEM_WEB_URL}
    )


# ------------------------------------------------------------------
# Happy path — 202 async flow
# ------------------------------------------------------------------


@respx.mock
def test_copy_file_returns_copied_true(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_source("Templates/contract.docx")
    _mock_resolve_dest_folder("Projects/ACME")
    _mock_copy_202()
    _mock_operation_completed()

    result = copy_file(SITE_URL, "Templates/contract.docx", "Projects/ACME/contract.docx")

    assert result["copied"] is True
    assert result["source"] == "Templates/contract.docx"
    assert result["destination"] == "Projects/ACME/contract.docx"
    assert result["web_url"] == NEW_ITEM_WEB_URL


@respx.mock
def test_copy_file_post_body_contains_parent_and_name(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_source("src.txt")
    _mock_resolve_dest_folder("dest-folder")
    route = _mock_copy_202()
    _mock_operation_completed()

    copy_file(SITE_URL, "src.txt", "dest-folder/copy.txt")

    import json

    body = json.loads(route.calls.last.request.content)
    assert body["parentReference"]["id"] == DEST_FOLDER_ITEM_ID
    assert body["name"] == "copy.txt"


@respx.mock
def test_copy_file_strips_leading_slashes(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_source("src/file.txt")
    _mock_resolve_dest_folder("dst")
    _mock_copy_202()
    _mock_operation_completed()

    result = copy_file(SITE_URL, "/src/file.txt", "/dst/file.txt")

    assert result["source"] == "src/file.txt"
    assert result["destination"] == "dst/file.txt"


@respx.mock
def test_copy_file_polls_until_completed(store_with_fresh_token: None) -> None:
    """Operation returns inProgress once, then completed."""
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_source("src.txt")
    _mock_resolve_dest_folder("dst")
    _mock_copy_202()
    respx.get(OPERATION_URL).side_effect = [
        httpx.Response(200, json={"status": "inProgress"}),
        httpx.Response(200, json={"status": "completed", "resourceLink": NEW_ITEM_WEB_URL}),
    ]

    result = copy_file(SITE_URL, "src.txt", "dst/src.txt")

    assert result["copied"] is True


@respx.mock
def test_copy_file_handles_synchronous_200(store_with_fresh_token: None) -> None:
    """Some Graph implementations return 200/201 directly (e.g. test tenants)."""
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_source("src.txt")
    _mock_resolve_dest_folder("dst")
    respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{SRC_ITEM_ID}/copy").respond(
        201, json={"id": "new-item", "webUrl": NEW_ITEM_WEB_URL}
    )

    result = copy_file(SITE_URL, "src.txt", "dst/src.txt")

    assert result["copied"] is True
    assert result["web_url"] == NEW_ITEM_WEB_URL


@respx.mock
def test_copy_file_handles_303_cdn_redirect(store_with_fresh_token: None) -> None:
    """Graph returns 303 See Other when copy completes via CDN redirect pattern."""
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_source("src.txt")
    _mock_resolve_dest_folder("dst")
    respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{SRC_ITEM_ID}/copy").respond(
        303, headers={"Location": NEW_ITEM_WEB_URL}
    )

    result = copy_file(SITE_URL, "src.txt", "dst/src.txt")

    assert result["copied"] is True
    assert result["web_url"] == NEW_ITEM_WEB_URL


# ------------------------------------------------------------------
# Error paths
# ------------------------------------------------------------------


@respx.mock
def test_copy_file_raises_on_failed_operation(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_source("src.txt")
    _mock_resolve_dest_folder("dst")
    _mock_copy_202()
    respx.get(OPERATION_URL).respond(
        json={
            "status": "failed",
            "error": {"code": "generalException", "message": "Copy failed"},
        }
    )

    with pytest.raises(RuntimeError, match="Copy failed"):
        copy_file(SITE_URL, "src.txt", "dst/src.txt")


@respx.mock
def test_copy_file_raises_timeout(
    store_with_fresh_token: None, monkeypatch: pytest.MonkeyPatch
) -> None:
    """If operation never completes within timeout, raises TimeoutError."""
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_source("src.txt")
    _mock_resolve_dest_folder("dst")
    _mock_copy_202()
    respx.get(OPERATION_URL).respond(json={"status": "inProgress"})

    # Patch sleep to avoid actual waiting; fake time by overriding monotonic.
    import time as time_module

    call_count = 0

    def fake_monotonic() -> float:
        nonlocal call_count
        call_count += 1
        # First call (deadline set) returns 0, subsequent calls advance past deadline.
        return 0.0 if call_count == 1 else 100.0

    monkeypatch.setattr(time_module, "monotonic", fake_monotonic)
    monkeypatch.setattr(time_module, "sleep", lambda _: None)

    with pytest.raises(TimeoutError, match="did not complete"):
        copy_file(SITE_URL, "src.txt", "dst/src.txt", timeout=60)


@respx.mock
def test_copy_file_raises_without_location_header(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_source("src.txt")
    _mock_resolve_dest_folder("dst")
    respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{SRC_ITEM_ID}/copy").respond(202)

    with pytest.raises(RuntimeError, match="no Location header"):
        copy_file(SITE_URL, "src.txt", "dst/src.txt")


@respx.mock
def test_copy_file_propagates_source_404(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/missing.txt").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drives").respond(json={"value": []})

    with pytest.raises(httpx.HTTPStatusError) as exc_info:
        copy_file(SITE_URL, "missing.txt", "dst/missing.txt")

    assert exc_info.value.response.status_code == 404


def test_copy_file_rejects_empty_site_url() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        copy_file("", "src.txt", "dst/src.txt")


def test_copy_file_rejects_empty_source() -> None:
    with pytest.raises(ValueError, match="non-empty source_path"):
        copy_file(SITE_URL, "", "dst/src.txt")


def test_copy_file_rejects_empty_destination() -> None:
    with pytest.raises(ValueError, match="non-empty destination_path"):
        copy_file(SITE_URL, "src.txt", "")


# ------------------------------------------------------------------
# _poll_copy_operation unit tests (pure logic)
# ------------------------------------------------------------------


@respx.mock
def test_poll_returns_resource_link_on_completed() -> None:
    respx.get(OPERATION_URL).respond(json={"status": "completed", "resourceLink": NEW_ITEM_WEB_URL})
    with httpx.Client() as client:
        url = _poll_copy_operation(client, OPERATION_URL, headers={}, timeout=5)
    assert url == NEW_ITEM_WEB_URL


@respx.mock
def test_poll_raises_runtime_error_on_failed() -> None:
    respx.get(OPERATION_URL).respond(
        json={
            "status": "failed",
            "error": {"code": "quota", "message": "quota exceeded"},
        }
    )
    with httpx.Client() as client:
        with pytest.raises(RuntimeError, match="quota exceeded"):
            _poll_copy_operation(client, OPERATION_URL, headers={}, timeout=5)


@respx.mock
def test_poll_handles_303_redirect() -> None:
    """Graph may respond to the operation-status poll with 303 See Other."""
    respx.get(OPERATION_URL).respond(303, headers={"Location": NEW_ITEM_WEB_URL})
    with httpx.Client() as client:
        url = _poll_copy_operation(client, OPERATION_URL, headers={}, timeout=5)
    assert url == NEW_ITEM_WEB_URL
