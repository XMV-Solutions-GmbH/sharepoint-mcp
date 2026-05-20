# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_drive_file_move (issue #95)."""

from __future__ import annotations

import time
from collections.abc import Iterator

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.move_file import move_file


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
WEB_URL = f"https://{SITE_HOST}{SITE_PATH}/Shared%20Documents/Archive/report.md"


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


def _mock_patch(web_url: str = WEB_URL) -> respx.Route:
    return respx.patch(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{SRC_ITEM_ID}").respond(
        json={"id": "item-new", "webUrl": web_url}
    )


# ------------------------------------------------------------------
# Happy path — cross-folder move
# ------------------------------------------------------------------


@respx.mock
def test_move_file_cross_folder(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_source("2026/Q2/report.md")
    _mock_resolve_dest_folder("Archive/2026/Q2")
    _mock_patch()

    result = move_file(SITE_URL, "2026/Q2/report.md", "Archive/2026/Q2/report.md")

    assert result["moved"] is True
    assert result["source"] == "2026/Q2/report.md"
    assert result["destination"] == "Archive/2026/Q2/report.md"
    assert result["web_url"] == WEB_URL


@respx.mock
def test_move_file_patch_body_contains_parent_and_name(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_source("old/file.txt")
    _mock_resolve_dest_folder("new-folder")
    route = _mock_patch()

    move_file(SITE_URL, "old/file.txt", "new-folder/file.txt")

    import json

    body = json.loads(route.calls.last.request.content)
    assert body["parentReference"]["id"] == DEST_FOLDER_ITEM_ID
    assert body["name"] == "file.txt"


@respx.mock
def test_move_file_rename_in_place(store_with_fresh_token: None) -> None:
    """Rename a file at the drive root — destination has no parent path segment."""
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_source("old-name.txt")
    # Root folder lookup
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root").respond(
        json={"id": "root-item-id", "parentReference": {"driveId": DRIVE_ID}}
    )
    route = _mock_patch(
        web_url="https://contoso.sharepoint.com/sites/foo/Shared%20Documents/new-name.txt"
    )

    result = move_file(SITE_URL, "old-name.txt", "new-name.txt")

    assert result["moved"] is True
    assert result["destination"] == "new-name.txt"
    import json

    body = json.loads(route.calls.last.request.content)
    assert body["name"] == "new-name.txt"


@respx.mock
def test_move_file_combined_move_and_rename(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_source("Work/draft.md")
    _mock_resolve_dest_folder("Archive")
    route = _mock_patch()

    move_file(SITE_URL, "Work/draft.md", "Archive/final.md")

    import json

    body = json.loads(route.calls.last.request.content)
    assert body["name"] == "final.md"
    assert body["parentReference"]["id"] == DEST_FOLDER_ITEM_ID


@respx.mock
def test_move_file_strips_leading_slashes(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_source("2026/file.txt")
    _mock_resolve_dest_folder("Archive")
    _mock_patch()

    result = move_file(SITE_URL, "/2026/file.txt", "/Archive/file.txt")

    assert result["source"] == "2026/file.txt"
    assert result["destination"] == "Archive/file.txt"


# ------------------------------------------------------------------
# Error paths
# ------------------------------------------------------------------


@respx.mock
def test_move_file_propagates_source_404(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/missing.txt").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drives").respond(json={"value": []})

    with pytest.raises(httpx.HTTPStatusError) as exc_info:
        move_file(SITE_URL, "missing.txt", "Archive/missing.txt")

    assert exc_info.value.response.status_code == 404


@respx.mock
def test_move_file_propagates_dest_folder_404(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_source("file.txt")
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/nonexistent-folder").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drives").respond(json={"value": []})

    with pytest.raises(httpx.HTTPStatusError) as exc_info:
        move_file(SITE_URL, "file.txt", "nonexistent-folder/file.txt")

    assert exc_info.value.response.status_code == 404


def test_move_file_rejects_empty_site_url() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        move_file("", "src.txt", "dst.txt")


def test_move_file_rejects_empty_source() -> None:
    with pytest.raises(ValueError, match="non-empty source_path"):
        move_file(SITE_URL, "", "dst.txt")


def test_move_file_rejects_empty_destination() -> None:
    with pytest.raises(ValueError, match="non-empty destination_path"):
        move_file(SITE_URL, "src.txt", "")


def test_move_file_rejects_whitespace_source() -> None:
    with pytest.raises(ValueError, match="non-empty source_path"):
        move_file(SITE_URL, "   ", "dst.txt")


def test_move_file_rejects_whitespace_destination() -> None:
    with pytest.raises(ValueError, match="non-empty destination_path"):
        move_file(SITE_URL, "src.txt", "   ")


# ------------------------------------------------------------------
# http injection
# ------------------------------------------------------------------


@respx.mock
def test_move_file_reuses_injected_client(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_resolve_source("old/file.txt")
    _mock_resolve_dest_folder("new-folder")
    _mock_patch()

    with httpx.Client() as client:
        move_file(SITE_URL, "old/file.txt", "new-folder/file.txt", http=client)
        assert not client.is_closed
