# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_download_binary."""

from __future__ import annotations

import base64
import time
from collections.abc import Iterator

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.download_binary import MAX_BYTES, download_binary


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


def _mock_site() -> None:
    respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(json={"id": SITE_ID})


def _mock_item(
    name: str = "photo.jpg",
    mime: str = "image/jpeg",
    size: int = 1024,
) -> None:
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/{name}").respond(
        json={
            "id": ITEM_ID,
            "name": name,
            "size": size,
            "file": {"mimeType": mime},
            "parentReference": {"driveId": DRIVE_ID},
        }
    )


def _mock_content(content: bytes) -> None:
    respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/content").respond(content=content)


# ---------------------------------------------------------------------------
# Happy path
# ---------------------------------------------------------------------------


@respx.mock
def test_download_binary_returns_envelope(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_item(name="photo.jpg", mime="image/jpeg", size=6)
    _mock_content(b"\xff\xd8\xff\xe0\x00\x10")

    result = download_binary(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/photo.jpg")
    assert result["filename"] == "photo.jpg"
    assert result["mime_type"] == "image/jpeg"
    assert result["size_bytes"] == 6
    assert base64.b64decode(result["base64"]) == b"\xff\xd8\xff\xe0\x00\x10"


@respx.mock
def test_download_binary_base64_is_valid_ascii(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_item()
    _mock_content(b"\x00\x01\x02\x03\xff\xfe")

    result = download_binary(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/photo.jpg")
    # Must be pure ASCII (RFC 4648 base64)
    result["base64"].encode("ascii")


@respx.mock
def test_download_binary_sends_bearer(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_item()
    content_route = respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/content").respond(
        content=b"abc"
    )

    download_binary(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/photo.jpg")
    assert content_route.calls.last.request.headers.get("authorization") == "Bearer AT-test"


@respx.mock
def test_download_binary_missing_mime_falls_back(store_with_fresh_token: None) -> None:
    """When Graph omits file.mimeType the tool returns application/octet-stream."""
    del store_with_fresh_token
    _mock_site()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/data.bin").respond(
        json={
            "id": ITEM_ID,
            "name": "data.bin",
            "size": 3,
            "parentReference": {"driveId": DRIVE_ID},
            # no "file" key at all
        }
    )
    _mock_content(b"XYZ")

    result = download_binary(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/data.bin")
    assert result["mime_type"] == "application/octet-stream"


@respx.mock
def test_download_binary_empty_file(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    _mock_item(name="empty.txt", mime="text/plain", size=0)
    _mock_content(b"")

    result = download_binary(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/empty.txt")
    assert result["size_bytes"] == 0
    assert result["base64"] == ""


@respx.mock
def test_download_binary_propagates_404(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/missing.jpg").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drives").respond(json={"value": []})

    with pytest.raises(httpx.HTTPStatusError):
        download_binary(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/missing.jpg")


# ---------------------------------------------------------------------------
# 10 MB guard
# ---------------------------------------------------------------------------


@respx.mock
def test_download_binary_rejects_oversized_by_metadata(store_with_fresh_token: None) -> None:
    """File larger than 10 MB is rejected before downloading content."""
    del store_with_fresh_token
    _mock_site()
    _mock_item(name="huge.zip", mime="application/zip", size=MAX_BYTES + 1)
    # content route must NOT be called; if it is, the test infrastructure
    # would complain about an unexpected call via respx strict mode.

    with pytest.raises(ValueError, match="10 MB"):
        download_binary(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/huge.zip")


@respx.mock
def test_download_binary_rejects_oversized_by_content(store_with_fresh_token: None) -> None:
    """Even if metadata reports size=0, actual content > 10 MB is rejected."""
    del store_with_fresh_token
    _mock_site()
    _mock_item(name="sneaky.bin", mime="application/octet-stream", size=0)
    # Simulate a server that reports zero size but sends more
    _mock_content(b"x" * (MAX_BYTES + 1))

    with pytest.raises(ValueError, match="10 MB"):
        download_binary(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/sneaky.bin")


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------


def test_download_binary_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        download_binary("")


def test_download_binary_rejects_blank_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        download_binary("   ")


def test_download_binary_rejects_site_url() -> None:
    with pytest.raises(ValueError, match="site/folder URL"):
        download_binary(f"https://{SITE_HOST}{SITE_PATH}")


def test_download_binary_rejects_folder_only_url() -> None:
    with pytest.raises(ValueError, match="site/folder URL"):
        download_binary(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents")
