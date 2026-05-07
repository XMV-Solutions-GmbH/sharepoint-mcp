# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_read."""

from __future__ import annotations

import time
from collections.abc import Iterator
from pathlib import Path

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.read import read_file


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


SITE_HOST = "contoso.sharepoint.com"
SITE_PATH = "/sites/foo"
SITE_ID = "contoso.sharepoint.com,site-guid,web-guid"


def _mock_site_lookup() -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(
        json={"id": SITE_ID},
    )


@respx.mock
def test_read_file_writes_temp_file_with_content(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    payload = b"# README\n\nharness seed content\n"
    _mock_site_lookup()
    respx.get(
        f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/README.md:/content",
    ).respond(content=payload)

    path = read_file(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/README.md")
    try:
        assert Path(path).read_bytes() == payload
        assert path.endswith(".md")  # extension preserved
    finally:
        Path(path).unlink(missing_ok=True)


@respx.mock
def test_read_file_subfolder_path(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(
        f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/policies/iso.docx:/content",
    ).respond(content=b"docx-bytes-here")

    path = read_file(
        f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/policies/iso.docx",
    )
    try:
        assert Path(path).read_bytes() == b"docx-bytes-here"
        assert path.endswith(".docx")
    finally:
        Path(path).unlink(missing_ok=True)


@respx.mock
def test_read_file_sends_bearer_on_both_calls(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    site_route = _mock_site_lookup()
    content_route = respx.get(
        f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/x.txt:/content",
    ).respond(content=b"x")
    path = read_file(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/x.txt")
    try:
        for call in (site_route.calls.last, content_route.calls.last):
            assert call.request.headers.get("authorization") == "Bearer AT-test"
    finally:
        Path(path).unlink(missing_ok=True)


@respx.mock
def test_read_file_propagates_404(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(
        f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/missing.txt:/content",
    ).respond(404, json={"error": {"code": "itemNotFound"}})
    with pytest.raises(httpx.HTTPStatusError):
        read_file(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/missing.txt")


# ---------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------


def test_read_file_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        read_file("")


def test_read_file_rejects_blank_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        read_file("   ")


def test_read_file_rejects_site_url_no_item() -> None:
    with pytest.raises(ValueError, match="needs a file URL"):
        read_file(f"https://{SITE_HOST}{SITE_PATH}")


def test_read_file_rejects_folder_url_no_item() -> None:
    """A site URL with just the library segment resolves to empty item_path."""
    with pytest.raises(ValueError, match="needs a file URL"):
        read_file(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents")
