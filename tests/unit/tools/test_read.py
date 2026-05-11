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


DRIVE_ID = "b!drive"
ITEM_ID = "01ITEM"


def _mock_share_lookup(file_name: str = "README.md") -> respx.Route:
    """Primary resolver: `/shares/{shareId}/driveItem`. Match any
    shareId — the encoding contains `:` and `/` which need URL-
    escaping in `respx.get(...)` arguments, and the input URL varies
    per test, so an `__regex` match is cleaner.
    """
    import re

    return respx.get(re.compile(rf"{re.escape(GRAPH_BASE)}/shares/u!.*?/driveItem")).respond(
        json={
            "id": ITEM_ID,
            "name": file_name,
            "parentReference": {"driveId": DRIVE_ID},
        },
    )


def _mock_site_lookup() -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(
        json={"id": SITE_ID},
    )


def _mock_item_lookup(item_path: str) -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/{item_path}").respond(
        json={
            "id": ITEM_ID,
            "name": item_path.rsplit("/", 1)[-1],
            "parentReference": {"driveId": DRIVE_ID},
        },
    )


@respx.mock
def test_read_file_writes_temp_file_with_content(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    payload = b"# README\n\nharness seed content\n"
    _mock_site_lookup()
    _mock_item_lookup("README.md")
    respx.get(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/content",
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
    _mock_item_lookup("policies/iso.docx")
    respx.get(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/content",
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
def test_read_file_localized_library_name_roundtrip(store_with_fresh_token: None) -> None:
    """Regression test for #79: German tenant URLs with 'Freigegebene
    Dokumente' (the German display name of the default library) must
    round-trip. resolve_drive_item_full's first-fallback strips the
    library segment and retries against the default drive root."""
    del store_with_fresh_token
    _mock_site_lookup()
    # Primary 404 (path includes the German library segment).
    respx.get(
        f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/Freigegebene Dokumente/Finanzen/steuer.pdf"
    ).respond(404, json={"error": {"code": "itemNotFound"}})
    # First fallback: strip "Freigegebene Dokumente", try default drive.
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/Finanzen/steuer.pdf").respond(
        json={
            "id": ITEM_ID,
            "name": "steuer.pdf",
            "parentReference": {"driveId": DRIVE_ID},
        },
    )
    respx.get(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/content",
    ).respond(content=b"pdf-bytes")
    path = read_file(
        f"https://{SITE_HOST}{SITE_PATH}/Freigegebene Dokumente/Finanzen/steuer.pdf",
    )
    try:
        assert Path(path).read_bytes() == b"pdf-bytes"
    finally:
        Path(path).unlink(missing_ok=True)


@respx.mock
def test_read_file_sends_bearer_on_both_calls(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    site_route = _mock_site_lookup()
    _mock_item_lookup("x.txt")
    content_route = respx.get(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/content",
    ).respond(content=b"x")
    path = read_file(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/x.txt")
    try:
        for call in (site_route.calls.last, content_route.calls.last):
            assert call.request.headers.get("authorization") == "Bearer AT-test"
    finally:
        Path(path).unlink(missing_ok=True)


@respx.mock
def test_read_file_propagates_404(store_with_fresh_token: None) -> None:
    """Primary 404 → strip-first-segment retry 404 → library-search
    finds no match → original 404 propagates."""
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/missing.txt").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drives").respond(json={"value": []})
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
