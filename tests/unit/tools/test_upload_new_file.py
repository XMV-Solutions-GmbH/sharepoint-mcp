# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_upload_new_file (issue #87)."""

from __future__ import annotations

import base64
import time
from collections.abc import Iterator

import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.upload_new_file import FileAlreadyExistsError, upload_new_file


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


def _mock_site() -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(
        json={"id": SITE_ID},
    )


def _b64(data: bytes) -> str:
    return base64.b64encode(data).decode()


ITEM_RESPONSE = {
    "id": "01ABCDEF",
    "eTag": '"abc123,1"',
    "webUrl": f"https://{SITE_HOST}{SITE_PATH}/Shared%20Documents/report.md",
    "size": 42,
}


# ---------------------------------------------------------------------
# Happy path
# ---------------------------------------------------------------------


@respx.mock
def test_upload_new_small_file_success(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/report.md").respond(
        404, json={"error": {"code": "itemNotFound"}},
    )
    respx.put(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/report.md:/content").respond(
        201, json=ITEM_RESPONSE,
    )

    result = upload_new_file(SITE_URL, "report.md", _b64(b"# Hello\n"))

    assert result["item_id"] == "01ABCDEF"
    assert result["etag"] == '"abc123,1"'
    assert result["web_url"].endswith("report.md")
    assert result["size"] == 42


@respx.mock
def test_upload_new_file_strips_library_prefix(store_with_fresh_token: None) -> None:
    """'Shared Documents/report.md' resolves identically to 'report.md'."""
    del store_with_fresh_token
    _mock_site()
    exist = respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/report.md").respond(404)
    put = respx.put(
        f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/report.md:/content"
    ).respond(201, json=ITEM_RESPONSE)

    upload_new_file(SITE_URL, "Shared Documents/report.md", _b64(b"x"))

    assert exist.called
    assert put.called


@respx.mock
def test_upload_new_file_nested_path(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/2026/Q2/report.md").respond(404)
    respx.put(
        f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/2026/Q2/report.md:/content"
    ).respond(201, json={**ITEM_RESPONSE, "size": 7})

    result = upload_new_file(SITE_URL, "2026/Q2/report.md", _b64(b"content"))

    assert result["size"] == 7


# ---------------------------------------------------------------------
# File-already-exists guard
# ---------------------------------------------------------------------


@respx.mock
def test_upload_new_file_409_raises_file_already_exists(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    # Existence check returns 200 → file is there
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/report.md").respond(
        200, json={"id": "01EXISTING", "name": "report.md"},
    )

    with pytest.raises(FileAlreadyExistsError, match="sp_open"):
        upload_new_file(SITE_URL, "report.md", _b64(b"x"))


# ---------------------------------------------------------------------
# Parent folder not found
# ---------------------------------------------------------------------


@respx.mock
def test_upload_new_file_404_parent_not_found_propagates(store_with_fresh_token: None) -> None:
    """Graph 404 on the PUT (parent folder missing) is raised as HTTPStatusError."""
    del store_with_fresh_token
    _mock_site()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/missing/report.md").respond(404)
    # PUT also 404 — parent folder doesn't exist
    respx.put(
        f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/missing/report.md:/content"
    ).respond(404, json={"error": {"code": "itemNotFound"}})

    import httpx

    with pytest.raises(httpx.HTTPStatusError):
        upload_new_file(SITE_URL, "missing/report.md", _b64(b"x"))


# ---------------------------------------------------------------------
# Base64 decoding + content
# ---------------------------------------------------------------------


@respx.mock
def test_upload_new_file_body_equals_decoded_bytes(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    payload = b"exact-content-bytes-XYZ"
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/f.bin").respond(404)
    put = respx.put(
        f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/f.bin:/content"
    ).respond(201, json={**ITEM_RESPONSE, "size": len(payload)})

    upload_new_file(SITE_URL, "f.bin", _b64(payload))

    assert put.calls.last.request.read() == payload


@respx.mock
def test_upload_new_file_empty_content_allowed(store_with_fresh_token: None) -> None:
    """0-byte file is valid — e.g. creating an empty placeholder."""
    del store_with_fresh_token
    _mock_site()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/placeholder.txt").respond(404)
    put = respx.put(
        f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/placeholder.txt:/content"
    ).respond(201, json={**ITEM_RESPONSE, "size": 0})

    result = upload_new_file(SITE_URL, "placeholder.txt", _b64(b""))

    assert put.calls.last.request.read() == b""
    assert result["size"] == 0


# ---------------------------------------------------------------------
# Content-Type inference
# ---------------------------------------------------------------------


@respx.mock
def test_upload_new_file_content_type_docx(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/doc.docx").respond(404)
    put = respx.put(
        f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/doc.docx:/content"
    ).respond(201, json=ITEM_RESPONSE)

    upload_new_file(SITE_URL, "doc.docx", _b64(b"<docx>"))

    ct = put.calls.last.request.headers.get("content-type", "")
    assert "wordprocessingml" in ct or "officedocument" in ct


@respx.mock
def test_upload_new_file_content_type_pdf(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/report.pdf").respond(404)
    put = respx.put(
        f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/report.pdf:/content"
    ).respond(201, json=ITEM_RESPONSE)

    upload_new_file(SITE_URL, "report.pdf", _b64(b"%PDF"))

    assert put.calls.last.request.headers.get("content-type") == "application/pdf"


@respx.mock
def test_upload_new_file_unknown_extension_uses_octet_stream(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/data.xyz123").respond(404)
    put = respx.put(
        f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/data.xyz123:/content"
    ).respond(201, json=ITEM_RESPONSE)

    upload_new_file(SITE_URL, "data.xyz123", _b64(b"binary"))

    assert put.calls.last.request.headers.get("content-type") == "application/octet-stream"


# ---------------------------------------------------------------------
# Auth header propagation
# ---------------------------------------------------------------------


@respx.mock
def test_upload_new_file_bearer_on_all_calls(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    site_r = _mock_site()
    exist_r = respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/f.txt").respond(404)
    put_r = respx.put(
        f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/f.txt:/content"
    ).respond(201, json=ITEM_RESPONSE)

    upload_new_file(SITE_URL, "f.txt", _b64(b"x"))

    for route in (site_r, exist_r, put_r):
        assert route.calls.last.request.headers.get("authorization") == "Bearer AT-test"


# ---------------------------------------------------------------------
# Input validation (no Graph calls needed)
# ---------------------------------------------------------------------


def test_upload_new_file_invalid_base64_raises() -> None:
    with pytest.raises(ValueError, match="not valid base64"):
        upload_new_file(SITE_URL, "f.txt", "!!!not-base64!!!")


def test_upload_new_file_oversized_content_raises() -> None:
    big = base64.b64encode(b"x" * (4 * 1024 * 1024 + 1)).decode()
    with pytest.raises(ValueError, match="4 MB"):
        upload_new_file(SITE_URL, "f.bin", big)


def test_upload_new_file_empty_site_url_raises() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        upload_new_file("", "f.txt", _b64(b"x"))


def test_upload_new_file_empty_path_raises() -> None:
    with pytest.raises(ValueError, match="non-empty path"):
        upload_new_file(SITE_URL, "", _b64(b"x"))


def test_upload_new_file_slash_only_path_raises() -> None:
    with pytest.raises(ValueError, match="filename"):
        upload_new_file(SITE_URL, "/", _b64(b"x"))
