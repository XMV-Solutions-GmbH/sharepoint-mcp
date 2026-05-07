# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_publish."""

from __future__ import annotations

import time
from collections.abc import Iterator
from pathlib import Path

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.publish import publish


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


def _mock_site_lookup() -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(
        json={"id": SITE_ID},
    )


# ---------------------------------------------------------------------
# Happy path: upload to root + sub-folder
# ---------------------------------------------------------------------


@respx.mock
def test_publish_into_default_drive_root(store_with_fresh_token: None, tmp_path: Path) -> None:
    del store_with_fresh_token
    src = tmp_path / "report.md"
    src.write_text("# Report v1\n", encoding="utf-8")

    _mock_site_lookup()
    # Existence check: 404 → free to publish
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/report.md").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    # Upload
    respx.put(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/report.md:/content").respond(
        json={
            "name": "report.md",
            "webUrl": f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/report.md",
            "eTag": '"new-etag,1"',
            "size": 14,
            "lastModifiedDateTime": "2026-05-07T12:00:00Z",
        }
    )

    result = publish(str(src), f"https://{SITE_HOST}{SITE_PATH}")
    assert result["name"] == "report.md"
    assert result["etag"] == '"new-etag,1"'
    assert result["size"] == 14
    assert result["web_url"].endswith("/report.md")


@respx.mock
def test_publish_into_subfolder(store_with_fresh_token: None, tmp_path: Path) -> None:
    del store_with_fresh_token
    src = tmp_path / "policy_draft.docx"
    src.write_bytes(b"<docx-bytes>")

    _mock_site_lookup()
    # New flow: resolve target folder driveItem, then existence + upload via /drives/...
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/drafts").respond(
        json={"id": "FID", "name": "drafts", "parentReference": {"driveId": "DID"}},
    )
    respx.get(f"{GRAPH_BASE}/drives/DID/items/FID:/policy_draft.docx").respond(404)
    respx.put(f"{GRAPH_BASE}/drives/DID/items/FID:/policy_draft.docx:/content").respond(
        json={"name": "policy_draft.docx", "webUrl": "x", "eTag": "y", "size": 12}
    )

    result = publish(
        str(src),
        f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/drafts",
    )
    assert result["name"] == "policy_draft.docx"


@respx.mock
def test_publish_with_explicit_name(store_with_fresh_token: None, tmp_path: Path) -> None:
    """The `name` parameter overrides the local file's basename."""
    del store_with_fresh_token
    src = tmp_path / "tmp_abc123.docx"
    src.write_bytes(b"x")

    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/iso27001-A.5.1.docx").respond(404)
    put_route = respx.put(
        f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/iso27001-A.5.1.docx:/content"
    ).respond(json={"name": "iso27001-A.5.1.docx", "webUrl": "x", "eTag": "y", "size": 1})

    result = publish(
        str(src),
        f"https://{SITE_HOST}{SITE_PATH}",
        name="iso27001-A.5.1.docx",
    )
    assert result["name"] == "iso27001-A.5.1.docx"
    # Confirm the PUT used the explicit name in the URL
    assert "iso27001-A.5.1.docx" in str(put_route.calls.last.request.url)


@respx.mock
def test_publish_uploads_actual_file_bytes(store_with_fresh_token: None, tmp_path: Path) -> None:
    """The PUT body is the local file's bytes verbatim."""
    del store_with_fresh_token
    payload = b"verifiable-content-XYZ-1234567890"
    src = tmp_path / "f.bin"
    src.write_bytes(payload)

    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/f.bin").respond(404)
    put_route = respx.put(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/f.bin:/content").respond(
        json={"name": "f.bin", "webUrl": "x", "eTag": "y"}
    )

    publish(str(src), f"https://{SITE_HOST}{SITE_PATH}")

    assert put_route.calls.last.request.read() == payload


@respx.mock
def test_publish_sends_bearer_on_both_calls(store_with_fresh_token: None, tmp_path: Path) -> None:
    del store_with_fresh_token
    src = tmp_path / "x.txt"
    src.write_bytes(b"y")

    site_route = _mock_site_lookup()
    exist_route = respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/x.txt").respond(404)
    put_route = respx.put(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/x.txt:/content").respond(
        json={"name": "x.txt", "webUrl": "x", "eTag": "y"}
    )

    publish(str(src), f"https://{SITE_HOST}{SITE_PATH}")
    for route in (site_route, exist_route, put_route):
        assert route.calls.last.request.headers.get("authorization") == "Bearer AT-test"


# ---------------------------------------------------------------------
# Refuse to overwrite existing files
# ---------------------------------------------------------------------


@respx.mock
def test_publish_refuses_when_target_exists(store_with_fresh_token: None, tmp_path: Path) -> None:
    del store_with_fresh_token
    src = tmp_path / "existing.md"
    src.write_text("local")

    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/existing.md").respond(
        json={"id": "01ITEM", "name": "existing.md"}
    )

    with pytest.raises(FileExistsError, match="Use sp_open"):
        publish(str(src), f"https://{SITE_HOST}{SITE_PATH}")


@respx.mock
def test_publish_propagates_403_on_existence_check(
    store_with_fresh_token: None, tmp_path: Path
) -> None:
    del store_with_fresh_token
    src = tmp_path / "f.md"
    src.write_text("y")

    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/f.md").respond(403)
    with pytest.raises(httpx.HTTPStatusError):
        publish(str(src), f"https://{SITE_HOST}{SITE_PATH}")


@respx.mock
def test_publish_propagates_500_on_upload(store_with_fresh_token: None, tmp_path: Path) -> None:
    del store_with_fresh_token
    src = tmp_path / "f.md"
    src.write_text("y")

    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/f.md").respond(404)
    respx.put(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/f.md:/content").respond(500)
    with pytest.raises(httpx.HTTPStatusError):
        publish(str(src), f"https://{SITE_HOST}{SITE_PATH}")


# ---------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------


def test_publish_rejects_empty_local_path() -> None:
    with pytest.raises(ValueError, match="non-empty local_path"):
        publish("", "https://example/foo")


def test_publish_rejects_blank_local_path() -> None:
    with pytest.raises(ValueError, match="non-empty local_path"):
        publish("   ", "https://example/foo")


def test_publish_rejects_empty_target_url() -> None:
    with pytest.raises(ValueError, match="non-empty target_folder_url"):
        publish("/tmp/x", "")


def test_publish_rejects_missing_local_file(tmp_path: Path) -> None:
    with pytest.raises(FileNotFoundError, match="not found"):
        publish(str(tmp_path / "no-such-file.txt"), "https://example/foo")


def test_publish_rejects_local_dir(tmp_path: Path) -> None:
    with pytest.raises(FileNotFoundError, match="not a file"):
        publish(str(tmp_path), "https://example/foo")


def test_publish_rejects_name_with_path_separators(tmp_path: Path) -> None:
    src = tmp_path / "x.txt"
    src.write_bytes(b"y")
    with pytest.raises(ValueError, match="bare filename"):
        publish(str(src), "https://example/foo", name="../escape.txt")
    with pytest.raises(ValueError, match="bare filename"):
        publish(str(src), "https://example/foo", name="sub/file.txt")
