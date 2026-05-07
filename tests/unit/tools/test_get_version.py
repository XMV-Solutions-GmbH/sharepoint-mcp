# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_get_version."""

from __future__ import annotations

import time
from collections.abc import Iterator
from pathlib import Path

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.get_version import get_version


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


def _mock_lookups(item_path: str = "policy.docx") -> None:
    respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(json={"id": SITE_ID})
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/{item_path}").respond(
        json={
            "id": ITEM_ID,
            "name": item_path.split("/")[-1],
            "parentReference": {"driveId": DRIVE_ID},
        }
    )


# ---------------------------------------------------------------------
# Happy path
# ---------------------------------------------------------------------


@respx.mock
def test_get_version_writes_temp_with_content(
    store_with_fresh_token: None,
) -> None:
    del store_with_fresh_token
    payload = b"<v2.0-bytes>"
    _mock_lookups()
    respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/versions/2.0/content").respond(
        content=payload
    )

    path = get_version(
        f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/policy.docx",
        version_id="2.0",
    )
    try:
        assert Path(path).read_bytes() == payload
        assert path.endswith(".docx"), f"extension preserved: {path}"
        # Naming convention: temp file includes _v<version-id>_ infix
        assert "_v2.0_" in Path(path).name or "_v2_0_" in Path(path).name
    finally:
        Path(path).unlink(missing_ok=True)


@respx.mock
def test_get_version_handles_subfolder_path(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_lookups("policies/iso.docx")
    respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/versions/1.0/content").respond(
        content=b"x"
    )

    path = get_version(
        f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/policies/iso.docx",
        version_id="1.0",
    )
    try:
        assert Path(path).exists()
    finally:
        Path(path).unlink(missing_ok=True)


@respx.mock
def test_get_version_propagates_404_on_unknown_version(
    store_with_fresh_token: None,
) -> None:
    del store_with_fresh_token
    _mock_lookups()
    respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/versions/99.0/content").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    with pytest.raises(httpx.HTTPStatusError):
        get_version(
            f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/policy.docx",
            version_id="99.0",
        )


# ---------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------


def test_get_version_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        get_version("", "1.0")


def test_get_version_rejects_empty_version_id() -> None:
    with pytest.raises(ValueError, match="non-empty version_id"):
        get_version("https://example/foo", "")


def test_get_version_rejects_blank_version_id() -> None:
    with pytest.raises(ValueError, match="non-empty version_id"):
        get_version("https://example/foo", "  ")


def test_get_version_rejects_site_only_url() -> None:
    with pytest.raises(ValueError, match="needs a file URL"):
        get_version(f"https://{SITE_HOST}{SITE_PATH}", "1.0")
