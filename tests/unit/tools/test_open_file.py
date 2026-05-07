# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_open."""

from __future__ import annotations

import time
from collections.abc import Iterator
from pathlib import Path

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.checkout_registry import CheckoutRegistry
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.open_file import CheckoutConflictError, open_file


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
DRIVE_ID = "b!drive-id"
ITEM_ID = "01ITEM"


def _mock_site_lookup() -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(
        json={"id": SITE_ID},
    )


def _mock_item_lookup(item_path: str, name: str = "policy.docx", etag: str = "v1") -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/{item_path}").respond(
        json={
            "id": ITEM_ID,
            "name": name,
            "eTag": etag,
            "parentReference": {"driveId": DRIVE_ID},
        },
    )


# ---------------------------------------------------------------------
# Happy path: lock + download + register
# ---------------------------------------------------------------------


@respx.mock
def test_open_file_full_flow(store_with_fresh_token: None, tmp_path: Path) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    _mock_item_lookup("policies/iso.docx", name="iso.docx", etag='"abc123,1"')
    respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/checkout").respond(204)
    respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/content").respond(
        content=b"<docx-bytes>",
    )

    local = open_file(
        f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/policies/iso.docx",
        base_dir=tmp_path,
        now=lambda: 1_900_000_000.0,
    )

    # Working file written to per-item subdir with original filename
    assert Path(local).exists()
    assert Path(local).read_bytes() == b"<docx-bytes>"
    assert Path(local).name == "iso.docx"
    assert ITEM_ID in str(Path(local).parent)

    # Registry entry has all the IDs + etag
    registry = CheckoutRegistry(profile="default", base_dir=tmp_path)
    [entry] = registry.list_all()
    assert entry.site_id == SITE_ID
    assert entry.drive_id == DRIVE_ID
    assert entry.item_id == ITEM_ID
    assert entry.etag == '"abc123,1"'
    assert entry.since == 1_900_000_000.0


@respx.mock
def test_open_file_uses_canonical_filename_from_response(
    store_with_fresh_token: None, tmp_path: Path
) -> None:
    """Microsoft's `name` field wins over what the URL suggested."""
    del store_with_fresh_token
    _mock_site_lookup()
    _mock_item_lookup("policies/iso.docx", name="iso27001-control-A.5.1.docx")
    respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/checkout").respond(204)
    respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/content").respond(content=b"x")

    local = open_file(
        f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/policies/iso.docx",
        base_dir=tmp_path,
    )
    assert Path(local).name == "iso27001-control-A.5.1.docx"


# ---------------------------------------------------------------------
# Concurrency conflict
# ---------------------------------------------------------------------


@respx.mock
def test_open_file_raises_CheckoutConflictError_on_423(
    store_with_fresh_token: None, tmp_path: Path
) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    _mock_item_lookup("policies/iso.docx")
    respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/checkout").respond(
        423, json={"error": {"code": "locked"}}
    )

    with pytest.raises(CheckoutConflictError, match="already checked out"):
        open_file(
            f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/policies/iso.docx",
            base_dir=tmp_path,
        )

    # No registry entry created when checkout fails
    registry = CheckoutRegistry(profile="default", base_dir=tmp_path)
    assert registry.list_all() == []


@respx.mock
def test_open_file_propagates_404_on_item_lookup(
    store_with_fresh_token: None, tmp_path: Path
) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/policies/missing.docx").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    with pytest.raises(httpx.HTTPStatusError):
        open_file(
            f"https://{SITE_HOST}{SITE_PATH}/Shared Documents/policies/missing.docx",
            base_dir=tmp_path,
        )


# ---------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------


def test_open_file_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        open_file("")


def test_open_file_rejects_site_url(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="needs a file URL"):
        open_file(f"https://{SITE_HOST}{SITE_PATH}", base_dir=tmp_path)


def test_open_file_rejects_folder_url(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="needs a file URL"):
        open_file(f"https://{SITE_HOST}{SITE_PATH}/Shared Documents", base_dir=tmp_path)
