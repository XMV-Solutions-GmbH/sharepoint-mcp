# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_save_file."""

from __future__ import annotations

import time
from collections.abc import Iterator
from pathlib import Path

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.checkout_registry import CheckedOutEntry, CheckoutRegistry
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.save import NotCheckedOutError, StaleWriteError, save


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


URL = "https://contoso.sharepoint.com/sites/foo/Shared Documents/policy.docx"
DRIVE_ID = "b!drive"
ITEM_ID = "01ITEM"
ETAG = '"abc123,1"'


def _seed_registry(tmp_path: Path, content: bytes = b"original-content") -> Path:
    """Set up registry + working file. Patches DEFAULT_REGISTRY_DIR via monkeypatch
    in callers; returns the working-file path."""
    registry = CheckoutRegistry(profile="default", base_dir=tmp_path)
    work_dir = tmp_path / "default" / "working" / ITEM_ID
    work_dir.mkdir(parents=True, exist_ok=True)
    work_file = work_dir / "policy.docx"
    work_file.write_bytes(content)
    registry.add(
        CheckedOutEntry(
            path=URL,
            site_id="S1",
            drive_id=DRIVE_ID,
            item_id=ITEM_ID,
            local_path=str(work_file),
            etag=ETAG,
            since=1_900_000_000.0,
        ),
    )
    return work_file


@pytest.fixture
def registry_with_seed(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    return _seed_registry(tmp_path)


# ---------------------------------------------------------------------
# Happy path: minor + major
# ---------------------------------------------------------------------


@respx.mock
def test_save_minor_version_round_trip(
    store_with_fresh_token: None, registry_with_seed: Path, tmp_path: Path
) -> None:
    del store_with_fresh_token
    work_file = registry_with_seed
    work_file.write_bytes(b"updated-content")

    put_route = respx.put(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/content").respond(
        json={"eTag": '"new-etag,1"', "webUrl": URL}
    )
    checkin_route = respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/checkin").respond(
        204
    )
    respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/versions").respond(
        json={"value": [{"id": "2.0", "lastModifiedDateTime": "2026-05-07T12:00:00Z"}]},
    )

    result = save(URL, comment="updated policy text", version="minor")

    # PUT carried If-Match, body had updated content
    put_request = put_route.calls.last.request
    assert put_request.headers.get("if-match") == ETAG
    assert put_request.read() == b"updated-content"

    # Checkin body included the comment, NOT checkInAs (minor = draft)
    checkin_body = checkin_route.calls.last.request.read().decode()
    assert "updated policy text" in checkin_body
    assert "checkInAs" not in checkin_body

    assert result["version_id"] == "2.0"
    assert result["etag"] == '"new-etag,1"'

    # Registry cleared, working file deleted
    assert CheckoutRegistry(profile="default", base_dir=tmp_path).get(URL) is None
    assert not work_file.exists()


@respx.mock
def test_save_major_version_sets_checkInAs_published(
    store_with_fresh_token: None, registry_with_seed: Path
) -> None:
    del store_with_fresh_token, registry_with_seed
    respx.put(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/content").respond(json={})
    checkin_route = respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/checkin").respond(
        204
    )
    respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/versions").respond(
        json={"value": [{"id": "3.0"}]}
    )

    save(URL, comment="major release", version="major")

    body = checkin_route.calls.last.request.read().decode()
    assert "published" in body  # checkInAs="published"


# ---------------------------------------------------------------------
# Stale-write detection (ETag mismatch)
# ---------------------------------------------------------------------


@respx.mock
def test_save_412_raises_StaleWriteError(
    store_with_fresh_token: None, registry_with_seed: Path, tmp_path: Path
) -> None:
    del store_with_fresh_token
    respx.put(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/content").respond(
        412, json={"error": {"code": "preconditionFailed"}}
    )
    with pytest.raises(StaleWriteError, match="changed under us"):
        save(URL, comment="updated", version="minor")

    # Registry NOT cleared — caller still owns the lock
    assert CheckoutRegistry(profile="default", base_dir=tmp_path).get(URL) is not None
    # Working file still present
    assert registry_with_seed.exists()


# ---------------------------------------------------------------------
# Pre-conditions
# ---------------------------------------------------------------------


def test_save_without_prior_open_raises(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    with pytest.raises(NotCheckedOutError, match="Call sp_open_file first"):
        save(URL, comment="updated", version="minor")


def test_save_missing_working_file_raises(
    store_with_fresh_token: None, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    del store_with_fresh_token
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    work_file = _seed_registry(tmp_path)
    work_file.unlink()
    with pytest.raises(FileNotFoundError, match="Working copy missing"):
        save(URL, comment="updated", version="minor")


# ---------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------


def test_save_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        save("", comment="x")


def test_save_rejects_empty_comment() -> None:
    with pytest.raises(ValueError, match="non-empty comment"):
        save(URL, comment="")


def test_save_rejects_blank_comment() -> None:
    with pytest.raises(ValueError, match="non-empty comment"):
        save(URL, comment="   ")


def test_save_rejects_invalid_version() -> None:
    with pytest.raises(ValueError, match="must be 'minor' or 'major'"):
        save(URL, comment="x", version="patch")  # type: ignore[arg-type]


# ---------------------------------------------------------------------
# Other HTTP errors propagate
# ---------------------------------------------------------------------


@respx.mock
def test_save_500_propagates(store_with_fresh_token: None, registry_with_seed: Path) -> None:
    del store_with_fresh_token, registry_with_seed
    respx.put(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/content").respond(500)
    with pytest.raises(httpx.HTTPStatusError):
        save(URL, comment="updated", version="minor")


# ---------------------------------------------------------------------
# Resumable upload path-switching (#38)
# ---------------------------------------------------------------------


@respx.mock
def test_save_uses_resumable_upload_when_file_exceeds_threshold(
    store_with_fresh_token: None,
    registry_with_seed: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """File larger than the threshold goes through createUploadSession."""
    del store_with_fresh_token
    monkeypatch.setenv("SP_CHUNKED_UPLOAD_THRESHOLD_MB", "1")  # 1 MB threshold
    work_file = registry_with_seed
    work_file.write_bytes(b"\x00" * (2 * 1024 * 1024))  # 2 MB

    # Single-shot route should NOT fire
    single_shot = respx.put(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/content").respond(200)
    create_session = respx.post(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createUploadSession",
    ).respond(json={"uploadUrl": "https://upload.example/abc"})
    respx.put("https://upload.example/abc").respond(
        201,
        json={"id": ITEM_ID, "eTag": '"new,1"', "webUrl": URL},
    )
    respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/checkin").respond(204)
    respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/versions").respond(
        json={"value": [{"id": "5.0"}]}
    )

    result = save(URL, comment="big upload", version="minor")
    assert single_shot.call_count == 0
    assert create_session.call_count == 1
    assert result["version_id"] == "5.0"
    assert result["etag"] == '"new,1"'


@respx.mock
def test_save_uses_single_shot_when_file_under_threshold(
    store_with_fresh_token: None,
    registry_with_seed: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """File at-or-below the threshold uses single-shot PUT."""
    del store_with_fresh_token
    monkeypatch.setenv("SP_CHUNKED_UPLOAD_THRESHOLD_MB", "100")
    # registry_with_seed writes ~16 bytes; well under 100 MB
    create_session = respx.post(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createUploadSession",
    ).respond(200)
    single_shot = respx.put(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/content").respond(
        json={"eTag": '"s,1"'}
    )
    respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/checkin").respond(204)
    respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/versions").respond(
        json={"value": [{"id": "1.0"}]}
    )
    save(URL, comment="small", version="minor")
    assert single_shot.call_count == 1
    assert create_session.call_count == 0


@respx.mock
def test_save_resumable_412_translates_to_stale_write_error(
    store_with_fresh_token: None,
    registry_with_seed: Path,
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """Resumable path also raises StaleWriteError on 412 — same agent contract
    as the single-shot path."""
    del store_with_fresh_token
    monkeypatch.setenv("SP_CHUNKED_UPLOAD_THRESHOLD_MB", "1")
    work_file = registry_with_seed
    work_file.write_bytes(b"\x00" * (2 * 1024 * 1024))
    respx.post(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createUploadSession",
    ).respond(412, json={"error": {"code": "preconditionFailed"}})
    with pytest.raises(StaleWriteError, match="changed under us"):
        save(URL, comment="big stale", version="minor")
    # Registry NOT cleared on stale-write
    assert CheckoutRegistry(profile="default", base_dir=tmp_path).get(URL) is not None
