# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for the resumable upload session helper (#38)."""

from __future__ import annotations

from pathlib import Path

import httpx
import pytest
import respx

from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools._upload import (
    CHUNK_SIZE_BYTES,
    DEFAULT_THRESHOLD_MB,
    StaleUploadSessionError,
    chunked_upload_threshold_bytes,
    upload_resumable,
)

DRIVE_ID = "b!drive"
ITEM_ID = "01ITEM"
UPLOAD_URL = "https://upload.example/session/abc123"


# ---------------------------------------------------------------------
# Threshold env-var parsing
# ---------------------------------------------------------------------


def test_threshold_default(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("SP_CHUNKED_UPLOAD_THRESHOLD_MB", raising=False)
    assert chunked_upload_threshold_bytes() == DEFAULT_THRESHOLD_MB * 1024 * 1024


def test_threshold_override(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("SP_CHUNKED_UPLOAD_THRESHOLD_MB", "5")
    assert chunked_upload_threshold_bytes() == 5 * 1024 * 1024


def test_threshold_invalid_falls_back_to_default(monkeypatch: pytest.MonkeyPatch) -> None:
    """Typo in env var must not silently disable chunked uploads."""
    monkeypatch.setenv("SP_CHUNKED_UPLOAD_THRESHOLD_MB", "twenty")
    assert chunked_upload_threshold_bytes() == DEFAULT_THRESHOLD_MB * 1024 * 1024


def test_threshold_empty_falls_back_to_default(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("SP_CHUNKED_UPLOAD_THRESHOLD_MB", "")
    assert chunked_upload_threshold_bytes() == DEFAULT_THRESHOLD_MB * 1024 * 1024


def test_chunk_size_is_multiple_of_320kib() -> None:
    """Microsoft requires chunk sizes that are multiples of 320 KiB."""
    assert CHUNK_SIZE_BYTES % (320 * 1024) == 0


# ---------------------------------------------------------------------
# Happy path: single chunk
# ---------------------------------------------------------------------


@respx.mock
def test_upload_resumable_single_chunk(tmp_path: Path) -> None:
    body = b"hello world" * 100
    local = tmp_path / "small.bin"
    local.write_bytes(body)

    respx.post(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createUploadSession",
    ).respond(json={"uploadUrl": UPLOAD_URL, "expirationDateTime": "2026-12-31T00:00:00Z"})
    respx.put(UPLOAD_URL).respond(
        201,
        json={"id": ITEM_ID, "eTag": "new-etag", "webUrl": "https://x/foo.bin"},
    )

    with httpx.Client() as client:
        result = upload_resumable(
            client,
            drive_id=DRIVE_ID,
            item_id=ITEM_ID,
            local_file=local,
            etag="prev-etag",
            auth_headers={"Authorization": "Bearer T"},
        )

    assert result["eTag"] == "new-etag"


@respx.mock
def test_upload_resumable_sends_if_match_on_session_creation(tmp_path: Path) -> None:
    local = tmp_path / "x.bin"
    local.write_bytes(b"x" * 1000)

    create_route = respx.post(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createUploadSession",
    ).respond(json={"uploadUrl": UPLOAD_URL})
    respx.put(UPLOAD_URL).respond(200, json={"id": ITEM_ID, "eTag": "e"})

    with httpx.Client() as client:
        upload_resumable(
            client,
            drive_id=DRIVE_ID,
            item_id=ITEM_ID,
            local_file=local,
            etag='"abc-123"',
            auth_headers={"Authorization": "Bearer T"},
        )

    request = create_route.calls.last.request
    assert request.headers.get("If-Match") == '"abc-123"'
    assert request.headers.get("Authorization") == "Bearer T"


@respx.mock
def test_upload_resumable_sends_correct_content_range(tmp_path: Path) -> None:
    body = b"a" * 1000
    local = tmp_path / "x.bin"
    local.write_bytes(body)

    respx.post(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createUploadSession",
    ).respond(json={"uploadUrl": UPLOAD_URL})
    chunk_route = respx.put(UPLOAD_URL).respond(200, json={"id": ITEM_ID, "eTag": "e"})

    with httpx.Client() as client:
        upload_resumable(
            client,
            drive_id=DRIVE_ID,
            item_id=ITEM_ID,
            local_file=local,
            etag="e0",
            auth_headers={"Authorization": "Bearer T"},
        )

    request = chunk_route.calls.last.request
    assert request.headers["Content-Range"] == "bytes 0-999/1000"
    assert request.headers["Content-Length"] == "1000"


@respx.mock
def test_upload_resumable_does_not_pass_authorization_to_uploadurl(tmp_path: Path) -> None:
    """Pre-signed uploadUrls should NOT receive bearer auth on chunk PUTs."""
    local = tmp_path / "x.bin"
    local.write_bytes(b"x" * 1000)
    respx.post(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createUploadSession",
    ).respond(json={"uploadUrl": UPLOAD_URL})
    chunk_route = respx.put(UPLOAD_URL).respond(200, json={"id": ITEM_ID, "eTag": "e"})

    with httpx.Client() as client:
        upload_resumable(
            client,
            drive_id=DRIVE_ID,
            item_id=ITEM_ID,
            local_file=local,
            etag="e0",
            auth_headers={"Authorization": "Bearer SECRET"},
        )

    assert "Authorization" not in chunk_route.calls.last.request.headers


# ---------------------------------------------------------------------
# Multi-chunk
# ---------------------------------------------------------------------


@respx.mock
def test_upload_resumable_multi_chunk(tmp_path: Path) -> None:
    """Uploads in chunks of `chunk_size`, sending each PUT in order."""
    chunk_size = 320 * 1024  # 320 KiB
    total = chunk_size * 3 + 100  # 4 chunks: 3 full + 1 partial
    local = tmp_path / "big.bin"
    local.write_bytes(b"\x00" * total)

    respx.post(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createUploadSession",
    ).respond(json={"uploadUrl": UPLOAD_URL})
    chunk_route = respx.put(UPLOAD_URL).mock(
        side_effect=[
            httpx.Response(202, json={"nextExpectedRanges": ["327680-"]}),
            httpx.Response(202, json={"nextExpectedRanges": ["655360-"]}),
            httpx.Response(202, json={"nextExpectedRanges": ["983040-"]}),
            httpx.Response(201, json={"id": ITEM_ID, "eTag": "final"}),
        ],
    )

    with httpx.Client() as client:
        result = upload_resumable(
            client,
            drive_id=DRIVE_ID,
            item_id=ITEM_ID,
            local_file=local,
            etag="e0",
            auth_headers={"Authorization": "Bearer T"},
            chunk_size=chunk_size,
        )

    assert result["eTag"] == "final"
    assert chunk_route.call_count == 4
    last_request = chunk_route.calls.last.request
    assert last_request.headers["Content-Range"] == f"bytes 983040-{total - 1}/{total}"


# ---------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------


@respx.mock
def test_upload_resumable_raises_stale_on_412(tmp_path: Path) -> None:
    local = tmp_path / "x.bin"
    local.write_bytes(b"x" * 100)
    respx.post(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createUploadSession",
    ).respond(412, json={"error": {"code": "preconditionFailed"}})

    with httpx.Client() as client, pytest.raises(StaleUploadSessionError, match="ETag mismatch"):
        upload_resumable(
            client,
            drive_id=DRIVE_ID,
            item_id=ITEM_ID,
            local_file=local,
            etag="stale",
            auth_headers={"Authorization": "Bearer T"},
        )


@respx.mock
def test_upload_resumable_propagates_other_4xx(tmp_path: Path) -> None:
    local = tmp_path / "x.bin"
    local.write_bytes(b"x" * 100)
    respx.post(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createUploadSession",
    ).respond(403, json={"error": {"code": "accessDenied"}})

    with httpx.Client() as client, pytest.raises(httpx.HTTPStatusError):
        upload_resumable(
            client,
            drive_id=DRIVE_ID,
            item_id=ITEM_ID,
            local_file=local,
            etag="e0",
            auth_headers={"Authorization": "Bearer T"},
        )


@respx.mock
def test_upload_resumable_retries_chunk_on_5xx(tmp_path: Path) -> None:
    local = tmp_path / "x.bin"
    local.write_bytes(b"x" * 1000)
    respx.post(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createUploadSession",
    ).respond(json={"uploadUrl": UPLOAD_URL})
    waits: list[float] = []
    chunk_route = respx.put(UPLOAD_URL).mock(
        side_effect=[
            httpx.Response(503),
            httpx.Response(503),
            httpx.Response(201, json={"id": ITEM_ID, "eTag": "ok"}),
        ],
    )
    with httpx.Client() as client:
        result = upload_resumable(
            client,
            drive_id=DRIVE_ID,
            item_id=ITEM_ID,
            local_file=local,
            etag="e0",
            auth_headers={"Authorization": "Bearer T"},
            sleep=waits.append,
        )
    assert result["eTag"] == "ok"
    assert chunk_route.call_count == 3
    assert waits == [2.0, 4.0]


@respx.mock
def test_upload_resumable_gives_up_after_max_chunk_retries(tmp_path: Path) -> None:
    local = tmp_path / "x.bin"
    local.write_bytes(b"x" * 1000)
    respx.post(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createUploadSession",
    ).respond(json={"uploadUrl": UPLOAD_URL})
    respx.put(UPLOAD_URL).respond(503)

    waits: list[float] = []
    with httpx.Client() as client, pytest.raises(httpx.HTTPStatusError):
        upload_resumable(
            client,
            drive_id=DRIVE_ID,
            item_id=ITEM_ID,
            local_file=local,
            etag="e0",
            auth_headers={"Authorization": "Bearer T"},
            sleep=waits.append,
        )
    # MAX_CHUNK_RETRIES=3 → 1 initial + 3 retries; so 3 sleeps before final raise.
    assert len(waits) == 3


@respx.mock
def test_upload_resumable_retries_chunk_on_connection_error(tmp_path: Path) -> None:
    local = tmp_path / "x.bin"
    local.write_bytes(b"x" * 1000)
    respx.post(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createUploadSession",
    ).respond(json={"uploadUrl": UPLOAD_URL})

    waits: list[float] = []
    chunk_route = respx.put(UPLOAD_URL).mock(
        side_effect=[
            httpx.ConnectError("boom"),
            httpx.Response(201, json={"id": ITEM_ID, "eTag": "ok"}),
        ],
    )
    with httpx.Client() as client:
        result = upload_resumable(
            client,
            drive_id=DRIVE_ID,
            item_id=ITEM_ID,
            local_file=local,
            etag="e0",
            auth_headers={"Authorization": "Bearer T"},
            sleep=waits.append,
        )
    assert result["eTag"] == "ok"
    assert chunk_route.call_count == 2
    assert waits == [2.0]


@respx.mock
def test_upload_resumable_propagates_4xx_on_chunk(tmp_path: Path) -> None:
    """Non-5xx status from a chunk PUT must not retry; propagate immediately."""
    local = tmp_path / "x.bin"
    local.write_bytes(b"x" * 1000)
    respx.post(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createUploadSession",
    ).respond(json={"uploadUrl": UPLOAD_URL})
    chunk_route = respx.put(UPLOAD_URL).respond(403)

    waits: list[float] = []
    with httpx.Client() as client, pytest.raises(httpx.HTTPStatusError):
        upload_resumable(
            client,
            drive_id=DRIVE_ID,
            item_id=ITEM_ID,
            local_file=local,
            etag="e0",
            auth_headers={"Authorization": "Bearer T"},
            sleep=waits.append,
        )
    assert chunk_route.call_count == 1  # no retry
    assert waits == []


def test_upload_resumable_rejects_empty_file(tmp_path: Path) -> None:
    """Empty file via resumable session is undefined behaviour; reject explicitly."""
    local = tmp_path / "empty.bin"
    local.write_bytes(b"")
    with httpx.Client() as client:
        with respx.mock:
            respx.post(
                f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createUploadSession",
            ).respond(json={"uploadUrl": UPLOAD_URL})
            with pytest.raises(ValueError, match="empty file"):
                upload_resumable(
                    client,
                    drive_id=DRIVE_ID,
                    item_id=ITEM_ID,
                    local_file=local,
                    etag="e0",
                    auth_headers={"Authorization": "Bearer T"},
                )
