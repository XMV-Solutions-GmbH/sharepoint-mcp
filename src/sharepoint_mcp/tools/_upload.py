# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Resumable upload session helper for sp_save_file (closes #38).

Microsoft Graph caps single-shot `PUT /content` at 250 MB per file
on SharePoint Online. For files above the threshold (default 100 MB,
overridable via `SP_CHUNKED_UPLOAD_THRESHOLD_MB`), `save()` uses a
resumable upload session: `POST /createUploadSession` returns a
pre-signed `uploadUrl`, then chunks are uploaded via `PUT` against
that URL with `Content-Range` headers.

Microsoft documents the chunk size must be a multiple of 320 KiB,
and recommends 5-10 MB. We use 5 MiB (= 16 x 320 KiB), the smaller
end of the recommended range — less waste on retry, more progress
checkpoints.

The pre-signed `uploadUrl` is bearer-auth-free; chunk PUTs do NOT
include the `Authorization` header. Only the createUploadSession
call needs it.

Stale-write detection: the `If-Match` header is passed on the
createUploadSession call. SharePoint returns 412 if the eTag
doesn't match, identical to single-shot behaviour, and we raise
`StaleWriteError` to match.
"""

from __future__ import annotations

import os
import time
from collections.abc import Callable
from pathlib import Path
from typing import Any

import httpx

from sharepoint_mcp.tools._common import GRAPH_BASE

# Multiple of 320 KiB (327680 bytes) per Graph docs. 5 MiB = 16 x 320 KiB.
CHUNK_SIZE_BYTES = 5 * 1024 * 1024
DEFAULT_THRESHOLD_MB = 100
THRESHOLD_ENV = "SP_CHUNKED_UPLOAD_THRESHOLD_MB"
MAX_CHUNK_RETRIES = 3


class StaleUploadSessionError(RuntimeError):
    """ETag mismatch on createUploadSession — file changed under us."""


def chunked_upload_threshold_bytes() -> int:
    """Return the byte threshold above which save() uses resumable uploads.

    Reads `SP_CHUNKED_UPLOAD_THRESHOLD_MB` from the environment. Falls
    back to `DEFAULT_THRESHOLD_MB` when unset, empty, or non-numeric
    (we don't want a typo to silently disable chunked uploads).
    """
    raw = os.environ.get(THRESHOLD_ENV, "").strip()
    if raw.isdigit():
        return int(raw) * 1024 * 1024
    return DEFAULT_THRESHOLD_MB * 1024 * 1024


def upload_resumable(
    client: httpx.Client,
    *,
    drive_id: str,
    item_id: str,
    local_file: Path,
    etag: str,
    auth_headers: dict[str, str],
    chunk_size: int = CHUNK_SIZE_BYTES,
    sleep: Callable[[float], None] = time.sleep,
) -> dict[str, Any]:
    """Upload `local_file` to `/drives/{id}/items/{id}` via a resumable session.

    Returns the final driveItem JSON (parsed) — same shape as a
    single-shot `PUT /content` response, so callers can extract
    `eTag`, `webUrl`, etc. uniformly.

    Raises:
        StaleUploadSessionError: 412 on createUploadSession (ETag drift).
        httpx.HTTPStatusError: any other non-2xx response not retryable
            (or retryable failures past `MAX_CHUNK_RETRIES`).
    """
    session_response = client.post(
        f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/createUploadSession",
        headers={
            **auth_headers,
            "If-Match": etag,
            "Content-Type": "application/json",
        },
        json={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
    )
    if session_response.status_code == 412:
        raise StaleUploadSessionError(
            "Upload session refused: ETag mismatch (412). File changed "
            "under us between sp_open_file and sp_save_file.",
        )
    session_response.raise_for_status()
    upload_url = str(session_response.json()["uploadUrl"])

    total_size = local_file.stat().st_size
    if total_size == 0:
        # createUploadSession with a 0-byte body isn't well-defined;
        # caller should have used single-shot. We don't reach here in
        # practice because the threshold check is well above 0, but
        # be safe.
        raise ValueError("Cannot upload an empty file via resumable session")

    last_response: httpx.Response | None = None
    with local_file.open("rb") as f:
        offset = 0
        while offset < total_size:
            chunk = f.read(chunk_size)
            if not chunk:
                break
            last_response = _put_chunk_with_retry(
                client,
                upload_url=upload_url,
                chunk=chunk,
                offset=offset,
                total_size=total_size,
                sleep=sleep,
            )
            offset += len(chunk)

    if last_response is None:
        raise RuntimeError("Resumable upload completed without uploading any chunks")
    return _final_payload(last_response)


def _put_chunk_with_retry(
    client: httpx.Client,
    *,
    upload_url: str,
    chunk: bytes,
    offset: int,
    total_size: int,
    sleep: Callable[[float], None],
) -> httpx.Response:
    """PUT one chunk with retry on 5xx + transient httpx errors."""
    chunk_len = len(chunk)
    content_range = f"bytes {offset}-{offset + chunk_len - 1}/{total_size}"
    headers = {
        "Content-Range": content_range,
        "Content-Length": str(chunk_len),
    }
    last_exc: Exception | None = None
    for attempt in range(MAX_CHUNK_RETRIES + 1):
        try:
            response = client.put(upload_url, headers=headers, content=chunk)
        except httpx.HTTPError as exc:
            last_exc = exc
            if attempt >= MAX_CHUNK_RETRIES:
                raise
            sleep(_backoff_seconds(attempt))
            continue
        if response.status_code in (200, 201, 202):
            return response
        if 500 <= response.status_code < 600 and attempt < MAX_CHUNK_RETRIES:
            sleep(_backoff_seconds(attempt))
            continue
        response.raise_for_status()
        # Defensive fallthrough — raise_for_status returns on 2xx, but
        # we already filtered 200/201/202 above, so any other 2xx is
        # unexpected. Treat as success for compatibility.
        return response
    # Unreachable in practice — loop returns or raises — but mypy
    # demands a terminal statement.
    if last_exc is not None:
        raise last_exc
    raise RuntimeError("chunk upload exhausted retries without resolution")


def _final_payload(response: httpx.Response) -> dict[str, Any]:
    """Parse the final-chunk response. 200/201 carry the driveItem JSON;
    202 means the session is still open, which shouldn't happen on the
    last chunk (the math is exact). We raise rather than return an
    incomplete payload.
    """
    if response.status_code in (200, 201):
        return dict(response.json())
    raise RuntimeError(
        f"Resumable upload final chunk got status {response.status_code} "
        "but expected 200 or 201; session was not committed.",
    )


def _backoff_seconds(attempt: int) -> float:
    """2, 4, 8 seconds — exponential backoff for chunk retries."""
    return float(2 ** (attempt + 1))
