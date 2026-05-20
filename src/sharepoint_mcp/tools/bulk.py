# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Bulk variants of sp_drive_file_checkout and sp_drive_file_checkin (closes #41).

Use case: agent edits 20 policy documents that all need the same
header update; doing it one at a time is 60 round-trips. The bulk
tools dispatch up to 4 underlying operations in parallel, respecting
Microsoft Graph throttling guidance.

Semantics (per #41):

- Per-item failures are reported in the result list, NOT raised. The
  caller decides whether to retry / continue / abort. No transactional
  semantics — Graph doesn't offer transactions across files, and we
  don't fake them.
- 429 / 503 responses with `Retry-After` headers are honoured (up to 2
  retries per item, max 30s wait per attempt).
- The result list preserves input order so callers can correlate by
  index without re-parsing URLs.

Result shape per item: `{"path": str, "status": "ok" | "error", ...}`
where the `...` is operation-specific (`local_path` for open,
`version_id` / `etag` / `web_url` for save) on success, or `error`
(human-readable string) on failure.
"""

from __future__ import annotations

import time
from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor
from typing import Any, TypedDict

import httpx

from sharepoint_mcp.tools.open_file import open_file as _do_open
from sharepoint_mcp.tools.save import save as _do_save

DEFAULT_CONCURRENCY = 4
MAX_RETRIES_ON_THROTTLE = 2
MAX_RETRY_WAIT_SECONDS = 30


class SaveOperation(TypedDict, total=False):
    """One save in a sp_drive_file_checkin_bulk batch.

    `version` is optional and defaults to "minor" — same default as
    sp_drive_file_checkin's single-call form. `comment` is required by sp_drive_file_checkin.
    """

    url: str
    comment: str
    version: str  # "minor" | "major"


def open_many(
    urls: list[str],
    *,
    profile: str = "default",
    concurrency: int = DEFAULT_CONCURRENCY,
    sleep: Callable[[float], None] = time.sleep,
) -> list[dict[str, Any]]:
    """Bulk variant of sp_drive_file_checkout. See module docstring for semantics.

    Returns one result per input url, in the original order. Each result
    is `{"path": <input url>, "status": "ok", "local_path": <str>}`
    or `{"path": <input url>, "status": "error", "error": <str>}`.

    Empty `urls` returns an empty list without any work.
    """
    if not urls:
        return []
    if concurrency < 1:
        raise ValueError(f"concurrency must be >= 1, got {concurrency!r}")

    def _work(url: str) -> dict[str, Any]:
        try:
            local = _retry_on_throttle(
                lambda: _do_open(url, profile=profile),
                sleep=sleep,
            )
            return {"path": url, "status": "ok", "local_path": local}
        except Exception as exc:
            return _error_result(url, exc)

    return _dispatch(urls, _work, concurrency=concurrency)


def save_many(
    operations: list[SaveOperation],
    *,
    profile: str = "default",
    concurrency: int = DEFAULT_CONCURRENCY,
    sleep: Callable[[float], None] = time.sleep,
) -> list[dict[str, Any]]:
    """Bulk variant of sp_drive_file_checkin. See module docstring for semantics.

    Each `operations` entry: `{"url": str, "comment": str, "version"?: "minor"|"major"}`.

    Returns one result per input op, in the original order. Each result
    is the sp_drive_file_checkin dict (`version_id`, `etag`, `web_url`) merged with
    `path` + `status="ok"`, or `{path, status="error", error}` on failure.

    Empty `operations` returns an empty list without any work.
    """
    if not operations:
        return []
    if concurrency < 1:
        raise ValueError(f"concurrency must be >= 1, got {concurrency!r}")

    def _work(op: SaveOperation) -> dict[str, Any]:
        url = op.get("url", "")
        comment = op.get("comment", "")
        version_raw = op.get("version", "minor")
        try:
            if version_raw not in ("minor", "major"):
                raise ValueError(
                    f"version must be 'minor' or 'major', got {version_raw!r}",
                )
            result = _retry_on_throttle(
                lambda: _do_save(
                    url,
                    comment=comment,
                    version=version_raw,  # type: ignore[arg-type]
                    profile=profile,
                ),
                sleep=sleep,
            )
            return {"path": url, "status": "ok", **result}
        except Exception as exc:
            return _error_result(url, exc)

    return _dispatch(operations, _work, concurrency=concurrency)


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------


def _dispatch(
    items: list[Any],
    work: Callable[[Any], dict[str, Any]],
    *,
    concurrency: int,
) -> list[dict[str, Any]]:
    """Run `work` over `items` with bounded concurrency, preserving input order.

    ThreadPoolExecutor.map preserves input order in the iterable it
    yields, so we can collect into a list directly.
    """
    if concurrency == 1 or len(items) == 1:
        return [work(it) for it in items]
    with ThreadPoolExecutor(max_workers=concurrency) as ex:
        return list(ex.map(work, items))


def _retry_on_throttle(
    call: Callable[[], Any],
    *,
    sleep: Callable[[float], None],
) -> Any:
    """Invoke `call`, retrying up to MAX_RETRIES_ON_THROTTLE times on 429/503.

    Honours `Retry-After` from the response (seconds, integer). Falls
    back to exponential backoff (2, 4 seconds) when the header is
    absent or unparseable. Any other exception propagates immediately.
    """
    for attempt in range(MAX_RETRIES_ON_THROTTLE + 1):
        try:
            return call()
        except httpx.HTTPStatusError as exc:
            if exc.response.status_code not in (429, 503):
                raise
            if attempt >= MAX_RETRIES_ON_THROTTLE:
                raise
            sleep(_compute_retry_wait(exc.response, attempt))
    # Unreachable — the loop always returns or raises — but mypy
    # demands a return path on the function.
    raise RuntimeError("retry loop exhausted without resolution")


def _compute_retry_wait(response: httpx.Response, attempt: int) -> float:
    """Compute seconds-to-wait for a throttled retry.

    Microsoft Graph populates `Retry-After` with a positive integer
    representing seconds. If absent or unparseable, fall back to
    exponential backoff: 2, 4 seconds. Cap at MAX_RETRY_WAIT_SECONDS
    so a buggy upstream that returns Retry-After: 86400 doesn't park
    the bulk operation for a day.
    """
    header = response.headers.get("Retry-After", "").strip()
    if header.isdigit():
        return float(min(int(header), MAX_RETRY_WAIT_SECONDS))
    fallback = float(2 ** (attempt + 1))
    return min(fallback, float(MAX_RETRY_WAIT_SECONDS))


def _error_result(url: str, exc: BaseException) -> dict[str, Any]:
    """Standard per-item error shape — name + message, never the traceback."""
    return {
        "path": url,
        "status": "error",
        "error": f"{type(exc).__name__}: {exc}",
    }
