# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_drive_file_checkout_bulk / sp_drive_file_checkin_bulk (#41).

We test the dispatcher / retry behaviour by injecting fake worker
functions, NOT by going all the way through respx-mocked open_file/
save flows. The underlying tools are already covered in their own
test modules; what's new here is concurrency, ordering, error
isolation, and Retry-After handling.
"""

from __future__ import annotations

import threading
import time
from typing import Any

import httpx
import pytest

from sharepoint_mcp.tools import bulk
from sharepoint_mcp.tools.bulk import (
    _compute_retry_wait,
    _retry_on_throttle,
    open_many,
    save_many,
)

# ---------------------------------------------------------------------
# open_many
# ---------------------------------------------------------------------


def test_open_many_empty_returns_empty() -> None:
    assert open_many([]) == []


def test_open_many_invalid_concurrency_raises() -> None:
    with pytest.raises(ValueError, match="concurrency must be >= 1"):
        open_many(["https://x"], concurrency=0)


def test_open_many_preserves_input_order(monkeypatch: pytest.MonkeyPatch) -> None:
    """Even if workers complete out of order, results match input positions."""

    def fake_open(url: str, *, profile: str = "default") -> str:
        del profile
        # Make later items finish first to challenge ordering.
        time.sleep(0.05 if "first" in url else 0.0)
        return f"/local/{url.rsplit('/', 1)[-1]}"

    monkeypatch.setattr(bulk, "_do_open", fake_open)
    urls = [
        "https://example/first.docx",
        "https://example/second.docx",
        "https://example/third.docx",
    ]
    out = open_many(urls)
    assert [r["path"] for r in out] == urls
    assert [r["status"] for r in out] == ["ok", "ok", "ok"]
    assert out[0]["local_path"].endswith("first.docx")
    assert out[2]["local_path"].endswith("third.docx")


def test_open_many_isolates_failures(monkeypatch: pytest.MonkeyPatch) -> None:
    """One failing item must not abort the rest."""

    def fake_open(url: str, *, profile: str = "default") -> str:
        del profile
        if "broken" in url:
            raise RuntimeError("simulated failure")
        return f"/local/{url.rsplit('/', 1)[-1]}"

    monkeypatch.setattr(bulk, "_do_open", fake_open)
    urls = [
        "https://example/ok1.docx",
        "https://example/broken.docx",
        "https://example/ok2.docx",
    ]
    [r1, r2, r3] = open_many(urls)
    assert r1["status"] == "ok"
    assert r2["status"] == "error"
    assert "simulated failure" in r2["error"]
    assert r3["status"] == "ok"


def test_open_many_runs_in_parallel(monkeypatch: pytest.MonkeyPatch) -> None:
    """4 items at 100ms each should finish in well under 400ms with concurrency=4."""

    def fake_open(url: str, *, profile: str = "default") -> str:
        del profile
        time.sleep(0.1)
        return f"/local/{url}"

    monkeypatch.setattr(bulk, "_do_open", fake_open)
    urls = [f"https://example/{i}.docx" for i in range(4)]
    t0 = time.monotonic()
    out = open_many(urls, concurrency=4)
    elapsed = time.monotonic() - t0
    assert all(r["status"] == "ok" for r in out)
    # Sequential would be ~0.4s; parallel should land near 0.1s. Allow
    # generous slack so a slow CI runner doesn't flake — anything under
    # 0.3s proves parallelism is happening.
    assert elapsed < 0.3, f"expected parallel execution, got {elapsed:.3f}s"


def test_open_many_concurrency_one_runs_sequentially(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """concurrency=1 should bypass the executor entirely."""
    threads: set[int] = set()

    def fake_open(url: str, *, profile: str = "default") -> str:
        del profile
        threads.add(threading.get_ident())
        return f"/local/{url}"

    monkeypatch.setattr(bulk, "_do_open", fake_open)
    open_many(["https://x", "https://y"], concurrency=1)
    assert threads == {threading.get_ident()}


# ---------------------------------------------------------------------
# save_many
# ---------------------------------------------------------------------


def test_save_many_empty_returns_empty() -> None:
    assert save_many([]) == []


def test_save_many_passes_comment_and_version(monkeypatch: pytest.MonkeyPatch) -> None:
    seen: list[dict[str, Any]] = []

    def fake_save(
        url: str,
        *,
        comment: str,
        version: str = "minor",
        profile: str = "default",
    ) -> dict[str, str]:
        del profile
        seen.append({"url": url, "comment": comment, "version": version})
        return {"version_id": "1.0", "etag": "e", "web_url": url}

    monkeypatch.setattr(bulk, "_do_save", fake_save)
    ops: list[bulk.SaveOperation] = [
        {"url": "https://x/a", "comment": "fix typo", "version": "minor"},
        {"url": "https://x/b", "comment": "publish", "version": "major"},
        {"url": "https://x/c", "comment": "tweak"},  # version defaults to minor
    ]
    out = save_many(ops, concurrency=1)
    assert [o["status"] for o in out] == ["ok", "ok", "ok"]
    seen_by_url = {s["url"]: s for s in seen}
    assert seen_by_url["https://x/a"]["version"] == "minor"
    assert seen_by_url["https://x/b"]["version"] == "major"
    assert seen_by_url["https://x/c"]["version"] == "minor"


def test_save_many_rejects_invalid_version_per_item(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Bad version on one op surfaces as that item's error, doesn't crash batch."""

    def fake_save(*args: Any, **kwargs: Any) -> dict[str, str]:
        del args, kwargs
        return {"version_id": "1.0", "etag": "e", "web_url": "x"}

    monkeypatch.setattr(bulk, "_do_save", fake_save)
    ops: list[bulk.SaveOperation] = [
        {"url": "https://x/a", "comment": "ok", "version": "minor"},
        {"url": "https://x/b", "comment": "bad", "version": "huge"},
    ]
    [r1, r2] = save_many(ops, concurrency=1)
    assert r1["status"] == "ok"
    assert r2["status"] == "error"
    assert "version must be 'minor' or 'major'" in r2["error"]


def test_save_many_merges_save_result_into_response(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    def fake_save(*args: Any, **kwargs: Any) -> dict[str, str]:
        return {"version_id": "3.0", "etag": "e3", "web_url": "https://x/a"}

    monkeypatch.setattr(bulk, "_do_save", fake_save)
    [out] = save_many(
        [{"url": "https://x/a", "comment": "msg"}],
        concurrency=1,
    )
    assert out["status"] == "ok"
    assert out["path"] == "https://x/a"
    assert out["version_id"] == "3.0"
    assert out["etag"] == "e3"
    assert out["web_url"] == "https://x/a"


# ---------------------------------------------------------------------
# Retry-on-throttle
# ---------------------------------------------------------------------


def _make_response(status: int, retry_after: str | None = None) -> httpx.Response:
    headers = {}
    if retry_after is not None:
        headers["Retry-After"] = retry_after
    return httpx.Response(
        status_code=status,
        headers=headers,
        request=httpx.Request("GET", "https://example/x"),
    )


def test_retry_on_throttle_retries_429_and_succeeds() -> None:
    calls = {"n": 0}
    waits: list[float] = []

    def call() -> str:
        calls["n"] += 1
        if calls["n"] < 3:
            raise httpx.HTTPStatusError(
                "429", request=_make_response(429, "1").request, response=_make_response(429, "1")
            )
        return "ok"

    result = _retry_on_throttle(call, sleep=waits.append)
    assert result == "ok"
    assert calls["n"] == 3
    assert waits == [1.0, 1.0]


def test_retry_on_throttle_gives_up_after_max_retries() -> None:
    waits: list[float] = []
    response = _make_response(429, "1")

    def always_throttle() -> None:
        raise httpx.HTTPStatusError("429", request=response.request, response=response)

    with pytest.raises(httpx.HTTPStatusError):
        _retry_on_throttle(always_throttle, sleep=waits.append)
    # MAX_RETRIES_ON_THROTTLE = 2 → 1 initial + 2 retries = 3 calls,
    # so 2 sleeps before the final raise.
    assert len(waits) == 2


def test_retry_on_throttle_propagates_other_status_codes_immediately() -> None:
    waits: list[float] = []
    response = _make_response(500)

    def fails_500() -> None:
        raise httpx.HTTPStatusError("500", request=response.request, response=response)

    with pytest.raises(httpx.HTTPStatusError):
        _retry_on_throttle(fails_500, sleep=waits.append)
    assert waits == []  # never slept


def test_retry_on_throttle_propagates_non_http_exceptions() -> None:
    waits: list[float] = []

    def raises_value_error() -> None:
        raise ValueError("nope")

    with pytest.raises(ValueError, match="nope"):
        _retry_on_throttle(raises_value_error, sleep=waits.append)
    assert waits == []


def test_retry_on_throttle_handles_503_too() -> None:
    calls = {"n": 0}
    waits: list[float] = []
    response = _make_response(503, "2")

    def call() -> str:
        calls["n"] += 1
        if calls["n"] < 2:
            raise httpx.HTTPStatusError("503", request=response.request, response=response)
        return "recovered"

    assert _retry_on_throttle(call, sleep=waits.append) == "recovered"
    assert waits == [2.0]


def test_compute_retry_wait_uses_retry_after_when_numeric() -> None:
    response = _make_response(429, "5")
    assert _compute_retry_wait(response, attempt=0) == 5.0


def test_compute_retry_wait_clamps_huge_retry_after() -> None:
    response = _make_response(429, "86400")  # one day
    assert _compute_retry_wait(response, attempt=0) == 30.0  # MAX_RETRY_WAIT_SECONDS


def test_compute_retry_wait_falls_back_to_exponential_on_non_numeric() -> None:
    response = _make_response(429, "soon")  # malformed header
    assert _compute_retry_wait(response, attempt=0) == 2.0  # 2^1
    assert _compute_retry_wait(response, attempt=1) == 4.0  # 2^2


def test_compute_retry_wait_falls_back_when_header_missing() -> None:
    response = _make_response(429)
    assert _compute_retry_wait(response, attempt=0) == 2.0


# ---------------------------------------------------------------------
# Bulk + retry integration
# ---------------------------------------------------------------------


def test_open_many_retries_throttled_item(monkeypatch: pytest.MonkeyPatch) -> None:
    counts = {"n": 0}
    waits: list[float] = []

    def fake_open(url: str, *, profile: str = "default") -> str:
        del profile
        counts["n"] += 1
        if counts["n"] == 1:
            response = _make_response(429, "1")
            raise httpx.HTTPStatusError("429", request=response.request, response=response)
        return f"/local/{url}"

    monkeypatch.setattr(bulk, "_do_open", fake_open)
    [out] = open_many(
        ["https://example/a"],
        concurrency=1,
        sleep=waits.append,
    )
    assert out["status"] == "ok"
    assert out["local_path"] == "/local/https://example/a"
    assert counts["n"] == 2
    assert waits == [1.0]
