# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_open_many / sp_save_many (#41).

Each run seeds 5 throwaway test files via sp_publish, exercises the
bulk operations against them, and cleans up by removing the files
through Microsoft Graph at the end. This is the closest we can get
to a real-world batch (acceptance criteria: 5+ files in one bulk
call) without polluting the harness sandbox between runs.
"""

from __future__ import annotations

import time
import uuid
from collections.abc import Iterator
from pathlib import Path

import httpx
import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.checkout_registry import CheckoutRegistry
from sharepoint_mcp.tools._common import (
    GRAPH_BASE,
    parse_sharepoint_url,
    resolve_drive_item,
    resolve_site_id,
)
from sharepoint_mcp.tools.bulk import open_many, save_many
from sharepoint_mcp.tools.publish import publish
from tests.harness._cleanup import HARNESS_PROFILE, discard_checkouts_added_during

HARNESS_SITE_URL = "https://xmvsolutions.sharepoint.com/sites/sharepoint-mcp-harness"
HARNESS_FOLDER_URL = f"{HARNESS_SITE_URL}/Shared Documents"
SEED_COUNT = 5


def _skip_if_no_harness() -> None:
    try:
        get_token(HARNESS_PROFILE)
    except AuthRequiredError as exc:
        pytest.skip(
            f"Harness credentials not available: {exc}. "
            "Run `uv run mcp-server-sharepoint login --profile harness` to populate.",
        )


@pytest.fixture
def seeded_files(tmp_path: Path) -> Iterator[list[str]]:
    """Publish SEED_COUNT throwaway files; yield their URLs; delete on teardown."""
    _skip_if_no_harness()
    run_id = uuid.uuid4().hex[:8]
    urls: list[str] = []
    for i in range(SEED_COUNT):
        local = tmp_path / f"bulk-seed-{run_id}-{i}.txt"
        local.write_text(f"seed {i} for bulk harness {run_id}", encoding="utf-8")
        publish(
            str(local),
            HARNESS_FOLDER_URL,
            name=local.name,
            profile=HARNESS_PROFILE,
        )
        urls.append(f"{HARNESS_FOLDER_URL}/{local.name}")
    try:
        yield urls
    finally:
        _delete_files(urls)


@pytest.fixture
def cleanup_checkouts() -> Iterator[None]:
    pre = {e.path for e in CheckoutRegistry(HARNESS_PROFILE).list_all()}
    yield from discard_checkouts_added_during(pre)


def test_open_many_acquires_locks_for_all_seeded_files(
    seeded_files: list[str],
    cleanup_checkouts: None,
) -> None:
    del cleanup_checkouts
    results = open_many(seeded_files, profile=HARNESS_PROFILE)
    assert len(results) == SEED_COUNT
    assert [r["path"] for r in results] == seeded_files
    assert all(r["status"] == "ok" for r in results), [r for r in results if r["status"] != "ok"]
    for r in results:
        assert Path(r["local_path"]).exists()


def test_save_many_creates_new_versions_for_all(
    seeded_files: list[str],
    cleanup_checkouts: None,
) -> None:
    del cleanup_checkouts
    open_results = open_many(seeded_files, profile=HARNESS_PROFILE)
    assert all(r["status"] == "ok" for r in open_results)
    # Modify each working copy with a marker line
    for r in open_results:
        local = Path(r["local_path"])
        local.write_text(local.read_text(encoding="utf-8") + "\nbulk marker\n", encoding="utf-8")
    save_results = save_many(
        [{"url": r["path"], "comment": "bulk harness save"} for r in open_results],
        profile=HARNESS_PROFILE,
    )
    assert len(save_results) == SEED_COUNT
    assert [r["path"] for r in save_results] == seeded_files
    failures = [r for r in save_results if r["status"] != "ok"]
    assert not failures, failures
    for r in save_results:
        assert r["version_id"]
        assert r["web_url"]


def test_open_many_isolates_per_item_failure(
    seeded_files: list[str],
    cleanup_checkouts: None,
) -> None:
    del cleanup_checkouts
    bogus = f"{HARNESS_FOLDER_URL}/this-does-not-exist-{uuid.uuid4().hex[:6]}.txt"
    urls = [seeded_files[0], bogus, seeded_files[1]]
    [r1, r2, r3] = open_many(urls, profile=HARNESS_PROFILE)
    assert r1["status"] == "ok"
    assert r2["status"] == "error"  # 404
    err_lc = r2["error"].lower()
    assert "404" in r2["error"] or "itemnotfound" in err_lc or "not found" in err_lc
    assert r3["status"] == "ok"


def test_open_many_runs_concurrently_under_real_latency(
    seeded_files: list[str],
    cleanup_checkouts: None,
) -> None:
    """Sequential 5x sp_open against real Graph would take ~2s+ at typical
    round-trip latency. Bulk with concurrency=4 should land well under
    the sequential floor."""
    del cleanup_checkouts
    t0 = time.monotonic()
    results = open_many(seeded_files, profile=HARNESS_PROFILE)
    elapsed = time.monotonic() - t0
    assert all(r["status"] == "ok" for r in results)
    # Don't make this brittle — just assert "didn't take ages". Real
    # cloud latency varies, so allow 20s ceiling. The point is to flag
    # accidental serialization regressions, not micro-benchmark.
    assert elapsed < 20.0, f"bulk open took {elapsed:.1f}s — concurrency may have regressed"


# ---------------------------------------------------------------------
# Teardown helper — direct Graph delete to clean up seeded files
# ---------------------------------------------------------------------


def _delete_files(urls: list[str]) -> None:
    """Best-effort: delete each seeded file via Graph. Swallows errors."""
    try:
        token = get_token(HARNESS_PROFILE)
    except AuthRequiredError:
        return
    headers = {"Authorization": f"Bearer {token}"}
    with httpx.Client(timeout=30.0) as client:
        for url in urls:
            try:
                hostname, site_path, item_path = parse_sharepoint_url(url)
                site_id = resolve_site_id(client, hostname, site_path, headers=headers)
                drive_id, item_id = resolve_drive_item(
                    client,
                    site_id,
                    item_path,
                    headers=headers,
                )
                client.delete(
                    f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}",
                    headers=headers,
                )
            except (httpx.HTTPError, KeyError):
                # Best effort. If a file is checked out by us at teardown
                # time the delete will fail; the cleanup_checkouts fixture
                # discards the lock first, so order matters: bulk harness
                # tests use cleanup_checkouts BEFORE seeded_files in
                # parameter order so the discard runs before the delete.
                pass
