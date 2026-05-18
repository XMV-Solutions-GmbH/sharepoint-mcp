# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_release_file."""

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
from sharepoint_mcp.tools.release import release


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


def _seed(tmp_path: Path) -> Path:
    work_dir = tmp_path / "default" / "working" / ITEM_ID
    work_dir.mkdir(parents=True, exist_ok=True)
    work_file = work_dir / "policy.docx"
    work_file.write_bytes(b"local-edits")
    CheckoutRegistry(profile="default", base_dir=tmp_path).add(
        CheckedOutEntry(
            path=URL,
            site_id="S1",
            drive_id=DRIVE_ID,
            item_id=ITEM_ID,
            local_path=str(work_file),
            etag='"abc"',
            since=1_900_000_000.0,
        ),
    )
    return work_file


@pytest.fixture
def seeded_registry(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    return _seed(tmp_path)


# ---------------------------------------------------------------------
# Happy path
# ---------------------------------------------------------------------


@respx.mock
def test_release_calls_discard_and_cleans_up_locally(
    store_with_fresh_token: None, seeded_registry: Path, tmp_path: Path
) -> None:
    del store_with_fresh_token
    work_file = seeded_registry
    discard_route = respx.post(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/discardCheckout",
    ).respond(204)

    release(URL)

    assert discard_route.called
    assert CheckoutRegistry(profile="default", base_dir=tmp_path).get(URL) is None
    assert not work_file.exists()


@respx.mock
def test_release_no_op_when_nothing_registered(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    # Should not call Graph at all; should not raise
    release(URL)
    assert not respx.routes


@respx.mock
def test_release_idempotent_after_success(
    store_with_fresh_token: None, seeded_registry: Path
) -> None:
    """Calling release twice in a row succeeds; second call is no-op."""
    del store_with_fresh_token, seeded_registry
    respx.post(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/discardCheckout",
    ).respond(204)
    release(URL)  # first
    release(URL)  # second — registry empty now, returns silently


# ---------------------------------------------------------------------
# Server-side error: local cleanup still happens
# ---------------------------------------------------------------------


@respx.mock
def test_release_local_cleanup_runs_even_if_server_errors(
    store_with_fresh_token: None, seeded_registry: Path, tmp_path: Path
) -> None:
    del store_with_fresh_token
    work_file = seeded_registry
    respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/discardCheckout").respond(500)

    with pytest.raises(httpx.HTTPStatusError):
        release(URL)

    # Even though the server-side discard failed, the local registry
    # entry and working file are cleaned up — the user can re-discard
    # the server-side lock manually if needed.
    assert CheckoutRegistry(profile="default", base_dir=tmp_path).get(URL) is None
    assert not work_file.exists()


# ---------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------


def test_release_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        release("")


def test_release_rejects_blank_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        release("   ")
