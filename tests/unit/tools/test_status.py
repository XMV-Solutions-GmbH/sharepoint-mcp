# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_status."""

from __future__ import annotations

from pathlib import Path

import pytest

from sharepoint_mcp.checkout_registry import CheckedOutEntry, CheckoutRegistry
from sharepoint_mcp.tools.status import status


def _seed(tmp_path: Path, profile: str, entries: list[CheckedOutEntry]) -> None:
    registry = CheckoutRegistry(profile=profile, base_dir=tmp_path)
    for entry in entries:
        registry.add(entry)


def test_status_empty_when_nothing_checked_out(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    assert status(profile="empty") == []


def test_status_returns_one_entry_per_checkout(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    _seed(
        tmp_path,
        "default",
        [
            CheckedOutEntry(
                path="https://example/foo.docx",
                site_id="S1",
                drive_id="D1",
                item_id="I1",
                local_path="/tmp/wc/foo.docx",
                etag="v1",
                since=1_900_000_000.0,
            ),
        ],
    )
    result = status(profile="default")
    assert len(result) == 1
    entry = result[0]
    assert entry["path"] == "https://example/foo.docx"
    assert entry["local_path"] == "/tmp/wc/foo.docx"
    # since is rendered as ISO datetime
    assert "2030" in entry["since"]  # 1_900_000_000 epoch ≈ 2030-03


def test_status_per_profile_isolation(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    _seed(
        tmp_path,
        "alpha",
        [
            CheckedOutEntry(
                path="A",
                site_id="s",
                drive_id="d",
                item_id="i",
                local_path="/x/A",
                etag="e",
                since=0.0,
            ),
        ],
    )
    _seed(
        tmp_path,
        "beta",
        [
            CheckedOutEntry(
                path="B",
                site_id="s",
                drive_id="d",
                item_id="i",
                local_path="/x/B",
                etag="e",
                since=0.0,
            ),
        ],
    )
    assert [e["path"] for e in status(profile="alpha")] == ["A"]
    assert [e["path"] for e in status(profile="beta")] == ["B"]


def test_status_does_not_leak_etag_or_internal_ids(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    """Public response keeps only the user-facing fields. ETag / drive_id /
    item_id are internal plumbing for sp_save and shouldn't leak via sp_status.
    """
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    _seed(
        tmp_path,
        "default",
        [
            CheckedOutEntry(
                path="P",
                site_id="S",
                drive_id="D",
                item_id="I",
                local_path="/L",
                etag="ETAG",
                since=0.0,
            ),
        ],
    )
    [entry] = status(profile="default")
    assert set(entry.keys()) == {"path", "since", "local_path"}
