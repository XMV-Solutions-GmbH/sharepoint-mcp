# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_drive_checkout_list."""

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
from sharepoint_mcp.tools.status import _query_lock_state, status


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
    item_id are internal plumbing for sp_drive_file_checkin and shouldn't
    leak via sp_drive_checkout_list.
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


# ---------------------------------------------------------------------
# verify=True
# ---------------------------------------------------------------------


def _entry(
    path: str = "https://example/foo.docx",
    drive_id: str = "D1",
    item_id: str = "I1",
) -> CheckedOutEntry:
    return CheckedOutEntry(
        path=path,
        site_id="S1",
        drive_id=drive_id,
        item_id=item_id,
        local_path="/tmp/wc/foo.docx",
        etag="v1",
        since=1_900_000_000.0,
    )


def test_status_verify_empty_skips_graph(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    store_with_fresh_token: None,
) -> None:
    """verify=True with no entries must not require any Graph calls."""
    del store_with_fresh_token
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    # No mock route; if the code tried to hit Graph we'd get an error.
    assert status(profile="empty", verify=True) == []


@respx.mock
def test_status_verify_reports_locked_when_checkout_user_present(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    store_with_fresh_token: None,
) -> None:
    del store_with_fresh_token
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    _seed(tmp_path, "default", [_entry()])
    respx.get(f"{GRAPH_BASE}/drives/D1/items/I1/listItem").respond(
        json={"fields": {"CheckoutUser": "David Koller"}},
    )
    [out] = status(profile="default", verify=True)
    assert out["server_locked"] is True
    assert out["lock_holder"] == "David Koller"


@respx.mock
def test_status_verify_reports_unlocked_when_checkout_user_absent(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    store_with_fresh_token: None,
) -> None:
    del store_with_fresh_token
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    _seed(tmp_path, "default", [_entry()])
    respx.get(f"{GRAPH_BASE}/drives/D1/items/I1/listItem").respond(json={"fields": {}})
    [out] = status(profile="default", verify=True)
    assert out["server_locked"] is False
    assert out["lock_holder"] is None


@respx.mock
def test_status_verify_returns_none_on_404(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    store_with_fresh_token: None,
) -> None:
    """Item deleted server-side → server_locked indeterminate, not crash."""
    del store_with_fresh_token
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    _seed(tmp_path, "default", [_entry()])
    respx.get(f"{GRAPH_BASE}/drives/D1/items/I1/listItem").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    [out] = status(profile="default", verify=True)
    assert out["server_locked"] is None
    assert out["lock_holder"] is None


@respx.mock
def test_status_verify_returns_none_on_network_error(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    store_with_fresh_token: None,
) -> None:
    del store_with_fresh_token
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    _seed(tmp_path, "default", [_entry()])
    respx.get(f"{GRAPH_BASE}/drives/D1/items/I1/listItem").mock(
        side_effect=httpx.ConnectError("boom")
    )
    [out] = status(profile="default", verify=True)
    assert out["server_locked"] is None
    assert out["lock_holder"] is None


@respx.mock
def test_status_verify_one_call_per_entry(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    store_with_fresh_token: None,
) -> None:
    del store_with_fresh_token
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    _seed(
        tmp_path,
        "default",
        [
            _entry(path="https://example/a.docx", drive_id="DA", item_id="IA"),
            _entry(path="https://example/b.docx", drive_id="DB", item_id="IB"),
        ],
    )
    route_a = respx.get(f"{GRAPH_BASE}/drives/DA/items/IA/listItem").respond(
        json={"fields": {"CheckoutUser": "Alice"}},
    )
    route_b = respx.get(f"{GRAPH_BASE}/drives/DB/items/IB/listItem").respond(
        json={"fields": {}},
    )
    out = status(profile="default", verify=True)
    assert route_a.called and route_a.call_count == 1
    assert route_b.called and route_b.call_count == 1
    locks = {e["path"]: e["server_locked"] for e in out}
    assert locks == {
        "https://example/a.docx": True,
        "https://example/b.docx": False,
    }


@respx.mock
def test_status_verify_dict_checkout_user(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    store_with_fresh_token: None,
) -> None:
    """Some Graph payloads return CheckoutUser as a dict with displayName."""
    del store_with_fresh_token
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    _seed(tmp_path, "default", [_entry()])
    respx.get(f"{GRAPH_BASE}/drives/D1/items/I1/listItem").respond(
        json={
            "fields": {
                "CheckoutUser": {
                    "displayName": "Bob",
                    "email": "bob@example.com",
                }
            }
        },
    )
    [out] = status(profile="default", verify=True)
    assert out["server_locked"] is True
    assert out["lock_holder"] == "Bob"


@respx.mock
def test_status_verify_list_lookup_value_checkout_user(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    store_with_fresh_token: None,
) -> None:
    """SharePoint list-fields wire shape uses [{"LookupValue": "..."}]."""
    del store_with_fresh_token
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    _seed(tmp_path, "default", [_entry()])
    respx.get(f"{GRAPH_BASE}/drives/D1/items/I1/listItem").respond(
        json={"fields": {"CheckoutUser": [{"LookupValue": "Carol"}]}},
    )
    [out] = status(profile="default", verify=True)
    assert out["server_locked"] is True
    assert out["lock_holder"] == "Carol"


@respx.mock
def test_status_default_skips_graph(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    store_with_fresh_token: None,
) -> None:
    """verify=False (default) must not hit Graph at all."""
    del store_with_fresh_token
    monkeypatch.setattr("sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR", tmp_path)
    _seed(tmp_path, "default", [_entry()])
    route = respx.get(f"{GRAPH_BASE}/drives/D1/items/I1/listItem").respond(
        json={"fields": {"CheckoutUser": "anyone"}},
    )
    [out] = status(profile="default")
    assert "server_locked" not in out
    assert "lock_holder" not in out
    assert not route.called


# ---------------------------------------------------------------------
# Helper directly
# ---------------------------------------------------------------------


@respx.mock
def test_query_lock_state_returns_locked_dict_with_lookup_value() -> None:
    respx.get(f"{GRAPH_BASE}/drives/D/items/I/listItem").respond(
        json={"fields": {"CheckoutUser": {"LookupValue": "Dave"}}},
    )
    with httpx.Client() as c:
        assert _query_lock_state(c, drive_id="D", item_id="I", headers={}) == (True, "Dave")


@respx.mock
def test_query_lock_state_unknown_user_object_shape_still_signals_locked() -> None:
    """Don't break if Graph returns an unrecognised CheckoutUser shape — at
    minimum we know there IS a lock."""
    respx.get(f"{GRAPH_BASE}/drives/D/items/I/listItem").respond(
        json={"fields": {"CheckoutUser": {"someOtherField": "value"}}},
    )
    with httpx.Client() as c:
        result = _query_lock_state(c, drive_id="D", item_id="I", headers={})
    assert result[0] is True
    assert result[1] is None
