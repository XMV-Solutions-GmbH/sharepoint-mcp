# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for the CheckoutRegistry — persistence + atomicity + dedup."""

from __future__ import annotations

from pathlib import Path

import pytest

from sharepoint_mcp.checkout_registry import CheckedOutEntry, CheckoutRegistry


def _entry(path: str = "https://example/foo", **overrides: str | float) -> CheckedOutEntry:
    base: dict[str, str | float] = {
        "path": path,
        "site_id": "S1",
        "drive_id": "D1",
        "item_id": "I1",
        "local_path": "/tmp/wc/foo.txt",
        "etag": "etag-v1",
        "since": 1_900_000_000.0,
    }
    base.update(overrides)
    return CheckedOutEntry(**base)  # type: ignore[arg-type]


# ---------------------------------------------------------------------
# list_all + add roundtrip
# ---------------------------------------------------------------------


def test_list_all_empty_when_no_file(tmp_path: Path) -> None:
    registry = CheckoutRegistry(profile="default", base_dir=tmp_path)
    assert registry.list_all() == []


def test_add_then_list(tmp_path: Path) -> None:
    registry = CheckoutRegistry(profile="default", base_dir=tmp_path)
    e = _entry()
    registry.add(e)
    assert registry.list_all() == [e]


def test_add_replaces_same_path(tmp_path: Path) -> None:
    registry = CheckoutRegistry(profile="default", base_dir=tmp_path)
    registry.add(_entry(path="https://example/foo", etag="v1"))
    registry.add(_entry(path="https://example/foo", etag="v2"))
    entries = registry.list_all()
    assert len(entries) == 1
    assert entries[0].etag == "v2"


def test_add_keeps_distinct_paths(tmp_path: Path) -> None:
    registry = CheckoutRegistry(profile="default", base_dir=tmp_path)
    registry.add(_entry(path="https://example/a"))
    registry.add(_entry(path="https://example/b"))
    paths = [e.path for e in registry.list_all()]
    assert sorted(paths) == ["https://example/a", "https://example/b"]


def test_persists_across_instances(tmp_path: Path) -> None:
    CheckoutRegistry(profile="default", base_dir=tmp_path).add(_entry())
    second = CheckoutRegistry(profile="default", base_dir=tmp_path)
    assert second.list_all() == [_entry()]


def test_per_profile_isolation(tmp_path: Path) -> None:
    a = CheckoutRegistry(profile="a", base_dir=tmp_path)
    b = CheckoutRegistry(profile="b", base_dir=tmp_path)
    a.add(_entry(path="A"))
    b.add(_entry(path="B"))
    assert [e.path for e in a.list_all()] == ["A"]
    assert [e.path for e in b.list_all()] == ["B"]


# ---------------------------------------------------------------------
# get + remove
# ---------------------------------------------------------------------


def test_get_returns_entry_for_path(tmp_path: Path) -> None:
    registry = CheckoutRegistry(profile="default", base_dir=tmp_path)
    registry.add(_entry(path="P", etag="v"))
    found = registry.get("P")
    assert found is not None
    assert found.etag == "v"


def test_get_returns_none_for_missing(tmp_path: Path) -> None:
    registry = CheckoutRegistry(profile="default", base_dir=tmp_path)
    assert registry.get("never-stored") is None


def test_remove_returns_removed_entry(tmp_path: Path) -> None:
    registry = CheckoutRegistry(profile="default", base_dir=tmp_path)
    e = _entry(path="https://example/x")
    registry.add(e)
    removed = registry.remove("https://example/x")
    assert removed == e
    assert registry.list_all() == []


def test_remove_returns_none_when_not_present(tmp_path: Path) -> None:
    registry = CheckoutRegistry(profile="default", base_dir=tmp_path)
    assert registry.remove("never-stored") is None


def test_remove_does_not_affect_other_entries(tmp_path: Path) -> None:
    registry = CheckoutRegistry(profile="default", base_dir=tmp_path)
    registry.add(_entry(path="A"))
    registry.add(_entry(path="B"))
    registry.remove("A")
    assert [e.path for e in registry.list_all()] == ["B"]


# ---------------------------------------------------------------------
# Persistence quirks — corrupt file, file mode
# ---------------------------------------------------------------------


def test_file_mode_is_owner_only(tmp_path: Path) -> None:
    registry = CheckoutRegistry(profile="default", base_dir=tmp_path)
    registry.add(_entry())
    f = tmp_path / "default" / "checked_out.json"
    assert (f.stat().st_mode & 0o777) == 0o600


def test_corrupt_file_is_treated_as_empty(tmp_path: Path) -> None:
    """A truncated/garbage registry doesn't crash callers — they see []."""
    profile_dir = tmp_path / "default"
    profile_dir.mkdir(parents=True)
    (profile_dir / "checked_out.json").write_text("{not valid json", encoding="utf-8")

    registry = CheckoutRegistry(profile="default", base_dir=tmp_path)
    assert registry.list_all() == []


def test_writes_are_atomic_no_temp_left_behind(tmp_path: Path) -> None:
    """After a successful add, no .tmp turd remains in the profile dir."""
    registry = CheckoutRegistry(profile="default", base_dir=tmp_path)
    registry.add(_entry())
    leftover = list((tmp_path / "default").glob("*.tmp"))
    assert leftover == []


def test_dataclass_round_trips_through_json(tmp_path: Path) -> None:
    """All CheckedOutEntry fields survive write+read."""
    e = CheckedOutEntry(
        path="P",
        site_id="S",
        drive_id="D",
        item_id="I",
        local_path="/tmp/wc/x",
        etag="abc123",
        since=1_777_777_777.5,
    )
    registry = CheckoutRegistry(profile="default", base_dir=tmp_path)
    registry.add(e)
    [recovered] = registry.list_all()
    assert recovered == e


def test_invalid_path_raises_on_write(tmp_path: Path) -> None:
    """If the directory can't be created, the error propagates cleanly."""
    # Make tmp_path read-only so mkdir of a child fails. macOS / Linux only.
    (tmp_path).chmod(0o500)
    try:
        registry = CheckoutRegistry(profile="default", base_dir=tmp_path)
        with pytest.raises(OSError):
            registry.add(_entry())
    finally:
        (tmp_path).chmod(0o700)
