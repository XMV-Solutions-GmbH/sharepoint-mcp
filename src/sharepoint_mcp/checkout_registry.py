# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Persistent registry of files currently checked out by this MCP profile.

Tracks the local working copies created by `sp_drive_file_checkout` so `sp_drive_file_checkin` can
look up the corresponding ETag (for stale-write detection) and Graph
ids without re-resolving from URL each time, and so `sp_drive_checkout_list` can
list them for the agent + the human.

Layout: `<base_dir>/<profile>/checked_out.json` with mode 0o600.
Atomic write via temp-file + rename so a crash mid-write doesn't
leave a half-truncated registry.

Server-side reconciliation (verifying that each entry is *still*
checked out on the SharePoint side) is deferred to v0.2 — for v0.1
we trust our local view, and `sp_drive_file_checkin` catches divergence via the
ETag round-trip.
"""

from __future__ import annotations

import json
import os
import tempfile
import threading
import time
from dataclasses import asdict, dataclass
from pathlib import Path

DEFAULT_REGISTRY_DIR = Path.home() / ".cache" / "sharepoint-mcp"

# Process-wide lock for registry mutations. Bulk operations
# (sp_drive_file_checkout_bulk / sp_drive_file_checkin_bulk) run multiple `add` / `remove`
# concurrently, and `add` does a non-atomic read-modify-write
# (list_all → mutate → _write). Without this lock, two concurrent
# adds can stomp each other and lose entries. The lock is held only
# for the duration of each mutation; underlying I/O is fast.
_REGISTRY_LOCK = threading.Lock()


@dataclass(frozen=True)
class CheckedOutEntry:
    """One row in the checked-out registry."""

    path: str  # original SharePoint URL passed to sp_drive_file_checkout
    site_id: str
    drive_id: str
    item_id: str
    local_path: str  # working-copy file on disk
    etag: str  # If-Match header for sp_drive_file_checkin's stale-write check
    since: float  # epoch seconds when sp_drive_file_checkout succeeded


class CheckoutRegistry:
    """File-backed registry of checked-out items, scoped to one profile."""

    def __init__(self, profile: str, base_dir: Path | None = None) -> None:
        self._dir = (base_dir if base_dir is not None else DEFAULT_REGISTRY_DIR) / profile
        self._registry_file = self._dir / "checked_out.json"

    def list_all(self) -> list[CheckedOutEntry]:
        """Return every currently-tracked entry. Empty list if none."""
        if not self._registry_file.exists():
            return []
        try:
            raw = json.loads(self._registry_file.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            # Corrupt file — treat as empty rather than crash. Caller can
            # delete the registry to recover.
            return []
        return [CheckedOutEntry(**row) for row in raw]

    def get(self, path: str) -> CheckedOutEntry | None:
        for entry in self.list_all():
            if entry.path == path:
                return entry
        return None

    def add(self, entry: CheckedOutEntry) -> None:
        """Add or replace the entry for `entry.path`. Thread-safe."""
        with _REGISTRY_LOCK:
            existing = [e for e in self.list_all() if e.path != entry.path]
            existing.append(entry)
            self._write(existing)

    def remove(self, path: str) -> CheckedOutEntry | None:
        """Remove the entry for `path`. Returns the removed entry, or None.
        Thread-safe."""
        with _REGISTRY_LOCK:
            existing = self.list_all()
            match = next((e for e in existing if e.path == path), None)
            if match is None:
                return None
            remaining = [e for e in existing if e.path != path]
            self._write(remaining)
        return match

    def _write(self, entries: list[CheckedOutEntry]) -> None:
        self._dir.mkdir(parents=True, exist_ok=True)
        # Atomic write: temp file in same directory, then rename.
        fd, tmp_path = tempfile.mkstemp(
            prefix="checked_out-",
            suffix=".json.tmp",
            dir=self._dir,
        )
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump([asdict(e) for e in entries], f, indent=2)
            os.chmod(tmp_path, 0o600)
            os.replace(tmp_path, self._registry_file)
        except OSError:
            try:
                Path(tmp_path).unlink()
            except FileNotFoundError:
                pass
            raise


def now() -> float:
    """Wall-clock epoch seconds; injectable for tests."""
    return time.time()
