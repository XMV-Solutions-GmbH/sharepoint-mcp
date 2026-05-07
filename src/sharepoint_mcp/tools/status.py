# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_status — list the files currently checked out by this profile.

Read-only. Returns the local view from the persistent CheckoutRegistry;
does NOT verify with SharePoint that the lock is still held server-side
(deferred to v0.2). That trust gap is intentional for v0.1: sp_save's
ETag round-trip catches divergence at write time, and that's the
moment that actually matters.
"""

from __future__ import annotations

from datetime import UTC
from typing import Any

from sharepoint_mcp.checkout_registry import CheckoutRegistry


def status(profile: str = "default") -> list[dict[str, Any]]:
    """Return one dict per currently-checked-out file under this profile.

    Each entry has `path` (original SharePoint URL), `since` (ISO
    datetime UTC of when sp_open succeeded), and `local_path` (the
    working copy on disk).

    Empty list when nothing is currently checked out.
    """
    from datetime import datetime

    registry = CheckoutRegistry(profile=profile)
    return [
        {
            "path": entry.path,
            "since": datetime.fromtimestamp(entry.since, tz=UTC).isoformat(),
            "local_path": entry.local_path,
        }
        for entry in registry.list_all()
    ]
