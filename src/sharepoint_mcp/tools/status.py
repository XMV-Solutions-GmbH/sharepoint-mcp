# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_status — list the files currently checked out by this profile.

Read-only. Two modes:

- `verify=False` (default): returns the local view from the persistent
  CheckoutRegistry only. Sub-second, no Graph calls. Sufficient for
  most use cases — sp_save_file's ETag round-trip catches divergence at
  the moment that actually matters.
- `verify=True`: for each registry entry, queries Microsoft Graph
  for the SharePoint listItem's `CheckoutUser` field to confirm the
  server-side lock state. Surfaces divergence the agent can act on
  (e.g., admin manually discarded our lock; lock held by someone
  else; item was deleted server-side).

The `verify=True` cost is one Graph call per registry entry, and
each entry adds the same call's worth of latency. Per
`docs/spikes/2026-05-07-working-dir-cleanup.md`, this is the v0.2
shape of the deferred reconciliation; v0.1 shipped without it.
"""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Any

import httpx

from sharepoint_mcp.auth import get_token
from sharepoint_mcp.checkout_registry import CheckoutRegistry
from sharepoint_mcp.tools._common import GRAPH_BASE


def status(
    profile: str = "default",
    *,
    verify: bool = False,
    http: httpx.Client | None = None,
) -> list[dict[str, Any]]:
    """Return one dict per currently-checked-out file under this profile.

    With `verify=False` (default), each entry has `path` (original
    SharePoint URL), `since` (ISO datetime UTC of when sp_open_file
    succeeded), and `local_path`. Empty list when nothing is open.

    With `verify=True`, each entry additionally has:

    - `server_locked` — `True` if SharePoint reports a lock,
      `False` if no lock, `None` if we couldn't determine (network
      error, item deleted, etc.).
    - `lock_holder` — display-name of the user holding the lock if
      known, `None` otherwise. Compare to the agent's identity to
      tell "our lock" from "someone else's lock".
    """
    registry = CheckoutRegistry(profile=profile)
    entries = registry.list_all()

    base_results: list[dict[str, Any]] = [
        {
            "path": entry.path,
            "since": datetime.fromtimestamp(entry.since, tz=UTC).isoformat(),
            "local_path": entry.local_path,
        }
        for entry in entries
    ]

    if not verify:
        return base_results
    if not base_results:
        return base_results

    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0)
    try:
        for result, entry in zip(base_results, entries, strict=True):
            server_locked, lock_holder = _query_lock_state(
                client,
                drive_id=entry.drive_id,
                item_id=entry.item_id,
                headers=headers,
            )
            result["server_locked"] = server_locked
            result["lock_holder"] = lock_holder
    finally:
        if http is None:
            client.close()
    return base_results


def _query_lock_state(
    client: httpx.Client,
    *,
    drive_id: str,
    item_id: str,
    headers: dict[str, str],
) -> tuple[bool | None, str | None]:
    """Query SharePoint's listItem fields for the CheckoutUser.

    Returns `(server_locked, lock_holder_display_name)`. On error
    (404, 403, network blip), returns `(None, None)` — caller can
    treat that as "couldn't determine" and decide what to do.
    """
    try:
        response = client.get(
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/listItem",
            headers=headers,
            params={"$expand": "fields($select=CheckoutUser)"},
        )
    except httpx.HTTPError:
        return None, None

    if response.status_code != 200:
        return None, None

    fields = response.json().get("fields") or {}
    checkout_user = fields.get("CheckoutUser")
    if not checkout_user:
        return False, None

    # CheckoutUser is typically a string (display-name) on SharePoint
    # lists, but Graph sometimes returns it as a structured user
    # object. Handle both shapes.
    if isinstance(checkout_user, str):
        return True, checkout_user
    if isinstance(checkout_user, dict):
        name = checkout_user.get("displayName") or checkout_user.get("LookupValue")
        return True, (str(name) if name else None)
    if isinstance(checkout_user, list) and checkout_user:
        first = checkout_user[0]
        if isinstance(first, dict):
            name = first.get("LookupValue") or first.get("displayName")
            return True, str(name) if name else None
    return True, None
