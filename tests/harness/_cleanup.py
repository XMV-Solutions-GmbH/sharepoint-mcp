# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Best-effort cleanup helper for harness tests that acquire checkouts.

Avoids leaving the sandbox in a "checked out by d.koller" state after
a test fails or times out. Used by sp_drive_file_checkout's harness tests until
sp_drive_file_checkout_discard exists, then by lifecycle tests as a defensive guard.
"""

from __future__ import annotations

from collections.abc import Iterator

import httpx

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.checkout_registry import CheckedOutEntry, CheckoutRegistry
from sharepoint_mcp.tools._common import GRAPH_BASE

HARNESS_PROFILE = "harness"


def discard_checkouts_added_during(
    pre_paths: set[str],
) -> Iterator[None]:
    """Generator that yields once, then discards new registry entries.

    Use as the body of a pytest fixture:

        @pytest.fixture
        def cleanup_after():
            pre = {e.path for e in CheckoutRegistry(HARNESS_PROFILE).list_all()}
            yield from discard_checkouts_added_during(pre)
    """
    yield
    registry = CheckoutRegistry(profile=HARNESS_PROFILE)
    new_entries = [e for e in registry.list_all() if e.path not in pre_paths]
    if not new_entries:
        return
    try:
        token = get_token(HARNESS_PROFILE)
    except AuthRequiredError:
        return
    headers = {"Authorization": f"Bearer {token}"}
    with httpx.Client(timeout=30.0) as client:
        for entry in new_entries:
            _force_discard(client, entry, headers)
            registry.remove(entry.path)


def _force_discard(
    client: httpx.Client,
    entry: CheckedOutEntry,
    headers: dict[str, str],
) -> None:
    """Best-effort: call /discardCheckout. Swallow errors."""
    try:
        client.post(
            f"{GRAPH_BASE}/drives/{entry.drive_id}/items/{entry.item_id}/discardCheckout",
            headers=headers,
        )
    except httpx.HTTPError:
        pass
