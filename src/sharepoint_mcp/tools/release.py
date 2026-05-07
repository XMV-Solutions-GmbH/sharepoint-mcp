# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_release — discard a pending checkout without saving.

`POST /drives/{id}/items/{id}/discardCheckout` releases the
server-side lock without committing any local changes. The local
working copy is also deleted, and the registry entry removed.

Idempotent in spirit: if the local registry doesn't have an entry
for `url`, that's already the desired post-state, so we return
silently rather than raising. Best-effort cleanup of the local
working file in any case.

The Graph call itself can fail (network blip, server-side state
diverged) — those errors propagate up so the caller sees them. The
local cleanup happens regardless; a leftover server-side lock will
be visible via SharePoint's web UI and can be discarded by the user
or an admin.
"""

from __future__ import annotations

from pathlib import Path

import httpx

from sharepoint_mcp.auth import get_token
from sharepoint_mcp.checkout_registry import CheckoutRegistry
from sharepoint_mcp.tools._common import GRAPH_BASE


def release(
    url: str,
    *,
    profile: str = "default",
    http: httpx.Client | None = None,
) -> None:
    """Discard the checkout for `url` without saving changes.

    No-op (returns silently) when nothing is registered locally for
    `url` — either we never sp_open'd it, or sp_release was already
    called. Either way, the desired state is "no local checkout for
    this path", which is true.

    Raises:
        ValueError: empty url.
        httpx.HTTPStatusError: server refused discardCheckout. Local
            registry entry + working file still get cleaned up.
    """
    if not url or not url.strip():
        raise ValueError("sp_release requires a non-empty url")

    registry = CheckoutRegistry(profile=profile)
    entry = registry.get(url)
    if entry is None:
        return

    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=30.0, follow_redirects=True)
    try:
        try:
            response = client.post(
                f"{GRAPH_BASE}/drives/{entry.drive_id}/items/{entry.item_id}/discardCheckout",
                headers=headers,
            )
            response.raise_for_status()
        finally:
            # Local cleanup happens regardless of server-side outcome.
            # The server-side lock can be discarded by the user later
            # if our discardCheckout failed; we don't want to leave
            # the local registry pretending the file is checked out.
            registry.remove(url)
            local_file = Path(entry.local_path)
            try:
                local_file.unlink()
            except FileNotFoundError:
                pass
            try:
                local_file.parent.rmdir()
            except OSError:
                pass
    finally:
        if http is None:
            client.close()
