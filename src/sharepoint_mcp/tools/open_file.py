# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_open — checkout a SharePoint file and download its content.

Acquires a server-side lock via `POST /drives/{id}/items/{id}/checkout`,
downloads the current content into a working-directory path, and
registers the entry in `CheckoutRegistry` so `sp_save` can find the
ETag for stale-write detection and `sp_release` knows what to discard.

Working-copy layout: `<base_dir>/<profile>/working/<item-id>/<filename>`.
The per-item-id subfolder prevents name collisions between checkouts
of different files that happen to share a basename, and makes
finding the working copy by content trivial.

Module is named `open_file` instead of `open` to avoid shadowing the
Python builtin.
"""

from __future__ import annotations

import time
from collections.abc import Callable
from pathlib import Path

import httpx

from sharepoint_mcp import checkout_registry as _registry_module
from sharepoint_mcp.auth import get_token
from sharepoint_mcp.checkout_registry import CheckedOutEntry, CheckoutRegistry
from sharepoint_mcp.tools._common import (
    GRAPH_BASE,
    parse_sharepoint_url,
    resolve_drive_item_full,
    resolve_site_id,
)

WORKING_DIR_NAME = "working"


class CheckoutConflictError(RuntimeError):
    """The file is already checked out by another user.

    Raised by `open_file` when SharePoint refuses the checkout because
    a different identity holds the lock. The agent / caller should
    surface the message to the human.
    """


def open_file(
    url: str,
    *,
    profile: str = "default",
    base_dir: Path | None = None,
    http: httpx.Client | None = None,
    now: Callable[[], float] = time.time,
) -> str:
    """Checkout the SharePoint file at `url` and return the local working path.

    Side effects (visible to other SharePoint users):

    1. Server-side checkout lock acquired (others see "checked out by
       <test-user>" until `sp_release` or `sp_save` is called).

    Side effects (local):

    2. File content written to
       `<base_dir>/<profile>/working/<item-id>/<filename>`.
    3. `CheckoutRegistry` entry created so subsequent `sp_save` knows
       which item-id, drive-id, and ETag to use.

    Raises:
        ValueError: empty URL or URL points at a site/folder, not a file.
        CheckoutConflictError: file is locked by another user.
        httpx.HTTPStatusError: any other non-2xx from Graph (404 if the
            item doesn't exist, 403 if the user can't see it).
        sharepoint_mcp.auth.AuthRequiredError: no cached token for `profile`.
    """
    if not url or not url.strip():
        raise ValueError("sp_open requires a non-empty url")

    hostname, site_path, item_path = parse_sharepoint_url(url)
    if not item_path:
        raise ValueError(f"sp_open needs a file URL, got {url!r}")

    resolved_base = base_dir if base_dir is not None else _registry_module.DEFAULT_REGISTRY_DIR
    token = get_token(profile)
    headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=60.0, follow_redirects=True)
    try:
        site_id = resolve_site_id(client, hostname, site_path, headers=headers)

        # Look up the driveItem to get id, drive_id, ETag, and the
        # canonical filename. Done before /checkout so we can fail
        # cleanly on 404 without leaving a dangling lock. The full
        # resolver handles non-default libraries transparently.
        item = resolve_drive_item_full(client, site_id, item_path, headers=headers)
        drive_id = item["parentReference"]["driveId"]
        item_id = item["id"]
        etag = str(item.get("eTag") or item.get("@odata.etag") or "")
        filename = item.get("name") or item_path.rsplit("/", 1)[-1]

        # Acquire checkout lock. 204 No Content on success;
        # 423 Locked if someone else has it.
        checkout_response = client.post(
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/checkout",
            headers=headers,
        )
        if checkout_response.status_code == 423:
            raise CheckoutConflictError(
                f"Cannot checkout {url!r}: file is already checked out by another user.",
            )
        checkout_response.raise_for_status()

        # Download current content (post-checkout reads the locked version).
        content_response = client.get(
            f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content",
            headers=headers,
        )
        content_response.raise_for_status()

        work_dir = resolved_base / profile / WORKING_DIR_NAME / item_id
        work_dir.mkdir(parents=True, exist_ok=True)
        local_path = work_dir / filename
        local_path.write_bytes(content_response.content)

        registry = CheckoutRegistry(profile=profile, base_dir=resolved_base)
        registry.add(
            CheckedOutEntry(
                path=url,
                site_id=site_id,
                drive_id=drive_id,
                item_id=item_id,
                local_path=str(local_path),
                etag=etag,
                since=now(),
            ),
        )
        return str(local_path)
    finally:
        if http is None:
            client.close()
