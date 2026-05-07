# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Cross-module integration: open → save and open → release lifecycles.

These tests don't go through the MCP-server / FastMCP layer; they
exercise the underlying Python functions across module boundaries
(`tools/open_file` + `checkout_registry` + `tools/save` +
`tools/release` + `auth/get_token`) with HTTP mocked at the boundary.
The point: catch wiring bugs that single-tool unit tests miss because
they only test one module at a time.
"""

from __future__ import annotations

from pathlib import Path

import pytest
import respx

from sharepoint_mcp.checkout_registry import CheckoutRegistry
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.open_file import open_file
from sharepoint_mcp.tools.release import release
from sharepoint_mcp.tools.save import save
from sharepoint_mcp.tools.status import status

URL = "https://contoso.sharepoint.com/sites/foo/Shared Documents/policy.docx"
SITE_HOST = "contoso.sharepoint.com"
SITE_PATH = "/sites/foo"
SITE_ID = "contoso.sharepoint.com,site-guid,web-guid"
DRIVE_ID = "b!drive"
ITEM_ID = "01ITEM"
ETAG_OPEN = '"open-etag,1"'
ETAG_AFTER_SAVE = '"saved-etag,2"'


def _mock_open_calls(content: bytes = b"original") -> None:
    respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(json={"id": SITE_ID})
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/policy.docx").respond(
        json={
            "id": ITEM_ID,
            "name": "policy.docx",
            "eTag": ETAG_OPEN,
            "parentReference": {"driveId": DRIVE_ID},
        },
    )
    respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/checkout").respond(204)
    respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/content").respond(content=content)


def _mock_save_calls(version_id: str = "2.0") -> None:
    respx.put(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/content").respond(
        json={"eTag": ETAG_AFTER_SAVE, "webUrl": URL},
    )
    respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/checkin").respond(204)
    respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/versions").respond(
        json={"value": [{"id": version_id}]},
    )


def _mock_discard() -> None:
    respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/discardCheckout").respond(204)


# ---------------------------------------------------------------------
# open → save lifecycle
# ---------------------------------------------------------------------


@respx.mock
def test_open_modify_save_full_chain(fake_token_store: None, isolated_registry_dir: Path) -> None:
    """sp_open populates the registry; sp_save reads the entry, uses its
    ETag for If-Match, succeeds, and removes the registry entry."""
    del fake_token_store
    _mock_open_calls(content=b"original-content")
    _mock_save_calls(version_id="2.0")

    # Open phase
    local_path = open_file(URL)
    assert Path(local_path).read_bytes() == b"original-content"

    registry_after_open = CheckoutRegistry(profile="default")
    entry = registry_after_open.get(URL)
    assert entry is not None
    assert entry.etag == ETAG_OPEN

    # Modify locally
    Path(local_path).write_bytes(b"updated-content")

    # Save phase
    result = save(URL, comment="updated", version="minor")
    assert result["version_id"] == "2.0"
    assert result["etag"] == ETAG_AFTER_SAVE

    # Registry cleared
    assert CheckoutRegistry(profile="default").get(URL) is None
    # Working file removed
    assert not Path(local_path).exists()


@respx.mock
def test_open_release_lifecycle(fake_token_store: None, isolated_registry_dir: Path) -> None:
    """sp_open populates registry; sp_release calls discardCheckout
    and removes the entry. No save call required."""
    del fake_token_store
    _mock_open_calls()
    _mock_discard()

    local_path = open_file(URL)
    assert CheckoutRegistry(profile="default").get(URL) is not None

    release(URL)

    assert CheckoutRegistry(profile="default").get(URL) is None
    assert not Path(local_path).exists()


# ---------------------------------------------------------------------
# sp_status sees what sp_open writes
# ---------------------------------------------------------------------


@respx.mock
def test_status_reflects_open(fake_token_store: None, isolated_registry_dir: Path) -> None:
    """Cross-module: sp_open's registry write is visible to sp_status."""
    del fake_token_store
    assert status() == []  # nothing checked out initially

    _mock_open_calls()
    open_file(URL)

    visible = status()
    assert len(visible) == 1
    entry = visible[0]
    assert entry["path"] == URL
    assert entry["local_path"]
    assert entry["since"]


# ---------------------------------------------------------------------
# Stale-write across modules: open A, foreign-modify, save A → 412
# ---------------------------------------------------------------------


@respx.mock
def test_save_412_propagates_through_lifecycle(
    fake_token_store: None, isolated_registry_dir: Path
) -> None:
    """Simulate: agent opens, someone else changes the file (server's
    eTag changes), agent's save fails 412 → StaleWriteError. The
    registry entry must NOT be cleared (caller still owns the lock and
    must release/re-open to recover)."""
    from sharepoint_mcp.tools.save import StaleWriteError

    del fake_token_store
    _mock_open_calls()

    # Open succeeds
    open_file(URL)
    assert CheckoutRegistry(profile="default").get(URL) is not None

    # Save: PUT comes back with 412
    respx.put(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/content").respond(
        412, json={"error": {"code": "preconditionFailed"}}
    )

    with pytest.raises(StaleWriteError, match="changed under us"):
        save(URL, comment="updated")

    # Registry MUST still hold the entry (ownership preserved)
    assert CheckoutRegistry(profile="default").get(URL) is not None
