# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_open against the harness sandbox.

Each test registers cleanup that releases any checkout it acquired,
even on failure, so the sandbox doesn't accumulate stale locks
across test runs.
"""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.checkout_registry import CheckoutRegistry
from sharepoint_mcp.tools.open_file import open_file
from tests.harness._cleanup import HARNESS_PROFILE, discard_checkouts_added_during

HARNESS_SITE_URL = "https://xmvsolutions.sharepoint.com/sites/sharepoint-mcp-harness"
HARNESS_README_URL = f"{HARNESS_SITE_URL}/Shared Documents/README.md"


def _skip_if_no_harness() -> None:
    try:
        get_token(HARNESS_PROFILE)
    except AuthRequiredError as exc:
        pytest.skip(
            f"Harness credentials not available: {exc}. "
            "Run `uv run mcp-server-sharepoint login --profile harness` to populate.",
        )


@pytest.fixture
def cleanup_checkouts() -> Iterator[None]:
    pre = {e.path for e in CheckoutRegistry(HARNESS_PROFILE).list_all()}
    yield from discard_checkouts_added_during(pre)


def test_sp_open_acquires_lock_and_downloads(cleanup_checkouts: None) -> None:
    """sp_open against the seed README: lock acquired, content downloaded,
    registry entry created with all the IDs and ETag we'll need for sp_save.
    """
    del cleanup_checkouts
    _skip_if_no_harness()

    local_path = open_file(HARNESS_README_URL, profile=HARNESS_PROFILE)

    # Working copy exists with the right content
    assert Path(local_path).exists()
    assert "sharepoint-mcp harness sandbox" in Path(local_path).read_text(encoding="utf-8")
    assert local_path.endswith("README.md")

    # Registry has the entry with all IDs populated
    entry = CheckoutRegistry(HARNESS_PROFILE).get(HARNESS_README_URL)
    assert entry is not None
    assert entry.site_id, "site_id should be populated"
    assert entry.drive_id, "drive_id should be populated"
    assert entry.item_id, "item_id should be populated"
    assert entry.etag, "ETag should be populated for stale-write detection in sp_save"
    assert Path(entry.local_path) == Path(local_path)


def test_sp_open_validation_does_not_need_harness() -> None:
    """Validation kicks in before any Graph or registry call."""
    with pytest.raises(ValueError, match="non-empty url"):
        open_file("", profile=HARNESS_PROFILE)
