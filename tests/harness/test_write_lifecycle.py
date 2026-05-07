# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""End-to-end harness tests for the open/save/release write lifecycle.

Each test runs against the real harness sandbox and follows the
discipline: every checkout it acquires gets released by the test
itself, with the cleanup fixture as a defensive backstop.

Save tests intentionally append harmless content (a marker line) to
the README.md so each run produces a new minor version in SharePoint.
The site fills with version history over time — that's fine, in fact
useful for observing the audit-log integration.
"""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.checkout_registry import CheckoutRegistry
from sharepoint_mcp.tools.open_file import open_file
from sharepoint_mcp.tools.release import release
from sharepoint_mcp.tools.save import save
from tests.harness._cleanup import HARNESS_PROFILE, discard_checkouts_added_during

HARNESS_SITE_URL = "https://xmvsolutions.sharepoint.com/sites/sharepoint-mcp-harness"
HARNESS_README_URL = f"{HARNESS_SITE_URL}/Shared Documents/README.md"


def _skip_if_no_harness() -> None:
    try:
        get_token(HARNESS_PROFILE)
    except AuthRequiredError as exc:
        pytest.skip(
            f"Harness credentials not available: {exc}. "
            "Run `uv run sharepoint-mcp login --profile harness` to populate.",
        )


@pytest.fixture
def cleanup_checkouts() -> Iterator[None]:
    pre = {e.path for e in CheckoutRegistry(HARNESS_PROFILE).list_all()}
    yield from discard_checkouts_added_during(pre)


def test_open_then_release_leaves_no_lasting_state(cleanup_checkouts: None) -> None:
    """Lifecycle 1: open → release. No version created, no lasting changes."""
    del cleanup_checkouts
    _skip_if_no_harness()

    local_path = open_file(HARNESS_README_URL, profile=HARNESS_PROFILE)
    assert Path(local_path).exists()
    assert CheckoutRegistry(HARNESS_PROFILE).get(HARNESS_README_URL) is not None

    release(HARNESS_README_URL, profile=HARNESS_PROFILE)

    # After release: registry empty, local file gone
    assert CheckoutRegistry(HARNESS_PROFILE).get(HARNESS_README_URL) is None
    assert not Path(local_path).exists()


def test_open_modify_save_creates_new_version(cleanup_checkouts: None) -> None:
    """Lifecycle 2: open → modify locally → save → registry cleaned up.

    Appends a harmless marker line to README.md so we can verify a
    new version was actually committed to SharePoint.
    """
    del cleanup_checkouts
    _skip_if_no_harness()

    local_path = open_file(HARNESS_README_URL, profile=HARNESS_PROFILE)
    original = Path(local_path).read_text(encoding="utf-8")
    marker = "\n<!-- harness-test marker: open-modify-save lifecycle -->\n"
    Path(local_path).write_text(original + marker, encoding="utf-8")

    result = save(
        HARNESS_README_URL,
        comment="harness lifecycle test: append marker line",
        version="minor",
        profile=HARNESS_PROFILE,
    )

    # Save returns metadata with non-empty version_id and etag
    assert result["version_id"], f"expected version_id, got {result!r}"
    assert result["etag"], f"expected etag, got {result!r}"
    assert result["web_url"]

    # Local cleanup — registry cleared, working file removed
    assert CheckoutRegistry(HARNESS_PROFILE).get(HARNESS_README_URL) is None
    assert not Path(local_path).exists()


def test_release_idempotent_on_unknown_path(cleanup_checkouts: None) -> None:
    """sp_release silently no-ops when nothing is checked out — even with valid creds."""
    del cleanup_checkouts
    _skip_if_no_harness()
    # Path NOT in registry. Should not raise, should not call Graph.
    release(
        f"{HARNESS_SITE_URL}/Shared Documents/this-was-never-opened.txt",
        profile=HARNESS_PROFILE,
    )
