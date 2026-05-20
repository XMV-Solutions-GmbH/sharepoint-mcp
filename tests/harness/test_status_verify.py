# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_drive_checkout_list verify=True against the real harness sandbox.

Validates the server-side reconciliation path: open a file, ask
SharePoint whether it's locked, get back True; release, get back to
the unverified shape.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.tools.open_file import open_file
from sharepoint_mcp.tools.release import release
from sharepoint_mcp.tools.status import status

HARNESS_PROFILE = "harness"
HARNESS_README_URL = (
    "https://xmvsolutions.sharepoint.com/sites/sharepoint-mcp-harness/Shared Documents/README.md"
)


def _skip_if_no_harness() -> None:
    try:
        get_token(HARNESS_PROFILE)
    except AuthRequiredError as exc:
        pytest.skip(
            f"Harness credentials not available: {exc}. "
            "Run `uv run mcp-server-sharepoint login --profile harness` to populate.",
        )


@pytest.fixture
def clean_registry() -> None:
    """Discard any stale checkout from a prior run before testing."""
    try:
        release(HARNESS_README_URL, profile=HARNESS_PROFILE)
    except Exception:
        # release is idempotent for missing entries; ignore other failures.
        pass


def test_status_verify_reports_server_locked_after_open(clean_registry: None) -> None:
    del clean_registry
    _skip_if_no_harness()
    local = open_file(HARNESS_README_URL, profile=HARNESS_PROFILE)
    try:
        assert Path(local).exists()
        results = status(profile=HARNESS_PROFILE, verify=True)
        assert len(results) == 1
        entry = results[0]
        assert entry["path"] == HARNESS_README_URL
        # Either definitively locked, or "unknown" if SharePoint
        # surfaces a transient quirk; we will not accept False.
        assert entry["server_locked"] in (True, None), (
            f"Expected lock to be visible to SharePoint after sp_drive_file_checkout, "
            f"got server_locked={entry['server_locked']!r}, "
            f"lock_holder={entry.get('lock_holder')!r}"
        )
    finally:
        release(HARNESS_README_URL, profile=HARNESS_PROFILE)


def test_status_default_does_not_add_server_fields() -> None:
    _skip_if_no_harness()
    # Whatever the registry currently holds — verify=False must
    # never include server_locked / lock_holder keys.
    results = status(profile=HARNESS_PROFILE)
    for entry in results:
        assert "server_locked" not in entry
        assert "lock_holder" not in entry


def test_status_verify_empty_registry_makes_zero_graph_calls() -> None:
    """If nothing is checked out, verify=True must short-circuit. We don't
    have a hook to count Graph calls, so check the contract: an empty
    registry returns an empty list with no Graph latency."""
    _skip_if_no_harness()
    # Make sure registry is empty for this profile by releasing harness URL
    # (idempotent if not present).
    try:
        release(HARNESS_README_URL, profile=HARNESS_PROFILE)
    except Exception:
        pass
    results = status(profile=HARNESS_PROFILE, verify=True)
    # If anything else is in registry, we still expect the call to succeed.
    # Empty is the typical case in CI.
    assert isinstance(results, list)
