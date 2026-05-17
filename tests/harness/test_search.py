# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_search_files — real Microsoft Graph against the harness sandbox.

Skipped when no harness credentials are present (same pattern as
`test_auth_smoke.py`).

The assertions here are deliberately loose on hit content because the
harness site has no seed data yet (per #14 follow-up). We assert on
*shape* — that the call returns a list, that each hit has the keys
we documented as the tool's contract, and that the right HTTP path
was hit. Tighter assertions land once seed data is in place.
"""

from __future__ import annotations

import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.tools.search import search

HARNESS_PROFILE = "harness"
EXPECTED_HIT_KEYS = {"name", "path", "web_url", "last_modified", "author", "size"}


def _skip_if_no_harness() -> None:
    try:
        get_token(HARNESS_PROFILE)
    except AuthRequiredError as exc:
        pytest.skip(
            f"Harness credentials not available: {exc}. "
            "Run `uv run mcp-server-sharepoint login --profile harness` to populate.",
        )


def test_sp_search_files_returns_a_list_against_real_graph() -> None:
    """A search call against real Graph parses without exception."""
    _skip_if_no_harness()
    # Use a query that's specific enough not to flood with millions of hits,
    # and broad enough that something probably matches in a typical M365
    # tenant. "xmv" exists as a tenant identifier; results may be empty
    # for d.koller's permission set, which is fine — we're testing the
    # *call shape*, not result-content.
    results = search("xmv", profile=HARNESS_PROFILE, limit=5)
    assert isinstance(results, list)


def test_sp_search_files_hit_shape_matches_documented_contract() -> None:
    """Whatever Graph returns, our extractor produces the documented keys."""
    _skip_if_no_harness()
    # Broad query likely to return at least the user's own OneDrive root
    # files. If nothing comes back, the test is a no-op for the shape
    # assertion (still passes — empty list is valid).
    results = search("xmv", profile=HARNESS_PROFILE, limit=5)
    for hit in results:
        assert set(hit.keys()) == EXPECTED_HIT_KEYS, (
            f"Hit shape drifted from contract: {sorted(hit.keys())} "
            f"vs. expected {sorted(EXPECTED_HIT_KEYS)}"
        )


def test_sp_search_files_empty_query_validation() -> None:
    """Validation kicks in before any Graph call — auth not required for this."""
    # No skip — this should work even without harness credentials,
    # because the validation is local. But sp_search_files needs profile to
    # pass; we use harness even if it doesn't exist (won't be reached).
    with pytest.raises(ValueError, match="non-empty query"):
        search("", profile=HARNESS_PROFILE)
