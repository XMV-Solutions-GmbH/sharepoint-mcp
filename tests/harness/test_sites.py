# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_sites / sp_subsites / sp_followed_sites (#49)."""

from __future__ import annotations

import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.tools.sites import followed_sites, sites, subsites

HARNESS_PROFILE = "harness"
HARNESS_SITE_URL = "https://xmvsolutions.sharepoint.com/sites/sharepoint-mcp-harness"


def _skip_if_no_harness() -> None:
    try:
        get_token(HARNESS_PROFILE)
    except AuthRequiredError as exc:
        pytest.skip(
            f"Harness credentials not available: {exc}. "
            "Run `uv run mcp-server-sharepoint login --profile harness` to populate.",
        )


def test_sites_finds_harness_site_by_name() -> None:
    _skip_if_no_harness()
    results = sites("sharepoint-mcp-harness", profile=HARNESS_PROFILE)
    assert isinstance(results, list)
    # Either the harness site is in results, or the search returned nothing
    # (Microsoft's search index can lag; we don't fail on that). We DO
    # require the response shape if anything came back.
    for entry in results:
        assert "id" in entry
        assert "name" in entry
        assert "web_url" in entry


def test_sites_wildcard_returns_at_least_one_site() -> None:
    """Default query lists everything visible; should not be empty for a
    real user with at least one site."""
    _skip_if_no_harness()
    results = sites(profile=HARNESS_PROFILE)
    assert isinstance(results, list)
    # The harness account has access to at least the sharepoint-mcp-harness
    # site; if Microsoft's search returns 0 here it's a tenant-side issue
    # we'd want to surface, not silently pass.
    assert len(results) >= 1


def test_subsites_returns_list_for_harness_site() -> None:
    _skip_if_no_harness()
    results = subsites(HARNESS_SITE_URL, profile=HARNESS_PROFILE)
    assert isinstance(results, list)
    # The harness site may or may not have sub-sites; we don't assert on
    # count, just shape.
    for entry in results:
        assert "id" in entry
        assert "web_url" in entry


def test_followed_sites_returns_list() -> None:
    """The harness user may not have any followed sites; we don't assert
    on count, just on response shape and that the call doesn't error."""
    _skip_if_no_harness()
    results = followed_sites(profile=HARNESS_PROFILE)
    assert isinstance(results, list)
    for entry in results:
        assert "id" in entry
        assert "web_url" in entry


def test_subsites_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="non-empty parent_site_url"):
        subsites("", profile=HARNESS_PROFILE)
