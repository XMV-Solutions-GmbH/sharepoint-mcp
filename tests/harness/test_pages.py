# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for SharePoint Pages tools (#45)."""

from __future__ import annotations

import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.tools.pages import page_read, pages_list

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


def test_pages_list_returns_at_least_default_home_page() -> None:
    """Every modern SharePoint site provisions a default Home.aspx page."""
    _skip_if_no_harness()
    out = pages_list(HARNESS_SITE_URL, profile=HARNESS_PROFILE)
    assert isinstance(out, list)
    assert len(out) >= 1, "harness site should have at least the default Home page"
    for entry in out:
        assert "id" in entry
        assert "name" in entry
        assert "title" in entry
        assert "web_url" in entry


def test_page_read_round_trips_an_existing_page() -> None:
    """Pick the first page from list and read it; canvas may be empty
    on a brand-new site, so we just assert structural fields."""
    _skip_if_no_harness()
    pages = pages_list(HARNESS_SITE_URL, profile=HARNESS_PROFILE)
    if not pages:
        pytest.skip("no pages on harness site to read")
    first = pages[0]
    web_url = first.get("web_url") or ""
    if "/SitePages/" not in web_url:
        pytest.skip(f"first page web_url isn't a /SitePages/ URL: {web_url!r}")
    out = page_read(web_url, profile=HARNESS_PROFILE)
    assert out["id"] == first["id"]
    assert out["name"] == first["name"]
    # canvas_layout key always present in read responses
    assert "canvas_layout" in out


def test_pages_list_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        pages_list("", profile=HARNESS_PROFILE)
    with pytest.raises(ValueError, match="expects a site URL"):
        pages_list(f"{HARNESS_SITE_URL}/Shared Documents/x.docx", profile=HARNESS_PROFILE)


def test_page_read_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="non-empty"):
        page_read("", profile=HARNESS_PROFILE)
    with pytest.raises(ValueError, match="SitePages"):
        page_read(f"{HARNESS_SITE_URL}/Shared Documents/x.docx", profile=HARNESS_PROFILE)
