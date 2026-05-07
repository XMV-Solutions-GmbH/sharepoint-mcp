# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_permissions (#46)."""

from __future__ import annotations

import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.tools.permissions import permissions

HARNESS_PROFILE = "harness"
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


def test_permissions_on_site_returns_at_least_one_grant() -> None:
    """Every SharePoint site has at least the site owner / member grants."""
    _skip_if_no_harness()
    out = permissions(HARNESS_SITE_URL, profile=HARNESS_PROFILE)
    assert isinstance(out, list)
    assert len(out) >= 1
    for entry in out:
        assert "id" in entry
        assert "roles" in entry
        assert "grantee" in entry
        assert "inherited" in entry


def test_permissions_on_known_file_returns_list() -> None:
    """The README file may inherit permissions; we don't assert on count,
    just on response shape."""
    _skip_if_no_harness()
    out = permissions(HARNESS_README_URL, profile=HARNESS_PROFILE)
    assert isinstance(out, list)
    for entry in out:
        grantee = entry.get("grantee", {})
        assert grantee.get("type") in {
            "user",
            "group",
            "link",
            "siteUser",
            "siteGroup",
            "application",
            "unknown",
        }


def test_permissions_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        permissions("", profile=HARNESS_PROFILE)
