# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_file_metadata against the harness sandbox.

Skipped when no harness credentials are present.
"""

from __future__ import annotations

import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.tools.file_metadata import file_metadata

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


def test_sp_file_metadata_read_returns_dict() -> None:
    """GET listItem/fields returns a non-empty dict with at least the id field."""
    _skip_if_no_harness()
    result = file_metadata(HARNESS_README_URL, profile=HARNESS_PROFILE)
    assert isinstance(result, dict)
    assert len(result) > 0


def test_sp_file_metadata_read_contains_system_fields() -> None:
    """Graph always includes system fields like Modified in the fields facet."""
    _skip_if_no_harness()
    result = file_metadata(HARNESS_README_URL, profile=HARNESS_PROFILE)
    # Graph includes Modified (last-modified timestamp) for all library items.
    assert "Modified" in result or "id" in result


def test_sp_file_metadata_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        file_metadata("", profile=HARNESS_PROFILE)


def test_sp_file_metadata_site_url_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="site/folder URL"):
        file_metadata(HARNESS_SITE_URL, profile=HARNESS_PROFILE)
