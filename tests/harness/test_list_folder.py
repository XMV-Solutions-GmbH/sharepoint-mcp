# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_list against the real harness sandbox.

Skipped when no harness credentials are present.
"""

from __future__ import annotations

import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.tools.list_folder import list_folder

HARNESS_PROFILE = "harness"
HARNESS_SITE_URL = "https://xmvsolutions.sharepoint.com/sites/sharepoint-mcp-harness"
HARNESS_DOCUMENTS_URL = f"{HARNESS_SITE_URL}/Shared%20Documents"
EXPECTED_ITEM_KEYS = {"name", "type", "size", "last_modified", "web_url"}


def _skip_if_no_harness() -> None:
    try:
        get_token(HARNESS_PROFILE)
    except AuthRequiredError as exc:
        pytest.skip(
            f"Harness credentials not available: {exc}. "
            "Run `uv run sharepoint-mcp login --profile harness` to populate.",
        )


def test_sp_list_against_real_harness_documents() -> None:
    """Listing the Documents library of the harness site succeeds.

    Site is currently empty — accept any list (empty or not) as long
    as the call succeeds and the response shape is correct.
    """
    _skip_if_no_harness()
    items = list_folder(HARNESS_DOCUMENTS_URL, profile=HARNESS_PROFILE, limit=50)
    assert isinstance(items, list)
    for item in items:
        assert set(item.keys()) == EXPECTED_ITEM_KEYS
        assert item["type"] in {"folder", "file"}


def test_sp_list_validation_does_not_need_harness() -> None:
    """Local validation kicks in before any Graph call."""
    with pytest.raises(ValueError, match="non-empty url"):
        list_folder("", profile=HARNESS_PROFILE)
