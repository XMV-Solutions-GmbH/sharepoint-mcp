# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_drive_file_read against the harness sandbox.

Skipped when no harness credentials are present.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.tools.read import read_file

HARNESS_PROFILE = "harness"
HARNESS_SITE_URL = "https://xmvsolutions.sharepoint.com/sites/sharepoint-mcp-harness"
HARNESS_README_URL = f"{HARNESS_SITE_URL}/Shared Documents/README.md"
HARNESS_POLICY_URL = f"{HARNESS_SITE_URL}/Shared Documents/policies/iso27001-control-A.5.1.md"


def _skip_if_no_harness() -> None:
    try:
        get_token(HARNESS_PROFILE)
    except AuthRequiredError as exc:
        pytest.skip(
            f"Harness credentials not available: {exc}. "
            "Run `uv run mcp-server-sharepoint login --profile harness` to populate.",
        )


def test_sp_read_file_downloads_seed_readme() -> None:
    """The README.md seed file uploaded during harness setup is readable."""
    _skip_if_no_harness()
    path = read_file(HARNESS_README_URL, profile=HARNESS_PROFILE)
    try:
        content = Path(path).read_text(encoding="utf-8")
        assert "sharepoint-mcp harness sandbox" in content
        assert path.endswith(".md")
    finally:
        Path(path).unlink(missing_ok=True)


def test_sp_read_file_downloads_subfolder_file() -> None:
    """File in a sub-folder works (drive-root:/folder/file:/content path shape)."""
    _skip_if_no_harness()
    path = read_file(HARNESS_POLICY_URL, profile=HARNESS_PROFILE)
    try:
        content = Path(path).read_text(encoding="utf-8")
        assert "ISO 27001" in content
    finally:
        Path(path).unlink(missing_ok=True)


def test_sp_read_file_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        read_file("", profile=HARNESS_PROFILE)
