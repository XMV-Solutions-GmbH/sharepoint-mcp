# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_file_history + sp_get_file_version against the real harness sandbox.

The seed README.md has accumulated versions through prior harness
lifecycle tests (each `test_open_modify_save_creates_new_version` run
appends a marker line and saves a new minor version). So we expect at
least one historical version to exist when these tests run.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.tools.get_version import get_version
from sharepoint_mcp.tools.history import history

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


def test_history_returns_at_least_one_version() -> None:
    _skip_if_no_harness()
    versions = history(HARNESS_README_URL, profile=HARNESS_PROFILE, limit=10)
    assert isinstance(versions, list)
    assert len(versions) >= 1
    for v in versions:
        assert v["id"]
        assert v["last_modified"]
        assert v["size"] > 0


def test_history_versions_ordered_newest_first() -> None:
    from itertools import pairwise

    _skip_if_no_harness()
    versions = history(HARNESS_README_URL, profile=HARNESS_PROFILE, limit=10)
    if len(versions) < 2:
        pytest.skip("Need at least two versions to test ordering")
    # Newest first: each pair (current, next) must have current >= next
    for current, nxt in pairwise(versions):
        assert current["last_modified"] >= nxt["last_modified"], (
            f"versions not ordered newest-first: "
            f"{current['last_modified']} < {nxt['last_modified']}"
        )


def test_get_version_downloads_historical_content() -> None:
    """Pick the oldest visible version from history; download it; verify it's parseable."""
    _skip_if_no_harness()
    versions = history(HARNESS_README_URL, profile=HARNESS_PROFILE, limit=10)
    if not versions:
        pytest.skip("No versions available on the harness README")
    target_id = versions[-1]["id"]  # oldest
    local_path = get_version(HARNESS_README_URL, version_id=target_id, profile=HARNESS_PROFILE)
    try:
        # Should be a real file with readable Markdown content
        content = Path(local_path).read_text(encoding="utf-8")
        assert "harness" in content.lower() or len(content) > 0
        # Naming preserved extension + carries version-id infix
        assert local_path.endswith(".md")
        assert f"_v{target_id.replace('/', '_')}_" in Path(local_path).name
    finally:
        Path(local_path).unlink(missing_ok=True)


def test_history_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        history("", profile=HARNESS_PROFILE)


def test_get_version_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        get_version("", "1.0", profile=HARNESS_PROFILE)
    with pytest.raises(ValueError, match="non-empty version_id"):
        get_version(HARNESS_README_URL, "", profile=HARNESS_PROFILE)
