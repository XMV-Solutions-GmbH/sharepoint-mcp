# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_download_binary against the harness sandbox.

Skipped when no harness credentials are present.
"""

from __future__ import annotations

import base64

import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.tools.download_binary import download_binary

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


def test_sp_download_binary_returns_envelope() -> None:
    """download_binary returns a dict with the expected keys."""
    _skip_if_no_harness()
    result = download_binary(HARNESS_README_URL, profile=HARNESS_PROFILE)
    assert "filename" in result
    assert "mime_type" in result
    assert "size_bytes" in result
    assert "base64" in result


def test_sp_download_binary_base64_decodes_to_markdown() -> None:
    """The seed README.md is readable text once decoded."""
    _skip_if_no_harness()
    result = download_binary(HARNESS_README_URL, profile=HARNESS_PROFILE)
    content = base64.b64decode(result["base64"]).decode("utf-8")
    assert "sharepoint-mcp harness sandbox" in content


def test_sp_download_binary_size_matches_content() -> None:
    """size_bytes matches the actual length of the decoded content."""
    _skip_if_no_harness()
    result = download_binary(HARNESS_README_URL, profile=HARNESS_PROFILE)
    decoded = base64.b64decode(result["base64"])
    assert result["size_bytes"] == len(decoded)


def test_sp_download_binary_filename_matches() -> None:
    """filename field matches the leaf name of the URL."""
    _skip_if_no_harness()
    result = download_binary(HARNESS_README_URL, profile=HARNESS_PROFILE)
    assert result["filename"] == "README.md"


def test_sp_download_binary_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        download_binary("", profile=HARNESS_PROFILE)


def test_sp_download_binary_site_url_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="site/folder URL"):
        download_binary(HARNESS_SITE_URL, profile=HARNESS_PROFILE)
