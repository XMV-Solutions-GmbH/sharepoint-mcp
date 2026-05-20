# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_share_permission_list (#46)."""

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


def test_permissions_on_site_returns_normalised_shape() -> None:
    """Site-level permissions require Sites.FullControl.All scope, which the
    delegated harness user typically lacks. Skip on 403 (well-known limitation,
    documented in tool description); on success, validate the wire shape."""
    import httpx as _httpx

    _skip_if_no_harness()
    try:
        out = permissions(HARNESS_SITE_URL, profile=HARNESS_PROFILE)
    except _httpx.HTTPStatusError as exc:
        if exc.response.status_code == 403:
            pytest.skip(
                "Site-level permissions require Sites.FullControl.All; the "
                "harness account has only ReadWrite.All. Validated locally "
                "via mocked unit tests."
            )
        raise
    assert isinstance(out, list)
    for entry in out:
        assert "id" in entry
        assert "roles" in entry
        assert "grantee" in entry
        assert "inherited" in entry


def test_permissions_on_known_file_returns_list() -> None:
    """File-level permissions are usually accessible to the file's owner via
    the same scopes our harness user has (Files.ReadWrite.All)."""
    import httpx as _httpx

    _skip_if_no_harness()
    try:
        out = permissions(HARNESS_README_URL, profile=HARNESS_PROFILE)
    except _httpx.HTTPStatusError as exc:
        if exc.response.status_code == 403:
            pytest.skip(
                "File-level permissions are gated on this tenant; harness "
                "user lacks the required scope. Wire shape validated via "
                "mocked unit tests."
            )
        raise
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
