# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_create_folder (and its integration with sp_upload_new_file).

These tests were absent in v0.6.0 — that omission allowed a URL bug in
sp_create_folder (missing `:` before `/children` in path-based Graph URLs) to
slip past CI. Unit tests only validate consistency of code with itself; the
harness catches API contract violations that mocks cannot.
"""

from __future__ import annotations

import time
from collections.abc import Iterator
from pathlib import Path

import httpx
import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.create_folder import create_folder
from sharepoint_mcp.tools.publish import publish

HARNESS_PROFILE = "harness"
HARNESS_SITE_URL = "https://xmvsolutions.sharepoint.com/sites/sharepoint-mcp-harness"
HARNESS_SITE_HOST = "xmvsolutions.sharepoint.com"
HARNESS_SITE_PATH = "/sites/sharepoint-mcp-harness"


def _skip_if_no_harness() -> None:
    try:
        get_token(HARNESS_PROFILE)
    except AuthRequiredError as exc:
        pytest.skip(
            f"Harness credentials not available: {exc}. "
            "Run `uv run mcp-server-sharepoint login --profile harness` to populate.",
        )


def _graph_delete(path: str) -> None:
    """Best-effort DELETE via Graph. Swallows errors (cleanup, not assertions)."""
    try:
        token = get_token(HARNESS_PROFILE)
    except AuthRequiredError:
        return
    headers = {"Authorization": f"Bearer {token}"}
    with httpx.Client(timeout=30.0, follow_redirects=True) as client:
        try:
            site_resp = client.get(
                f"{GRAPH_BASE}/sites/{HARNESS_SITE_HOST}:{HARNESS_SITE_PATH}",
                headers=headers,
            )
            site_resp.raise_for_status()
            site_id = site_resp.json()["id"]
            client.delete(
                f"{GRAPH_BASE}/sites/{site_id}/drive/root:/{path}",
                headers=headers,
            )
        except httpx.HTTPError:
            pass


@pytest.fixture
def harness_test_root() -> Iterator[str]:
    """Yield a unique top-level folder path; delete it from SharePoint after the test."""
    folder = f"harness-create-{int(time.time())}"
    yield folder
    _graph_delete(folder)


# ---------------------------------------------------------------------
# sp_create_folder — real Graph API
# ---------------------------------------------------------------------


def test_create_folder_single_segment(harness_test_root: str) -> None:
    """Create one folder under the default library root."""
    _skip_if_no_harness()
    result = create_folder(HARNESS_SITE_URL, harness_test_root, profile=HARNESS_PROFILE)

    assert harness_test_root in result["created"]
    assert result["already_existed"] == []
    assert result["web_url"]


def test_create_folder_deep_path(harness_test_root: str) -> None:
    """Create a three-level hierarchy in one call."""
    _skip_if_no_harness()
    path = f"{harness_test_root}/L2/L3"
    result = create_folder(HARNESS_SITE_URL, path, profile=HARNESS_PROFILE)

    assert len(result["created"]) == 3
    assert result["already_existed"] == []
    assert result["web_url"]


def test_create_folder_idempotent(harness_test_root: str) -> None:
    """Calling create_folder twice on the same path must not raise."""
    _skip_if_no_harness()
    create_folder(HARNESS_SITE_URL, harness_test_root, profile=HARNESS_PROFILE)

    result = create_folder(HARNESS_SITE_URL, harness_test_root, profile=HARNESS_PROFILE)

    assert result["created"] == []
    assert harness_test_root in result["already_existed"]
    assert result["web_url"]


def test_create_folder_partial_pre_existence(harness_test_root: str) -> None:
    """First segment already exists; second is new — both handled correctly."""
    _skip_if_no_harness()
    create_folder(HARNESS_SITE_URL, harness_test_root, profile=HARNESS_PROFILE)

    result = create_folder(
        HARNESS_SITE_URL,
        f"{harness_test_root}/new-child",
        profile=HARNESS_PROFILE,
    )

    assert result["already_existed"] == [harness_test_root]
    assert result["created"] == [f"{harness_test_root}/new-child"]


# ---------------------------------------------------------------------
# sp_create_folder + sp_upload_new_file integration — real Graph API
# ---------------------------------------------------------------------


def test_create_folder_then_publish(harness_test_root: str, tmp_path: Path) -> None:
    """Create a folder and publish a file into it — the standard new-document workflow."""
    _skip_if_no_harness()
    create_folder(HARNESS_SITE_URL, harness_test_root, profile=HARNESS_PROFILE)

    src = tmp_path / "test.md"
    src.write_text(f"# Harness test\n\nGenerated at {time.time()}\n", encoding="utf-8")
    folder_url = f"{HARNESS_SITE_URL}/Shared Documents/{harness_test_root}"

    result = publish(str(src), folder_url, profile=HARNESS_PROFILE)

    assert result["name"] == "test.md"
    assert result["web_url"]
    assert result["size"] == src.stat().st_size


def test_publish_refuses_overwrite_in_created_folder(
    harness_test_root: str, tmp_path: Path
) -> None:
    """Publishing the same filename twice into the same folder raises FileExistsError."""
    _skip_if_no_harness()
    create_folder(HARNESS_SITE_URL, harness_test_root, profile=HARNESS_PROFILE)

    src = tmp_path / "once.txt"
    src.write_text("first", encoding="utf-8")
    folder_url = f"{HARNESS_SITE_URL}/Shared Documents/{harness_test_root}"

    publish(str(src), folder_url, profile=HARNESS_PROFILE)

    src.write_text("second (must not land)", encoding="utf-8")
    with pytest.raises(FileExistsError, match="sp_open_file"):
        publish(str(src), folder_url, profile=HARNESS_PROFILE)
