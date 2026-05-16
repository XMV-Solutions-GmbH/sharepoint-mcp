# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_delete_file, sp_move_file, sp_copy_file (#92, #95, #96).

These tests run against the real SharePoint harness sandbox and validate the
Graph API contract that unit tests with mocks cannot catch.
"""

from __future__ import annotations

import tempfile
import time
from collections.abc import Iterator
from pathlib import Path

import httpx
import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.tools.copy_file import copy_file
from sharepoint_mcp.tools.create_folder import create_folder
from sharepoint_mcp.tools.delete_file import delete_file
from sharepoint_mcp.tools.move_file import move_file
from sharepoint_mcp.tools.publish import publish
from sharepoint_mcp.tools.trash import trash_list

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


def _publish_text(path: str, label: str) -> None:
    """Write a small text file to a temp file and publish it at `path`
    (relative to the default document library root)."""
    p = Path(path)
    folder_url = f"{HARNESS_SITE_URL}/Shared Documents/{p.parent}"
    content = f"# Harness test\n{label}\n".encode()
    with tempfile.NamedTemporaryFile(suffix=p.suffix, delete=False) as f:
        f.write(content)
        tmp = Path(f.name)
    try:
        publish(str(tmp), folder_url, name=p.name, profile=HARNESS_PROFILE)
    finally:
        tmp.unlink(missing_ok=True)


@pytest.fixture
def harness_root() -> Iterator[str]:
    """Unique top-level folder; deleted from SharePoint after the test via sp_delete_file."""
    folder = f"harness-drive-ops-{int(time.time())}"
    yield folder
    # Best-effort cleanup — swallow errors so fixture teardown never masks test failures.
    try:
        delete_file(HARNESS_SITE_URL, folder, profile=HARNESS_PROFILE)
    except Exception:
        pass


# ------------------------------------------------------------------
# sp_delete_file
# ------------------------------------------------------------------


def test_delete_file_moves_to_recycle_bin(harness_root: str) -> None:
    _skip_if_no_harness()
    create_folder(HARNESS_SITE_URL, harness_root, profile=HARNESS_PROFILE)
    path = f"{harness_root}/to-delete.txt"
    _publish_text(path, "delete me")

    result = delete_file(HARNESS_SITE_URL, path, profile=HARNESS_PROFILE)

    assert result["deleted"] is True
    assert result["path"] == path

    # Verify item ended up in the recycle bin.
    trash = trash_list(HARNESS_SITE_URL, profile=HARNESS_PROFILE)
    names = [item["name"] for item in trash]
    assert "to-delete.txt" in names


def test_delete_file_folder_moves_to_recycle_bin(harness_root: str) -> None:
    _skip_if_no_harness()
    folder = f"{harness_root}/sub"
    create_folder(HARNESS_SITE_URL, folder, profile=HARNESS_PROFILE)

    # Delete the sub-folder itself (not just its contents).
    result = delete_file(HARNESS_SITE_URL, folder, profile=HARNESS_PROFILE)

    assert result["deleted"] is True


# ------------------------------------------------------------------
# sp_move_file
# ------------------------------------------------------------------


def test_move_file_to_different_folder(harness_root: str) -> None:
    _skip_if_no_harness()
    src_folder = f"{harness_root}/src"
    dst_folder = f"{harness_root}/dst"
    create_folder(HARNESS_SITE_URL, src_folder, profile=HARNESS_PROFILE)
    create_folder(HARNESS_SITE_URL, dst_folder, profile=HARNESS_PROFILE)

    src_path = f"{src_folder}/move-me.txt"
    _publish_text(src_path, "move me")

    dst_path = f"{dst_folder}/move-me.txt"
    result = move_file(HARNESS_SITE_URL, src_path, dst_path, profile=HARNESS_PROFILE)

    assert result["moved"] is True
    assert result["source"] == src_path
    assert result["destination"] == dst_path
    assert result["web_url"]


def test_move_file_rename_in_place(harness_root: str) -> None:
    _skip_if_no_harness()
    create_folder(HARNESS_SITE_URL, harness_root, profile=HARNESS_PROFILE)
    src_path = f"{harness_root}/before.txt"
    _publish_text(src_path, "rename me")

    dst_path = f"{harness_root}/after.txt"
    result = move_file(HARNESS_SITE_URL, src_path, dst_path, profile=HARNESS_PROFILE)

    assert result["moved"] is True
    assert "after.txt" in result["web_url"]


# ------------------------------------------------------------------
# sp_copy_file
# ------------------------------------------------------------------


def test_copy_file_creates_independent_copy(harness_root: str) -> None:
    _skip_if_no_harness()
    src_folder = f"{harness_root}/templates"
    dst_folder = f"{harness_root}/projects"
    create_folder(HARNESS_SITE_URL, src_folder, profile=HARNESS_PROFILE)
    create_folder(HARNESS_SITE_URL, dst_folder, profile=HARNESS_PROFILE)

    src_path = f"{src_folder}/template.txt"
    _publish_text(src_path, "template")

    dst_path = f"{dst_folder}/instance.txt"
    result = copy_file(HARNESS_SITE_URL, src_path, dst_path, profile=HARNESS_PROFILE)

    assert result["copied"] is True
    assert result["source"] == src_path
    assert result["destination"] == dst_path
    assert result["web_url"]


# ------------------------------------------------------------------
# Error paths — verifying real Graph API behaviour
#
# These tests do NOT create any folders or files; they verify that
# the real Microsoft Graph API returns 404 for non-existent resources.
# Mocks cannot catch this: a mock returns whatever shape the author
# wrote in; only the real API confirms the actual error contract.
# ------------------------------------------------------------------


def test_delete_nonexistent_file_propagates_404() -> None:
    """Graph returns 404 when deleting a path that does not exist."""
    _skip_if_no_harness()
    nonexistent = f"harness-nonexistent-{int(time.time())}/ghost.txt"

    with pytest.raises(httpx.HTTPStatusError) as exc_info:
        delete_file(HARNESS_SITE_URL, nonexistent, profile=HARNESS_PROFILE)

    assert exc_info.value.response.status_code == 404


def test_move_nonexistent_source_propagates_404() -> None:
    """Graph returns 404 when the move source path does not exist."""
    _skip_if_no_harness()
    ts = int(time.time())
    nonexistent = f"harness-nonexistent-{ts}/ghost.txt"
    dest = f"harness-nonexistent-{ts}/moved.txt"

    with pytest.raises(httpx.HTTPStatusError) as exc_info:
        move_file(HARNESS_SITE_URL, nonexistent, dest, profile=HARNESS_PROFILE)

    assert exc_info.value.response.status_code == 404


def test_copy_nonexistent_source_propagates_404() -> None:
    """Graph returns 404 when the copy source path does not exist."""
    _skip_if_no_harness()
    ts = int(time.time())
    nonexistent = f"harness-nonexistent-{ts}/ghost.txt"
    dest = f"harness-nonexistent-{ts}/copy.txt"

    with pytest.raises(httpx.HTTPStatusError) as exc_info:
        copy_file(HARNESS_SITE_URL, nonexistent, dest, profile=HARNESS_PROFILE)

    assert exc_info.value.response.status_code == 404
