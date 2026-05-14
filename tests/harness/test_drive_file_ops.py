# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_delete_file, sp_move_file, sp_copy_file (#92, #95, #96).

These tests run against the real SharePoint harness sandbox and validate the
Graph API contract that unit tests with mocks cannot catch.
"""

from __future__ import annotations

import base64
import time
from collections.abc import Iterator

import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.tools.copy_file import copy_file
from sharepoint_mcp.tools.create_folder import create_folder
from sharepoint_mcp.tools.delete_file import delete_file
from sharepoint_mcp.tools.move_file import move_file
from sharepoint_mcp.tools.trash import trash_list
from sharepoint_mcp.tools.upload_new_file import upload_new_file

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


def _txt(label: str) -> str:
    """Minimal text file content as base64."""
    return base64.b64encode(f"# Harness test\n{label}\n".encode()).decode()


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
    upload_new_file(HARNESS_SITE_URL, path, _txt("delete me"), profile=HARNESS_PROFILE)

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
    upload_new_file(HARNESS_SITE_URL, src_path, _txt("move me"), profile=HARNESS_PROFILE)

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
    upload_new_file(HARNESS_SITE_URL, src_path, _txt("rename me"), profile=HARNESS_PROFILE)

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
    upload_new_file(HARNESS_SITE_URL, src_path, _txt("template"), profile=HARNESS_PROFILE)

    dst_path = f"{dst_folder}/instance.txt"
    result = copy_file(HARNESS_SITE_URL, src_path, dst_path, profile=HARNESS_PROFILE)

    assert result["copied"] is True
    assert result["source"] == src_path
    assert result["destination"] == dst_path
    assert result["web_url"]
