# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_publish against the real harness sandbox."""

from __future__ import annotations

import time
from collections.abc import Iterator
from pathlib import Path

import httpx
import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.publish import publish

HARNESS_PROFILE = "harness"
HARNESS_SITE_URL = "https://xmvsolutions.sharepoint.com/sites/sharepoint-mcp-harness"
HARNESS_DRAFTS_URL = f"{HARNESS_SITE_URL}/Shared Documents/drafts"


def _skip_if_no_harness() -> None:
    try:
        get_token(HARNESS_PROFILE)
    except AuthRequiredError as exc:
        pytest.skip(
            f"Harness credentials not available: {exc}. "
            "Run `uv run mcp-server-sharepoint login --profile harness` to populate.",
        )


@pytest.fixture
def cleanup_published_files() -> Iterator[list[str]]:
    """Track published filenames and delete them from SharePoint after the test.

    Each test that uses this fixture appends to the yielded list whatever it
    publishes. After the test, we DELETE those items via Graph so the harness
    doesn't accumulate cruft across runs.
    """
    published: list[str] = []
    yield published
    if not published:
        return
    try:
        token = get_token(HARNESS_PROFILE)
    except AuthRequiredError:
        return
    headers = {"Authorization": f"Bearer {token}"}
    # Resolve site id once
    with httpx.Client(timeout=30.0, follow_redirects=True) as client:
        try:
            site_response = client.get(
                f"{GRAPH_BASE}/sites/xmvsolutions.sharepoint.com:/sites/sharepoint-mcp-harness",
                headers=headers,
            )
            site_response.raise_for_status()
            site_id = site_response.json()["id"]
        except httpx.HTTPError:
            return
        for filename in published:
            try:
                client.delete(
                    f"{GRAPH_BASE}/sites/{site_id}/drive/root:/drafts/{filename}",
                    headers=headers,
                )
            except httpx.HTTPError:
                pass


def test_publish_creates_new_file_in_drafts(
    tmp_path: Path, cleanup_published_files: list[str]
) -> None:
    """Publishes a fresh file into /drafts. Verifies upload + content + cleanup."""
    _skip_if_no_harness()
    # Unique name per run so cleanup-failures don't poison subsequent runs
    unique = f"harness-publish-{int(time.time())}.md"
    cleanup_published_files.append(unique)

    src = tmp_path / unique
    src.write_text(
        f"# Harness publish test\n\nGenerated at {time.time()}\n",
        encoding="utf-8",
    )

    result = publish(str(src), HARNESS_DRAFTS_URL, profile=HARNESS_PROFILE)

    assert result["name"] == unique
    assert result["web_url"]
    assert result["etag"]
    assert result["size"] > 0


def test_publish_refuses_to_overwrite_existing_file(
    tmp_path: Path, cleanup_published_files: list[str]
) -> None:
    """Publishing twice with the same name raises FileExistsError on the second call."""
    _skip_if_no_harness()
    unique = f"harness-overwrite-{int(time.time())}.md"
    cleanup_published_files.append(unique)

    src = tmp_path / unique
    src.write_text("first version", encoding="utf-8")

    # First publish: succeeds
    publish(str(src), HARNESS_DRAFTS_URL, profile=HARNESS_PROFILE)

    # Second publish with same name: refused
    src.write_text("second version (should never land)", encoding="utf-8")
    with pytest.raises(FileExistsError, match="Use sp_open"):
        publish(str(src), HARNESS_DRAFTS_URL, profile=HARNESS_PROFILE)


def test_publish_validation_does_not_need_harness(tmp_path: Path) -> None:
    with pytest.raises(ValueError, match="non-empty"):
        publish("", HARNESS_DRAFTS_URL, profile=HARNESS_PROFILE)
    with pytest.raises(FileNotFoundError):
        publish(str(tmp_path / "no-such.txt"), HARNESS_DRAFTS_URL, profile=HARNESS_PROFILE)
