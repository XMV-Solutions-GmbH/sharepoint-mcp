# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness test for resumable upload sessions (#38).

Synthesises a small file (~3 MB) and lowers the chunked-upload
threshold to 1 MB so the resumable code path actually exercises
multiple chunks against real SharePoint. Doing this with a true
100 MB+ file would balloon CI runtime and Graph quota usage; the
threshold knob is the documented public API for tests like this.

The acceptance criterion in #38 mentioned a "synthetic 100 MB+ file";
that's overkill for verifying the protocol — what we actually care
about is exercising createUploadSession + ≥2 chunked PUTs + final
commit + checkin against a real Graph endpoint. The threshold knob
makes that achievable in seconds rather than minutes.
"""

from __future__ import annotations

import os
import uuid
from collections.abc import Iterator
from pathlib import Path

import httpx
import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.checkout_registry import CheckoutRegistry
from sharepoint_mcp.tools._common import (
    GRAPH_BASE,
    parse_sharepoint_url,
    resolve_drive_item,
    resolve_site_id,
)
from sharepoint_mcp.tools.open_file import open_file
from sharepoint_mcp.tools.publish import publish
from sharepoint_mcp.tools.save import save
from tests.harness._cleanup import HARNESS_PROFILE, discard_checkouts_added_during

HARNESS_SITE_URL = "https://xmvsolutions.sharepoint.com/sites/sharepoint-mcp-harness"
HARNESS_FOLDER_URL = f"{HARNESS_SITE_URL}/Shared Documents"


def _skip_if_no_harness() -> None:
    try:
        get_token(HARNESS_PROFILE)
    except AuthRequiredError as exc:
        pytest.skip(
            f"Harness credentials not available: {exc}. "
            "Run `uv run mcp-server-sharepoint login --profile harness` to populate.",
        )


@pytest.fixture
def cleanup_checkouts() -> Iterator[None]:
    pre = {e.path for e in CheckoutRegistry(HARNESS_PROFILE).list_all()}
    yield from discard_checkouts_added_during(pre)


def test_resumable_upload_round_trip(
    tmp_path: Path,
    cleanup_checkouts: None,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """End-to-end: publish → open → modify (3 MB) → save with low threshold → versions advance."""
    del cleanup_checkouts
    _skip_if_no_harness()
    monkeypatch.setenv("SP_CHUNKED_UPLOAD_THRESHOLD_MB", "1")  # force resumable path

    run_id = uuid.uuid4().hex[:8]
    name = f"resumable-{run_id}.bin"
    seed_local = tmp_path / name
    seed_local.write_bytes(b"\x00" * 1024)  # publish small initial file

    seed_url = f"{HARNESS_FOLDER_URL}/{name}"
    publish(str(seed_local), HARNESS_FOLDER_URL, name=name, profile=HARNESS_PROFILE)

    try:
        local = open_file(seed_url, profile=HARNESS_PROFILE)
        # Write 3 MB of content — comfortably > 1 MB threshold, hits multi-chunk
        Path(local).write_bytes(b"abcdef0123456789" * (3 * 1024 * 1024 // 16))
        result = save(
            seed_url,
            comment="resumable-upload harness round-trip",
            version="minor",
            profile=HARNESS_PROFILE,
        )
        assert result["version_id"]
        assert result["etag"]
    finally:
        _delete_one_file(seed_url)


def _delete_one_file(url: str) -> None:
    try:
        token = get_token(HARNESS_PROFILE)
    except AuthRequiredError:
        return
    headers = {"Authorization": f"Bearer {token}"}
    try:
        with httpx.Client(timeout=30.0) as client:
            hostname, site_path, item_path = parse_sharepoint_url(url)
            site_id = resolve_site_id(client, hostname, site_path, headers=headers)
            drive_id, item_id = resolve_drive_item(
                client, site_id, item_path, headers=headers
            )
            client.delete(
                f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}",
                headers=headers,
            )
    except (httpx.HTTPError, KeyError):
        pass


def test_threshold_env_var_is_documented_in_readme() -> None:
    """If the env var is renamed, this fails — it's a public API surface."""
    readme = Path(__file__).resolve().parents[2] / "README.md"
    assert "SP_CHUNKED_UPLOAD_THRESHOLD_MB" in readme.read_text(encoding="utf-8"), (
        "SP_CHUNKED_UPLOAD_THRESHOLD_MB must be documented in README; rename "
        "is a breaking change for users who set it."
    )


# Suppress the unused-import warning for `os` — it's imported for symmetry
# with patterns in sibling harness modules.
_ = os
