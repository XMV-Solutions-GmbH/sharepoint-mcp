# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Harness tests for sp_drive_change_track (#51) — delta-query change tracking.

End-to-end lifecycle:
1. First call returns full item list + initial cursor.
2. Publish a new file.
3. Second call with cursor sees that file (and only that file, modulo
   any other concurrent activity on the harness sandbox).
4. Cursor advances and stays opaque.
"""

from __future__ import annotations

import uuid
from collections.abc import Iterator
from pathlib import Path

import httpx
import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token
from sharepoint_mcp.tools._common import (
    GRAPH_BASE,
    parse_sharepoint_url,
    resolve_drive_item,
    resolve_site_id,
)
from sharepoint_mcp.tools.changes import changes
from sharepoint_mcp.tools.publish import publish

HARNESS_PROFILE = "harness"
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
def published_file(tmp_path: Path) -> Iterator[tuple[str, str]]:
    """Yield (name, url) of a newly-published throwaway file; clean up at teardown."""
    _skip_if_no_harness()
    name = f"changes-harness-{uuid.uuid4().hex[:8]}.txt"
    url = f"{HARNESS_FOLDER_URL}/{name}"
    local = tmp_path / name
    local.write_text("changes harness\n", encoding="utf-8")
    publish(str(local), HARNESS_FOLDER_URL, name=name, profile=HARNESS_PROFILE)
    try:
        yield name, url
    finally:
        _delete_via_graph(url)


def test_first_call_returns_items_and_cursor() -> None:
    """Initial call: full list of items currently in the drive + a cursor."""
    _skip_if_no_harness()
    out = changes(HARNESS_SITE_URL, profile=HARNESS_PROFILE)
    assert isinstance(out, dict)
    assert "items" in out
    assert "cursor" in out
    assert isinstance(out["items"], list)
    assert out["cursor"], "expected a non-empty cursor on initial sync"


def test_second_call_sees_a_newly_published_file(
    published_file: tuple[str, str],
) -> None:
    """Lifecycle: get cursor -> publish file -> next call sees the file."""
    name, _ = published_file

    # The fixture already published. Take a cursor BEFORE publishing
    # would require ordering changes; instead, do the canonical
    # workflow: take cursor first, then publish, then call again.
    # We re-publish here under a different name to demonstrate the
    # delta cleanly.
    initial = changes(HARNESS_SITE_URL, profile=HARNESS_PROFILE)
    cursor = initial["cursor"]

    second_name = f"delta-{uuid.uuid4().hex[:8]}.txt"
    second_url = f"{HARNESS_FOLDER_URL}/{second_name}"
    try:
        local = Path("/tmp") / second_name
        local.write_text("delta harness\n", encoding="utf-8")
        publish(str(local), HARNESS_FOLDER_URL, name=second_name, profile=HARNESS_PROFILE)
        try:
            local.unlink(missing_ok=True)
        except OSError:
            pass

        delta = changes(HARNESS_SITE_URL, since=cursor, profile=HARNESS_PROFILE)
        names = {it["name"] for it in delta["items"]}
        # The fixture-published file was already in the cursor's snapshot;
        # we expect ONLY items changed since: at least the new file we
        # just published. Other tests in the same harness run might also
        # have committed during this window, so we don't assert exact
        # equality — just presence + cursor advance.
        assert second_name in names, (
            f"newly-published {second_name!r} not in delta items: {sorted(names)!r}"
        )
        assert delta["cursor"]
        assert delta["cursor"] != cursor, "cursor should advance after activity"
        # Suppress unused-name lint
        _ = name
    finally:
        _delete_via_graph(second_url)


def test_validation_does_not_need_harness() -> None:
    with pytest.raises(ValueError, match="non-empty scope_url"):
        changes("", profile=HARNESS_PROFILE)
    with pytest.raises(ValueError, match="site URL"):
        changes(f"{HARNESS_SITE_URL}/Shared Documents/x.docx", profile=HARNESS_PROFILE)


# ---------------------------------------------------------------------
# Helper
# ---------------------------------------------------------------------


def _delete_via_graph(url: str) -> None:
    try:
        token = get_token(HARNESS_PROFILE)
    except AuthRequiredError:
        return
    headers = {"Authorization": f"Bearer {token}"}
    try:
        with httpx.Client(timeout=30.0) as client:
            hostname, site_path, item_path = parse_sharepoint_url(url)
            site_id = resolve_site_id(client, hostname, site_path, headers=headers)
            drive_id, item_id = resolve_drive_item(client, site_id, item_path, headers=headers)
            client.delete(
                f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}",
                headers=headers,
            )
    except (httpx.HTTPError, KeyError):
        pass
