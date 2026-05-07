# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for the shared helpers in tools/_common.py.

Covers the library-fallback behaviour added for #48 — the
"transparently route URLs into non-default document libraries"
behaviour that every read/write tool depends on.
"""

from __future__ import annotations

import httpx
import pytest
import respx

from sharepoint_mcp.tools._common import (
    GRAPH_BASE,
    _find_drive_id_by_name,
    list_site_drives,
    parse_sharepoint_url,
    resolve_drive_item,
    resolve_drive_item_full,
)

SITE_ID = "contoso.sharepoint.com,site,web"
HEADERS = {"Authorization": "Bearer T"}


# ---------------------------------------------------------------------
# parse_sharepoint_url — happy paths
# ---------------------------------------------------------------------


def test_parse_strips_default_library_segment() -> None:
    host, site_path, item_path = parse_sharepoint_url(
        "https://contoso.sharepoint.com/sites/foo/Shared Documents/policies/iso.docx"
    )
    assert host == "contoso.sharepoint.com"
    assert site_path == "/sites/foo"
    assert item_path == "policies/iso.docx"


def test_parse_keeps_non_default_library_segment() -> None:
    """SiteAssets is not a default library prefix; preserved in item_path."""
    _, _, item_path = parse_sharepoint_url(
        "https://contoso.sharepoint.com/sites/foo/SiteAssets/icon.png"
    )
    assert item_path == "SiteAssets/icon.png"


# ---------------------------------------------------------------------
# resolve_drive_item_full — default-drive happy path
# ---------------------------------------------------------------------


@respx.mock
def test_resolve_drive_item_full_returns_default_drive_item() -> None:
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/policies/iso.docx").respond(
        json={
            "id": "01ITEM",
            "name": "iso.docx",
            "parentReference": {"driveId": "DEFAULT"},
        }
    )
    with httpx.Client() as c:
        item = resolve_drive_item_full(c, SITE_ID, "policies/iso.docx", headers=HEADERS)
    assert item["id"] == "01ITEM"
    assert item["parentReference"]["driveId"] == "DEFAULT"


# ---------------------------------------------------------------------
# resolve_drive_item_full — library fallback (#48)
# ---------------------------------------------------------------------


@respx.mock
def test_resolve_drive_item_full_falls_back_to_library_on_404() -> None:
    """Default drive doesn't have SiteAssets/icon.png → fallback finds the
    SiteAssets library and retries against its root."""
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/SiteAssets/icon.png").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drives").respond(
        json={
            "value": [
                {"id": "DEF", "name": "Documents"},
                {"id": "SITEASSETS", "name": "Site Assets"},
                {"id": "SA2", "name": "SiteAssets"},
            ],
        },
    )
    respx.get(f"{GRAPH_BASE}/drives/SA2/root:/icon.png").respond(
        json={"id": "ICON", "name": "icon.png", "parentReference": {"driveId": "SA2"}},
    )
    with httpx.Client() as c:
        item = resolve_drive_item_full(c, SITE_ID, "SiteAssets/icon.png", headers=HEADERS)
    assert item["id"] == "ICON"
    assert item["parentReference"]["driveId"] == "SA2"


@respx.mock
def test_resolve_drive_item_full_library_match_is_case_insensitive() -> None:
    """URL casing varies (Site Assets / SiteAssets / siteassets); fallback
    must match regardless."""
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/siteassets/x.png").respond(404)
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drives").respond(
        json={"value": [{"id": "SA", "name": "SiteAssets"}]},
    )
    respx.get(f"{GRAPH_BASE}/drives/SA/root:/x.png").respond(
        json={"id": "X", "parentReference": {"driveId": "SA"}},
    )
    with httpx.Client() as c:
        item = resolve_drive_item_full(c, SITE_ID, "siteassets/x.png", headers=HEADERS)
    assert item["id"] == "X"


@respx.mock
def test_resolve_drive_item_full_library_only_path_resolves_to_drive_root() -> None:
    """URL like https://x/sites/foo/SiteAssets (no sub-path) → fallback
    returns the library's root driveItem."""
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/SiteAssets").respond(404)
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drives").respond(
        json={"value": [{"id": "SA", "name": "SiteAssets"}]},
    )
    respx.get(f"{GRAPH_BASE}/drives/SA/root").respond(
        json={"id": "SA-root", "parentReference": {"driveId": "SA"}},
    )
    with httpx.Client() as c:
        item = resolve_drive_item_full(c, SITE_ID, "SiteAssets", headers=HEADERS)
    assert item["id"] == "SA-root"


@respx.mock
def test_resolve_drive_item_full_no_matching_library_propagates_404() -> None:
    """If no library matches the first segment, the original 404 propagates."""
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/Bogus/foo.txt").respond(404)
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drives").respond(
        json={"value": [{"id": "DEF", "name": "Documents"}]},
    )
    with httpx.Client() as c, pytest.raises(httpx.HTTPStatusError):
        resolve_drive_item_full(c, SITE_ID, "Bogus/foo.txt", headers=HEADERS)


@respx.mock
def test_resolve_drive_item_full_disable_fallback_propagates_404_directly() -> None:
    """allow_library_fallback=False: 404 raises immediately, no /drives lookup."""
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/anything.txt").respond(404)
    with httpx.Client() as c, pytest.raises(httpx.HTTPStatusError):
        resolve_drive_item_full(
            c, SITE_ID, "anything.txt", headers=HEADERS, allow_library_fallback=False
        )


# ---------------------------------------------------------------------
# resolve_drive_item — wrapper-style return
# ---------------------------------------------------------------------


@respx.mock
def test_resolve_drive_item_returns_tuple() -> None:
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/x.docx").respond(
        json={"id": "I", "parentReference": {"driveId": "D"}},
    )
    with httpx.Client() as c:
        drive_id, item_id = resolve_drive_item(c, SITE_ID, "x.docx", headers=HEADERS)
    assert (drive_id, item_id) == ("D", "I")


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------


@respx.mock
def test_find_drive_id_by_name_returns_None_when_no_match() -> None:
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drives").respond(json={"value": []})
    with httpx.Client() as c:
        assert _find_drive_id_by_name(c, SITE_ID, "anything", headers=HEADERS) is None


@respx.mock
def test_list_site_drives_returns_dicts() -> None:
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drives").respond(
        json={"value": [{"id": "D1", "name": "X"}, "not-a-dict"]},
    )
    with httpx.Client() as c:
        result = list_site_drives(c, SITE_ID, headers=HEADERS)
    assert len(result) == 1
    assert result[0]["id"] == "D1"
