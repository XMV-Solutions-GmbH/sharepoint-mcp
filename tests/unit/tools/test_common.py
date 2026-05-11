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
    # Strip-segment retry (#79 fallback 1) — also 404, falls through to library search.
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/icon.png").respond(404)
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
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/x.png").respond(404)
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
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/foo.txt").respond(404)
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


# ---------------------------------------------------------------------
# resolve_drive_item_by_share_url — /shares/{u!base64}/driveItem (#79)
# ---------------------------------------------------------------------


@respx.mock
def test_resolve_drive_item_by_share_url_happy_path() -> None:
    """Encoded webUrl → /shares lookup → driveItem returned."""
    import base64

    from sharepoint_mcp.tools._common import resolve_drive_item_by_share_url

    url = "https://contoso.sharepoint.com/sites/foo/Shared Documents/readme.md"
    encoded = base64.urlsafe_b64encode(url.encode()).rstrip(b"=").decode()
    share_id = f"u!{encoded}"
    respx.get(f"{GRAPH_BASE}/shares/{share_id}/driveItem").respond(
        json={
            "id": "01XYZ",
            "name": "readme.md",
            "parentReference": {"driveId": "b!drive"},
        },
    )
    with httpx.Client() as c:
        item = resolve_drive_item_by_share_url(c, url, headers=HEADERS)
    assert item["id"] == "01XYZ"
    assert item["parentReference"]["driveId"] == "b!drive"


@respx.mock
def test_resolve_drive_item_by_share_url_localized_german() -> None:
    """The German tenant case from #79: 'Freigegebene Dokumente' in
    the URL resolves cleanly via /shares (whereas the legacy
    /sites/{id}/drive/root:/... path returns 404)."""
    import base64

    from sharepoint_mcp.tools._common import resolve_drive_item_by_share_url

    url = "https://contoso.sharepoint.com/sites/Foo/Freigegebene Dokumente/Finanzen/steuer.pdf"
    encoded = base64.urlsafe_b64encode(url.encode("utf-8")).rstrip(b"=").decode()
    share_id = f"u!{encoded}"
    respx.get(f"{GRAPH_BASE}/shares/{share_id}/driveItem").respond(
        json={
            "id": "01DEU",
            "name": "steuer.pdf",
            "parentReference": {"driveId": "b!drive-de"},
        },
    )
    with httpx.Client() as c:
        item = resolve_drive_item_by_share_url(c, url, headers=HEADERS)
    assert item["id"] == "01DEU"


@respx.mock
def test_resolve_drive_item_by_share_url_404_propagates() -> None:
    """Caller catches HTTPStatusError to decide on fallback strategy."""
    import re

    from sharepoint_mcp.tools._common import resolve_drive_item_by_share_url

    respx.get(re.compile(rf"{re.escape(GRAPH_BASE)}/shares/u!.*?/driveItem")).respond(404)
    with httpx.Client() as c:
        with pytest.raises(httpx.HTTPStatusError):
            resolve_drive_item_by_share_url(
                c,
                "https://contoso.sharepoint.com/sites/foo/bad",
                headers=HEADERS,
            )


def test_resolve_drive_item_by_share_url_encoding_no_padding() -> None:
    """Verify the encoded share-id strips trailing `=` padding (per Graph spec)."""
    import base64

    # 28-byte URL produces padded base64; encoding must strip the pad.
    url = "https://x.sharepoint.com/Foo"
    raw = base64.b64encode(url.encode()).decode()
    assert raw.endswith("="), f"test setup: {url!r} should produce padded base64, got {raw!r}"
    safe = base64.urlsafe_b64encode(url.encode()).rstrip(b"=").decode()
    assert not safe.endswith("=")
    # Encoding alphabet matches what Graph expects (urlsafe variant: -_, not +/).
    assert all(c.isalnum() or c in "-_" for c in safe)


# ---------------------------------------------------------------------
# resolve_drive_item_full — localized-library fallback (#79)
# ---------------------------------------------------------------------


@respx.mock
def test_resolve_drive_item_full_strip_localized_library_segment() -> None:
    """Primary 404 → strip first segment, retry against default drive.
    Covers the German `Freigegebene Dokumente` / Italian
    `Documenti condivisi` case from #79."""
    respx.get(
        f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/Freigegebene Dokumente/policies/iso.docx"
    ).respond(404)
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/policies/iso.docx").respond(
        json={"id": "I", "parentReference": {"driveId": "DEFAULT"}},
    )
    with httpx.Client() as c:
        item = resolve_drive_item_full(
            c,
            SITE_ID,
            "Freigegebene Dokumente/policies/iso.docx",
            headers=HEADERS,
        )
    assert item["id"] == "I"


@respx.mock
def test_resolve_drive_item_full_localized_retry_still_404_falls_to_library_search() -> None:
    """If the strip-segment retry ALSO 404s, we fall through to the
    existing `_find_drive_id_by_name` library-search path."""
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/SiteAssets/icon.png").respond(404)
    # Strip-segment retry: also 404 (the path isn't in the default drive).
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/icon.png").respond(404)
    # Library-search finds SiteAssets.
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drives").respond(
        json={"value": [{"id": "SA", "name": "SiteAssets"}]},
    )
    respx.get(f"{GRAPH_BASE}/drives/SA/root:/icon.png").respond(
        json={"id": "ICON", "parentReference": {"driveId": "SA"}},
    )
    with httpx.Client() as c:
        item = resolve_drive_item_full(c, SITE_ID, "SiteAssets/icon.png", headers=HEADERS)
    assert item["id"] == "ICON"
