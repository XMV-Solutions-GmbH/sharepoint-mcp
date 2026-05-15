# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for SharePoint modern Pages tools (#45).

Edge cases deliberately covered:
- URL parser: case-insensitive SitePages segment, /teams/ root,
  URL-encoded names with spaces and apostrophes, missing segments.
- Resolver: $filter escapes single quotes, raises PageNotFoundError
  on empty list, raises on response missing id.
- Normaliser: canvas_layout omitted in list responses, included in
  read; lastModifiedBy with no displayName but an email; tolerant
  of malformed lastModifiedBy / non-dict canvas / non-list value.
"""

from __future__ import annotations

import time
from collections.abc import Iterator

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.pages import (
    PageNotFoundError,
    _extract_pages,
    _one_page,
    _resolve_page_id,
    page_read,
    pages_list,
    parse_page_url,
)


class _MemStore:
    def __init__(self, value: bytes | None) -> None:
        self._v = value

    def get(self, profile: str) -> bytes | None:
        return self._v

    def set(self, profile: str, value: bytes) -> None:
        self._v = value

    def delete(self, profile: str) -> None:
        self._v = None


@pytest.fixture
def store_with_fresh_token(monkeypatch: pytest.MonkeyPatch) -> Iterator[None]:
    cached = CachedToken(
        access_token="AT-test",
        refresh_token="RT-test",
        expires_at=time.time() + 3600,
        scope="",
    )
    monkeypatch.setattr(
        "sharepoint_mcp.auth.get_token_store",
        lambda: _MemStore(cached.to_json().encode()),
    )
    monkeypatch.delenv("SP_AUTH_MODE", raising=False)
    monkeypatch.delenv("SP_CLIENT_SECRET", raising=False)
    yield


SITE_HOST = "contoso.sharepoint.com"
SITE_PATH = "/sites/foo"
SITE_ID = "contoso.sharepoint.com,site,web"
SITE_URL = f"https://{SITE_HOST}{SITE_PATH}"
PAGE_NAME = "Onboarding.aspx"
PAGE_URL = f"{SITE_URL}/SitePages/{PAGE_NAME}"
PAGE_GUID = "page-guid-123"


def _mock_site_lookup() -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(
        json={"id": SITE_ID},
    )


def _mock_resolve_page() -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/pages").respond(
        json={"value": [{"id": PAGE_GUID, "name": PAGE_NAME}]},
    )


# ---------------------------------------------------------------------
# parse_page_url — happy path + edge cases
# ---------------------------------------------------------------------


def test_parse_page_url_basic() -> None:
    assert parse_page_url(PAGE_URL) == (SITE_HOST, SITE_PATH, PAGE_NAME)


def test_parse_page_url_url_decodes_name() -> None:
    url = f"{SITE_URL}/SitePages/My%20Page.aspx"
    assert parse_page_url(url) == (SITE_HOST, SITE_PATH, "My Page.aspx")


def test_parse_page_url_url_decodes_apostrophe() -> None:
    """SharePoint allows apostrophes in page titles; the URL form percent-encodes."""
    url = f"{SITE_URL}/SitePages/Bob%27s%20Notes.aspx"
    assert parse_page_url(url) == (SITE_HOST, SITE_PATH, "Bob's Notes.aspx")


def test_parse_page_url_case_insensitive_sitepages_segment() -> None:
    url = f"{SITE_URL}/sitepages/Lowercase.aspx"
    assert parse_page_url(url) == (SITE_HOST, SITE_PATH, "Lowercase.aspx")


def test_parse_page_url_works_with_teams_root() -> None:
    url = f"https://{SITE_HOST}/teams/eng/SitePages/Runbook.aspx"
    assert parse_page_url(url) == (SITE_HOST, "/teams/eng", "Runbook.aspx")


def test_parse_page_url_rejects_empty() -> None:
    with pytest.raises(ValueError, match="non-empty"):
        parse_page_url("")


def test_parse_page_url_rejects_relative() -> None:
    with pytest.raises(ValueError, match="absolute URL"):
        parse_page_url("/sites/foo/SitePages/X.aspx")


def test_parse_page_url_rejects_no_sitepages_segment() -> None:
    with pytest.raises(ValueError, match="SitePages"):
        parse_page_url(f"{SITE_URL}/Shared Documents/foo.docx")


def test_parse_page_url_rejects_sitepages_without_name() -> None:
    with pytest.raises(ValueError, match="SitePages"):
        parse_page_url(f"{SITE_URL}/SitePages")


def test_parse_page_url_rejects_personal_root() -> None:
    """Personal OneDrive URLs aren't sites; reject."""
    with pytest.raises(ValueError, match="must look like"):
        parse_page_url(f"https://{SITE_HOST}/personal/me/SitePages/X.aspx")


# ---------------------------------------------------------------------
# pages_list
# ---------------------------------------------------------------------


@respx.mock
def test_pages_list_returns_normalised_entries(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/pages").respond(
        json={
            "value": [
                {
                    "id": "p1",
                    "name": "Home.aspx",
                    "title": "Welcome",
                    "webUrl": "https://x/sites/foo/SitePages/Home.aspx",
                    "description": "landing page",
                    "pageLayout": "home",
                    "thumbnailWebUrl": "https://x/thumb.png",
                    "lastModifiedDateTime": "2026-04-01T08:00:00Z",
                    "lastModifiedBy": {"user": {"displayName": "Alice"}},
                },
            ],
        },
    )
    [out] = pages_list(SITE_URL)
    assert out["id"] == "p1"
    assert out["title"] == "Welcome"
    assert out["page_layout"] == "home"
    assert out["last_modified_by"] == "Alice"
    assert "canvas_layout" not in out  # list excludes canvas


def test_pages_list_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty site_url"):
        pages_list("")


def test_pages_list_rejects_file_url() -> None:
    with pytest.raises(ValueError, match="expects a site URL"):
        pages_list(f"{SITE_URL}/Shared Documents/x.docx")


# ---------------------------------------------------------------------
# page_read
# ---------------------------------------------------------------------


@respx.mock
def test_page_read_returns_page_with_canvas(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    _mock_resolve_page()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/pages/{PAGE_GUID}/microsoft.graph.sitePage").respond(
        json={
            "id": PAGE_GUID,
            "name": PAGE_NAME,
            "title": "Onboarding",
            "canvasLayout": {
                "horizontalSections": [{"layout": "oneColumn", "columns": [{"id": "col-1"}]}],
            },
        },
    )
    out = page_read(PAGE_URL)
    assert out["id"] == PAGE_GUID
    assert "canvas_layout" in out
    assert out["canvas_layout"]["horizontalSections"][0]["layout"] == "oneColumn"


@respx.mock
def test_page_read_passes_filter_with_escaped_apostrophe(store_with_fresh_token: None) -> None:
    """OData $filter requires '' to escape an apostrophe in the literal."""
    del store_with_fresh_token
    _mock_site_lookup()
    route = respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/pages").respond(
        json={"value": [{"id": "G", "name": "Bob's Notes.aspx"}]},
    )
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/pages/G/microsoft.graph.sitePage").respond(
        json={"id": "G", "name": "Bob's Notes.aspx", "canvasLayout": {}},
    )
    page_read(f"{SITE_URL}/SitePages/Bob%27s%20Notes.aspx")
    url = str(route.calls.last.request.url)
    assert "Bob%27%27s" in url or "Bob''s" in url


@respx.mock
def test_page_read_raises_PageNotFoundError_when_filter_returns_empty(
    store_with_fresh_token: None,
) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/pages").respond(json={"value": []})
    with pytest.raises(PageNotFoundError, match="No SharePoint page named"):
        page_read(PAGE_URL)


@respx.mock
def test_page_read_raises_PageNotFoundError_when_response_missing_id(
    store_with_fresh_token: None,
) -> None:
    """Defensive: if Graph returns a malformed entry without id."""
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/pages").respond(
        json={"value": [{"name": PAGE_NAME}]},  # no id
    )
    with pytest.raises(PageNotFoundError, match="missing the id field"):
        page_read(PAGE_URL)


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------


@respx.mock
def test_resolve_page_id_escapes_apostrophe_in_filter() -> None:
    route = respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/pages").respond(
        json={"value": [{"id": "G", "name": "Bob's.aspx"}]},
    )
    with httpx.Client() as c:
        out = _resolve_page_id(c, SITE_ID, "Bob's.aspx", headers={})
    assert out == "G"
    sent = str(route.calls.last.request.url)
    # The single apostrophe must be doubled in the OData literal
    assert "Bob%27%27s" in sent or "Bob''s" in sent


@respx.mock
def test_resolve_page_id_raises_when_value_field_is_not_list() -> None:
    """Defensive: malformed response without a list value."""
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/pages").respond(
        json={"value": "not-a-list"},
    )
    with httpx.Client() as c, pytest.raises(PageNotFoundError):
        _resolve_page_id(c, SITE_ID, "x.aspx", headers={})


def test_extract_pages_handles_missing_value() -> None:
    assert _extract_pages({}, include_canvas=False) == []


def test_extract_pages_filters_non_dict_entries() -> None:
    payload = {"value": [{"id": "a", "name": "A.aspx"}, "junk", None]}
    [out] = _extract_pages(payload, include_canvas=False)
    assert out["id"] == "a"


def test_one_page_handles_missing_lastModifiedBy() -> None:
    out = _one_page({"id": "p"}, include_canvas=False)
    assert out["last_modified_by"] == ""


def test_one_page_handles_email_only_user() -> None:
    out = _one_page(
        {"id": "p", "lastModifiedBy": {"user": {"email": "x@y.com"}}},
        include_canvas=False,
    )
    assert out["last_modified_by"] == "x@y.com"


def test_one_page_handles_non_dict_lastModifiedBy() -> None:
    """Defensive: malformed payload doesn't crash."""
    out = _one_page({"id": "p", "lastModifiedBy": "not-a-dict"}, include_canvas=False)
    assert out["last_modified_by"] == ""


def test_one_page_handles_non_dict_canvasLayout() -> None:
    """Graph occasionally returns canvasLayout: null on empty pages."""
    out = _one_page({"id": "p", "canvasLayout": None}, include_canvas=True)
    assert out["canvas_layout"] == {}


def test_one_page_normalises_missing_fields() -> None:
    out = _one_page({}, include_canvas=False)
    assert out["id"] == ""
    assert out["title"] == ""
    assert out["page_layout"] == ""
    assert out["thumbnail_web_url"] == ""


def test_one_page_omits_canvas_when_include_canvas_false() -> None:
    out = _one_page({"id": "p", "canvasLayout": {"x": 1}}, include_canvas=False)
    assert "canvas_layout" not in out
