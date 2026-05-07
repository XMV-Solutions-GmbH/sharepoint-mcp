# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sharing-link tools (#47).

Includes deliberate edge-case coverage:
- alternate Graph wire shapes for the createLink response
- conservative defaults exercised explicitly
- security-relevant params (anonymous + edit) flow through to the body
- expiration / password forwarded only when set
- response-shape normalisation tolerates malformed payloads
"""

from __future__ import annotations

import time
from collections.abc import Iterator

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.sharing import (
    VALID_LINK_SCOPES,
    VALID_LINK_TYPES,
    _normalise_create_response,
    _normalise_existing_link,
    share_create,
    share_list,
    share_revoke,
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
DRIVE_ID = "b!drive"
ITEM_ID = "01ITEM"
SITE_URL = f"https://{SITE_HOST}{SITE_PATH}"
FILE_URL = f"{SITE_URL}/Shared Documents/policy.docx"
LINK_ID = "perm-link-1"


def _mock_resolution() -> None:
    respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(json={"id": SITE_ID})
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/policy.docx").respond(
        json={"id": ITEM_ID, "parentReference": {"driveId": DRIVE_ID}},
    )


# ---------------------------------------------------------------------
# share_create — happy path + body construction
# ---------------------------------------------------------------------


@respx.mock
def test_share_create_default_view_organization(store_with_fresh_token: None) -> None:
    """Defaults must be conservative: type='view', scope='organization'."""
    del store_with_fresh_token
    _mock_resolution()
    route = respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createLink").respond(
        json={
            "id": "p1",
            "roles": ["read"],
            "link": {
                "type": "view",
                "scope": "organization",
                "webUrl": "https://x/share/abc",
                "preventsDownload": False,
            },
            "expirationDateTime": None,
            "hasPassword": False,
        },
    )
    out = share_create(FILE_URL)
    assert out["type"] == "view"
    assert out["scope"] == "organization"
    assert out["web_url"] == "https://x/share/abc"
    body = route.calls.last.request.read().decode()
    assert '"type":"view"' in body or '"type": "view"' in body
    assert '"scope":"organization"' in body or '"scope": "organization"' in body
    # Optional fields must not appear when not set
    assert "expirationDateTime" not in body
    assert "password" not in body


@respx.mock
def test_share_create_anonymous_edit_passes_through(store_with_fresh_token: None) -> None:
    """The risky combo (anonymous + edit) is allowed but explicit, so the
    request body includes both — the agent had to opt in."""
    del store_with_fresh_token
    _mock_resolution()
    route = respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createLink").respond(
        json={
            "id": "p2",
            "roles": ["write"],
            "link": {
                "type": "edit",
                "scope": "anonymous",
                "webUrl": "https://x/share/anon-edit",
            },
        },
    )
    out = share_create(FILE_URL, type="edit", scope="anonymous")
    body = route.calls.last.request.read().decode()
    assert "edit" in body
    assert "anonymous" in body
    assert out["type"] == "edit"
    assert out["scope"] == "anonymous"


@respx.mock
def test_share_create_includes_expires_and_password_when_set(
    store_with_fresh_token: None,
) -> None:
    del store_with_fresh_token
    _mock_resolution()
    route = respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createLink").respond(
        json={
            "id": "p3",
            "roles": ["read"],
            "link": {
                "type": "view",
                "scope": "anonymous",
                "webUrl": "https://x/share/locked",
            },
            "expirationDateTime": "2026-12-31T23:59:59Z",
            "hasPassword": True,
        },
    )
    out = share_create(
        FILE_URL,
        type="view",
        scope="anonymous",
        expires="2026-12-31T23:59:59Z",
        password="hunter2",
    )
    body = route.calls.last.request.read().decode()
    assert "2026-12-31" in body
    assert "hunter2" in body
    assert out["expiration_date_time"] == "2026-12-31T23:59:59Z"
    assert out["has_password"] is True


# ---------------------------------------------------------------------
# share_create — validation
# ---------------------------------------------------------------------


def test_share_create_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        share_create("")


@pytest.mark.parametrize(
    "type_arg",
    ["VIEW", "delete", "open", "", " edit", "anything"],
)
def test_share_create_rejects_unknown_type(type_arg: str) -> None:
    with pytest.raises(ValueError, match="type must be one of"):
        share_create(FILE_URL, type=type_arg)


@pytest.mark.parametrize(
    "scope_arg",
    ["public", "private", "PUBLIC", "", "Organization"],
)
def test_share_create_rejects_unknown_scope(scope_arg: str) -> None:
    with pytest.raises(ValueError, match="scope must be one of"):
        share_create(FILE_URL, scope=scope_arg)


def test_share_create_rejects_site_url_without_item() -> None:
    """sp_share_create requires a file/folder URL, not a site URL."""
    with pytest.raises(ValueError, match="file/folder URL"):
        share_create(SITE_URL)


# ---------------------------------------------------------------------
# share_create — error propagation
# ---------------------------------------------------------------------


@respx.mock
def test_share_create_propagates_403(store_with_fresh_token: None) -> None:
    """Tenant has anonymous-sharing disabled -> Graph returns 403."""
    del store_with_fresh_token
    _mock_resolution()
    respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createLink").respond(
        403, json={"error": {"code": "accessDenied"}}
    )
    with pytest.raises(httpx.HTTPStatusError):
        share_create(FILE_URL, scope="anonymous")


@respx.mock
def test_share_create_propagates_404(store_with_fresh_token: None) -> None:
    """File doesn't exist -> driveItem lookup 404s before createLink."""
    del store_with_fresh_token
    respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(json={"id": SITE_ID})
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/policy.docx").respond(404)
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drives").respond(json={"value": []})
    with pytest.raises(httpx.HTTPStatusError):
        share_create(FILE_URL)


# ---------------------------------------------------------------------
# share_revoke
# ---------------------------------------------------------------------


@respx.mock
def test_share_revoke_deletes_permission(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_resolution()
    route = respx.delete(
        f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/permissions/{LINK_ID}"
    ).respond(204)
    share_revoke(FILE_URL, LINK_ID)
    assert route.call_count == 1


@respx.mock
def test_share_revoke_propagates_404_for_already_gone_link(
    store_with_fresh_token: None,
) -> None:
    """Re-revoking is not idempotent at the Graph level — propagates 404."""
    del store_with_fresh_token
    _mock_resolution()
    respx.delete(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/permissions/{LINK_ID}").respond(
        404, json={"error": {"code": "itemNotFound"}}
    )
    with pytest.raises(httpx.HTTPStatusError):
        share_revoke(FILE_URL, LINK_ID)


def test_share_revoke_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        share_revoke("", LINK_ID)


def test_share_revoke_rejects_empty_link_id() -> None:
    with pytest.raises(ValueError, match="non-empty link_id"):
        share_revoke(FILE_URL, "")


def test_share_revoke_rejects_blank_link_id() -> None:
    with pytest.raises(ValueError, match="non-empty link_id"):
        share_revoke(FILE_URL, "   ")


def test_share_revoke_rejects_site_url_without_item() -> None:
    with pytest.raises(ValueError, match="file/folder URL"):
        share_revoke(SITE_URL, LINK_ID)


# ---------------------------------------------------------------------
# share_list
# ---------------------------------------------------------------------


@respx.mock
def test_share_list_filters_to_link_grantees_only(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_resolution()
    respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/permissions").respond(
        json={
            "value": [
                {
                    "id": "p-user",
                    "roles": ["read"],
                    "grantedToV2": {"user": {"displayName": "Alice"}},
                },
                {
                    "id": "p-link",
                    "roles": ["read"],
                    "link": {
                        "type": "view",
                        "scope": "organization",
                        "webUrl": "https://x/share/foo",
                    },
                },
                {
                    "id": "p-link-2",
                    "roles": ["write"],
                    "link": {
                        "type": "edit",
                        "scope": "anonymous",
                        "webUrl": "https://x/share/anon",
                    },
                },
            ],
        },
    )
    out = share_list(FILE_URL)
    assert len(out) == 2  # user grant excluded
    [view, edit] = out
    assert view["id"] == "p-link"
    assert view["type"] == "view"
    assert view["web_url"] == "https://x/share/foo"
    assert edit["scope"] == "anonymous"


@respx.mock
def test_share_list_returns_empty_when_no_links(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_resolution()
    respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/permissions").respond(
        json={
            "value": [
                {
                    "id": "p-user-only",
                    "roles": ["read"],
                    "grantedToV2": {"user": {"displayName": "A"}},
                }
            ],
        },
    )
    assert share_list(FILE_URL) == []


def test_share_list_rejects_site_url() -> None:
    with pytest.raises(ValueError, match="file/folder URL"):
        share_list(SITE_URL)


def test_share_list_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        share_list("")


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------


def test_valid_link_types_match_microsoft_documented_set() -> None:
    """Pin the type set so accidental loosening triggers a test failure."""
    assert VALID_LINK_TYPES == {"view", "edit", "embed", "blocksDownload"}


def test_valid_link_scopes_match_microsoft_documented_set() -> None:
    assert VALID_LINK_SCOPES == {"anonymous", "organization", "users"}


def test_normalise_create_response_handles_missing_link_facet() -> None:
    """Defensive: Graph response without `link` key shouldn't crash."""
    out = _normalise_create_response({"id": "p", "roles": ["read"]})
    assert out["id"] == "p"
    assert out["web_url"] == ""
    assert out["type"] == ""


def test_normalise_create_response_handles_non_dict_link() -> None:
    """Pure defensiveness — Graph wouldn't actually return this, but a
    test catches accidental destructuring on bad data."""
    out = _normalise_create_response({"id": "p", "link": "not-a-dict"})
    assert out["web_url"] == ""


def test_normalise_create_response_expiration_None_translates_to_None() -> None:
    out = _normalise_create_response({"id": "p", "expirationDateTime": None})
    assert out["expiration_date_time"] is None


def test_normalise_create_response_handles_non_list_roles() -> None:
    out = _normalise_create_response({"id": "p", "roles": "not-a-list"})
    assert out["roles"] == []


def test_normalise_create_response_carries_prevents_download() -> None:
    out = _normalise_create_response(
        {
            "id": "p",
            "link": {
                "type": "view",
                "scope": "organization",
                "webUrl": "https://x",
                "preventsDownload": True,
            },
        }
    )
    assert out["prevents_download"] is True


def test_normalise_existing_link_pulls_web_url_from_grantee() -> None:
    """sp_share_list relies on sp_permissions including link_web_url in
    its grantee normalisation."""
    out = _normalise_existing_link(
        {
            "id": "p-link",
            "roles": ["read"],
            "grantee": {
                "type": "link",
                "link_type": "view",
                "link_scope": "organization",
                "link_web_url": "https://x/share/abc",
            },
        }
    )
    assert out["web_url"] == "https://x/share/abc"
    assert out["type"] == "view"


def test_normalise_existing_link_handles_missing_grantee() -> None:
    out = _normalise_existing_link({"id": "p", "roles": ["read"]})
    assert out["web_url"] == ""
    assert out["type"] == ""


def test_normalise_existing_link_handles_non_dict_grantee() -> None:
    out = _normalise_existing_link({"id": "p", "grantee": "not-a-dict"})
    assert out["web_url"] == ""


# ---------------------------------------------------------------------
# Server-layer integration: defaults flow through correctly
# ---------------------------------------------------------------------


@respx.mock
def test_share_create_with_only_valid_types(store_with_fresh_token: None) -> None:
    """Smoke-test all four documented type values — any one regression is caught."""
    del store_with_fresh_token
    _mock_resolution()
    respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createLink").respond(
        json={"id": "p", "link": {"type": "view", "scope": "organization", "webUrl": "x"}},
    )
    for t in VALID_LINK_TYPES:
        share_create(FILE_URL, type=t)  # must not raise


@respx.mock
def test_share_create_does_not_send_password_when_empty_string(
    store_with_fresh_token: None,
) -> None:
    """Empty password should NOT appear in the body — Graph rejects empty values."""
    del store_with_fresh_token
    _mock_resolution()
    route = respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createLink").respond(
        json={"id": "p", "link": {"type": "view", "scope": "organization", "webUrl": "x"}},
    )
    share_create(FILE_URL, password="")
    body = route.calls.last.request.read().decode()
    assert "password" not in body


@respx.mock
def test_share_create_does_not_send_expires_when_empty_string(
    store_with_fresh_token: None,
) -> None:
    del store_with_fresh_token
    _mock_resolution()
    route = respx.post(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/createLink").respond(
        json={"id": "p", "link": {"type": "view", "scope": "organization", "webUrl": "x"}},
    )
    share_create(FILE_URL, expires="")
    body = route.calls.last.request.read().decode()
    assert "expirationDateTime" not in body
