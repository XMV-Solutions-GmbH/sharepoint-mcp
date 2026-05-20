# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_share_permission_list (#46)."""

from __future__ import annotations

import time
from collections.abc import Iterator

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools.permissions import (
    _extract_grantee,
    _extract_permissions,
    _normalise_identity,
    _one_permission,
    permissions,
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


def _mock_site_lookup() -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_HOST}:{SITE_PATH}").respond(
        json={"id": SITE_ID},
    )


def _mock_item_lookup() -> respx.Route:
    return respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/drive/root:/policy.docx").respond(
        json={"id": ITEM_ID, "parentReference": {"driveId": DRIVE_ID}},
    )


# ---------------------------------------------------------------------
# Site-level permissions
# ---------------------------------------------------------------------


@respx.mock
def test_permissions_site_url_hits_site_endpoint(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    route = respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/permissions").respond(
        json={
            "value": [
                {
                    "id": "p1",
                    "roles": ["read"],
                    "grantedToV2": {
                        "user": {"displayName": "Alice", "email": "a@x"},
                    },
                }
            ],
        },
    )
    [out] = permissions(SITE_URL)
    assert out["id"] == "p1"
    assert out["roles"] == ["read"]
    assert out["grantee"]["type"] == "user"
    assert out["grantee"]["display_name"] == "Alice"
    assert route.called


# ---------------------------------------------------------------------
# Item-level permissions
# ---------------------------------------------------------------------


@respx.mock
def test_permissions_item_url_hits_drive_endpoint(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    _mock_item_lookup()
    respx.get(f"{GRAPH_BASE}/drives/{DRIVE_ID}/items/{ITEM_ID}/permissions").respond(
        json={
            "value": [
                {
                    "id": "p2",
                    "roles": ["write"],
                    "grantedToIdentitiesV2": [
                        {"user": {"displayName": "Bob"}},
                    ],
                }
            ],
        },
    )
    [out] = permissions(FILE_URL)
    assert out["roles"] == ["write"]
    assert out["grantee"]["display_name"] == "Bob"


@respx.mock
def test_permissions_normalises_sharing_link(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/permissions").respond(
        json={
            "value": [
                {
                    "id": "link-1",
                    "roles": ["read"],
                    "link": {"type": "view", "scope": "anonymous"},
                }
            ],
        },
    )
    [out] = permissions(SITE_URL)
    assert out["grantee"]["type"] == "link"
    assert out["grantee"]["link_type"] == "view"
    assert out["grantee"]["link_scope"] == "anonymous"
    assert out["grantee"]["display_name"] == ""


@respx.mock
def test_permissions_marks_inherited(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/permissions").respond(
        json={
            "value": [
                {
                    "id": "p3",
                    "roles": ["owner"],
                    "grantedToV2": {"user": {"displayName": "Owner"}},
                    "inheritedFrom": {"driveId": "X", "id": "Y"},
                },
                {
                    "id": "p4",
                    "roles": ["read"],
                    "grantedToV2": {"user": {"displayName": "Direct"}},
                },
            ],
        },
    )
    [inherited, direct] = permissions(SITE_URL)
    assert inherited["inherited"] is True
    assert direct["inherited"] is False


@respx.mock
def test_permissions_empty(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/permissions").respond(json={"value": []})
    assert permissions(SITE_URL) == []


@respx.mock
def test_permissions_propagates_403(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    _mock_site_lookup()
    respx.get(f"{GRAPH_BASE}/sites/{SITE_ID}/permissions").respond(
        403, json={"error": {"code": "accessDenied"}}
    )
    with pytest.raises(httpx.HTTPStatusError):
        permissions(SITE_URL)


def test_permissions_rejects_empty_url() -> None:
    with pytest.raises(ValueError, match="non-empty url"):
        permissions("")


# ---------------------------------------------------------------------
# Helpers — grantee normalisation
# ---------------------------------------------------------------------


def test_extract_grantee_prefers_user_over_group() -> None:
    """When both user and group facets are populated (rare), user wins."""
    out = _extract_grantee(
        {"grantedToV2": {"user": {"displayName": "Alice"}, "group": {"displayName": "Sales"}}}
    )
    assert out["type"] == "user"
    assert out["display_name"] == "Alice"


def test_extract_grantee_falls_back_to_grantedToIdentitiesV2() -> None:
    out = _extract_grantee({"grantedToIdentitiesV2": [{"group": {"displayName": "Eng"}}]})
    assert out["type"] == "group"
    assert out["display_name"] == "Eng"


def test_extract_grantee_handles_application() -> None:
    out = _extract_grantee({"grantedToV2": {"application": {"displayName": "MyApp"}}})
    assert out["type"] == "application"
    assert out["display_name"] == "MyApp"


def test_extract_grantee_empty_returns_unknown() -> None:
    out = _extract_grantee({})
    assert out["type"] == "unknown"
    assert out["display_name"] == ""


def test_normalise_identity_uses_loginName_fallback() -> None:
    """siteUser sometimes only has loginName, not displayName."""
    out = _normalise_identity({"siteUser": {"loginName": "i:0#.f|membership|x@y"}})
    assert out["type"] == "siteUser"
    assert out["display_name"] == "i:0#.f|membership|x@y"


def test_normalise_identity_unknown_shape() -> None:
    out = _normalise_identity({"weirdField": {"foo": "bar"}})
    assert out["type"] == "unknown"


def test_extract_permissions_handles_missing_value() -> None:
    assert _extract_permissions({}) == []


def test_one_permission_handles_non_list_roles() -> None:
    """Defensive: shouldn't crash on a malformed roles field."""
    out = _one_permission({"id": "x", "roles": "not-a-list"})
    assert out["roles"] == []
