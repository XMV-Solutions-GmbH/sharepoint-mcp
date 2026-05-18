# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for sp_search_files.

Mocks both Microsoft Graph (via respx) and the auth layer (by giving
search() an in-memory TokenStore via auth.get_token's keyword args).
No real network, no real keyring.
"""

from __future__ import annotations

import json
import time
from collections.abc import Iterator

import httpx
import pytest
import respx

from sharepoint_mcp.auth.tokens import CachedToken
from sharepoint_mcp.tools.search import (
    GRAPH_SEARCH_URL,
    _extract_hits,
    _extract_path,
    _extract_user,
    search,
)

# ---------------------------------------------------------------------
# Token-store fixture: a minimal in-memory TokenStore that returns a
# fresh CachedToken so search() can resolve get_token() without
# touching the real keyring or hitting Microsoft Identity.
# ---------------------------------------------------------------------


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
    """Patch get_token_store to return a store holding a fresh token.

    Used by tests that don't pass `store=` explicitly.
    """
    cached = CachedToken(
        access_token="AT-test",
        refresh_token="RT-test",
        expires_at=time.time() + 3600,
        scope="",
    )
    fake = _MemStore(cached.to_json().encode())
    monkeypatch.setattr("sharepoint_mcp.auth.get_token_store", lambda: fake)
    yield


# ---------------------------------------------------------------------
# search() — happy path + filter expansion + bearer-header
# ---------------------------------------------------------------------


@respx.mock
def test_search_returns_parsed_hits(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    respx.post(GRAPH_SEARCH_URL).respond(
        json={
            "value": [
                {
                    "hitsContainers": [
                        {
                            "total": 1,
                            "hits": [
                                {
                                    "resource": {
                                        "@odata.type": "#microsoft.graph.driveItem",
                                        "name": "policy.docx",
                                        "webUrl": "https://contoso.sharepoint.com/sites/foo/policy.docx",
                                        "lastModifiedDateTime": "2026-04-15T10:00:00Z",
                                        "lastModifiedBy": {
                                            "user": {"displayName": "Alice", "email": "a@x.de"}
                                        },
                                        "size": 1234,
                                        "parentReference": {
                                            "path": "/sites/foo/root:/Shared Documents/policies"
                                        },
                                    },
                                },
                            ],
                        },
                    ],
                },
            ],
        }
    )
    results = search("policy")
    assert len(results) == 1
    hit = results[0]
    assert hit["name"] == "policy.docx"
    assert hit["author"] == "Alice"
    assert hit["size"] == 1234
    assert hit["last_modified"] == "2026-04-15T10:00:00Z"
    assert hit["web_url"].endswith("policy.docx")
    assert hit["path"] == "/Shared Documents/policies/policy.docx"


@respx.mock
def test_search_empty_results(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    respx.post(GRAPH_SEARCH_URL).respond(
        json={"value": [{"hitsContainers": [{"total": 0, "hits": []}]}]},
    )
    assert search("no-such-thing") == []


@respx.mock
def test_search_sends_bearer_token(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    route = respx.post(GRAPH_SEARCH_URL).respond(json={"value": []})
    search("anything")
    auth_header = route.calls.last.request.headers.get("authorization")
    assert auth_header == "Bearer AT-test"


@respx.mock
def test_search_filters_translate_to_kql(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    route = respx.post(GRAPH_SEARCH_URL).respond(json={"value": []})
    search(
        "policy",
        site="https://contoso.sharepoint.com/sites/foo",
        folder="/Shared Documents/policies",
        file_type="docx",
        modified_after="2024-01-01",
        limit=10,
    )
    body = json.loads(route.calls.last.request.read())
    request0 = body["requests"][0]
    qs = request0["query"]["queryString"]
    # All filters appear in the queryString as proper KQL
    assert qs.startswith("policy ")
    assert 'site:"https://contoso.sharepoint.com/sites/foo"' in qs
    assert 'path:"/Shared Documents/policies"' in qs
    assert "fileExtension:docx" in qs
    assert "lastModifiedDateTime>=2024-01-01" in qs
    # And the limit is honoured
    assert request0["size"] == 10


@respx.mock
def test_search_propagates_http_errors(store_with_fresh_token: None) -> None:
    del store_with_fresh_token
    respx.post(GRAPH_SEARCH_URL).respond(
        401, json={"error": {"code": "InvalidAuthenticationToken"}}
    )
    with pytest.raises(httpx.HTTPStatusError):
        search("anything")


# ---------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------


def test_search_rejects_empty_query() -> None:
    with pytest.raises(ValueError, match="non-empty query"):
        search("")


def test_search_rejects_blank_query() -> None:
    with pytest.raises(ValueError, match="non-empty query"):
        search("   ")


def test_search_rejects_non_positive_limit() -> None:
    with pytest.raises(ValueError, match="limit must be positive"):
        search("query", limit=0)


# ---------------------------------------------------------------------
# Helper extraction logic — unit-level
# ---------------------------------------------------------------------


def test_extract_hits_handles_missing_value() -> None:
    assert _extract_hits({}) == []


def test_extract_hits_handles_empty_value() -> None:
    assert _extract_hits({"value": []}) == []


def test_extract_hits_handles_no_hits_container() -> None:
    assert _extract_hits({"value": [{"hitsContainers": []}]}) == []


def test_extract_user_with_displayName() -> None:
    assert _extract_user({"user": {"displayName": "Alice"}}) == "Alice"


def test_extract_user_falls_back_to_email() -> None:
    assert _extract_user({"user": {"email": "a@x.de"}}) == "a@x.de"


def test_extract_user_returns_none_when_missing() -> None:
    assert _extract_user(None) is None
    assert _extract_user({}) is None
    assert _extract_user({"user": None}) is None


def test_extract_path_strips_root_prefix() -> None:
    resource = {
        "name": "file.docx",
        "parentReference": {"path": "/sites/foo/root:/Shared Documents/folder"},
    }
    assert _extract_path(resource) == "/Shared Documents/folder/file.docx"


def test_extract_path_no_parent_ref() -> None:
    assert _extract_path({"name": "file.docx"}) == "file.docx"


def test_extract_path_empty_path() -> None:
    assert _extract_path({"name": "file.docx", "parentReference": {}}) == "file.docx"
