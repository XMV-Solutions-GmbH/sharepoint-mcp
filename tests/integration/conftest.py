# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Shared fixtures for integration tests.

These tests use boundary mocks at the HTTP layer (`respx` against
graph.microsoft.com) and at the keyring/token-store layer (in-memory
fake), but exercise the actual Python modules between those
boundaries. So they catch wiring bugs that pure unit tests can miss.
"""

from __future__ import annotations

import time
from collections.abc import Iterator
from pathlib import Path

import pytest

from sharepoint_mcp.auth.tokens import CachedToken


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
def fake_token_store(monkeypatch: pytest.MonkeyPatch) -> Iterator[None]:
    """Patch get_token_store to return a store with a fresh token.

    Avoids touching real keyring or filesystem.
    """
    cached = CachedToken(
        access_token="AT-integration",
        refresh_token="RT-integration",
        expires_at=time.time() + 3600,
        scope="Files.ReadWrite.All Sites.ReadWrite.All",
    )
    monkeypatch.setattr(
        "sharepoint_mcp.auth.get_token_store",
        lambda: _MemStore(cached.to_json().encode()),
    )
    yield


@pytest.fixture
def isolated_registry_dir(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """Redirect the checkout-registry default dir to a tmp_path."""
    monkeypatch.setattr(
        "sharepoint_mcp.checkout_registry.DEFAULT_REGISTRY_DIR",
        tmp_path,
    )
    return tmp_path
