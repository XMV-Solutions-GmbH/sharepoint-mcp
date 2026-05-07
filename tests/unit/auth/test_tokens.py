# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for the CachedToken value type."""

from __future__ import annotations

import time

from sharepoint_mcp.auth.tokens import CachedToken


def test_serialization_roundtrip() -> None:
    original = CachedToken(
        access_token="AT-XYZ",
        refresh_token="RT-ABC",
        expires_at=1_900_000_000.0,
        scope="Files.ReadWrite.All offline_access",
    )
    recovered = CachedToken.from_json(original.to_json())
    assert recovered == original


def test_serialization_with_no_refresh_token() -> None:
    original = CachedToken(
        access_token="AT",
        refresh_token=None,
        expires_at=1_900_000_000.0,
        scope="User.Read",
    )
    recovered = CachedToken.from_json(original.to_json())
    assert recovered == original
    assert recovered.refresh_token is None


def test_is_expired_far_future() -> None:
    token = CachedToken(
        access_token="AT",
        refresh_token="RT",
        expires_at=time.time() + 3600,
        scope="",
    )
    assert token.is_expired() is False


def test_is_expired_now() -> None:
    token = CachedToken(
        access_token="AT",
        refresh_token="RT",
        expires_at=time.time() - 1,
        scope="",
    )
    assert token.is_expired() is True


def test_is_expired_within_default_buffer() -> None:
    """Token that expires in 30s should be considered expired (60s default buffer)."""
    token = CachedToken(
        access_token="AT",
        refresh_token="RT",
        expires_at=time.time() + 30,
        scope="",
    )
    assert token.is_expired() is True


def test_is_expired_with_custom_zero_buffer() -> None:
    token = CachedToken(
        access_token="AT",
        refresh_token="RT",
        expires_at=time.time() + 30,
        scope="",
    )
    assert token.is_expired(buffer=0) is False
