# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""End-to-end harness smoke test — the gate test from § 5.

Validates that our auth + Graph stack actually works against the real
Microsoft Identity + Microsoft Graph endpoints (NOT mocked). This is
the test that proves our model of the API matches reality.

Prerequisites:

- A `harness` profile token cache populated by running
  `uv run mcp-server-sharepoint login --profile harness` once (see
  docs/testconcept.md once that lands).
- The dedicated harness SharePoint site provisioned per issue #14
  (M365 group `sharepoint-mcp-harness`, test user `d.koller@xmv.de`
  added as member).

When no harness credentials are present (e.g., a unit/integration
CI run without the harness secret), the tests are SKIPPED rather
than failing — so an absent harness setup doesn't break the lower
layers' CI.
"""

from __future__ import annotations

import httpx
import pytest

from sharepoint_mcp.auth import AuthRequiredError, get_token

HARNESS_PROFILE = "harness"
EXPECTED_TENANT_DOMAIN = "@xmv.de"
EXPECTED_HARNESS_SITE_NAME = "sharepoint-mcp-harness"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def _harness_token_or_skip() -> str:
    """Return cached harness access token; skip the test if unavailable."""
    try:
        return get_token(HARNESS_PROFILE)
    except AuthRequiredError as exc:
        pytest.skip(
            f"Harness credentials not available: {exc}. "
            "Run `uv run mcp-server-sharepoint login --profile harness` to populate.",
        )


def test_harness_token_authenticates_to_graph_me() -> None:
    """Cached token works end-to-end against Microsoft Graph /me.

    Proves:
    - Token store (keyring or plain file) holds a valid CachedToken.
    - get_token() returns the access token (silent path, no refresh
      this run since we just logged in).
    - The access token is in the right format for Graph
      Authorization-Bearer headers.
    - Microsoft Graph accepts our token and returns the harness
      user's profile.
    """
    token = _harness_token_or_skip()

    with httpx.Client(timeout=10.0) as client:
        response = client.get(
            f"{GRAPH_BASE}/me",
            headers={"Authorization": f"Bearer {token}"},
        )
    response.raise_for_status()
    me = response.json()

    upn = me.get("userPrincipalName", "")
    assert upn.endswith(EXPECTED_TENANT_DOMAIN), (
        f"Harness user must be in {EXPECTED_TENANT_DOMAIN} tenant; got {upn!r}. "
        "Did you sign in as the wrong account?"
    )
    assert me.get("id"), "Microsoft Graph /me must return an id"


def test_harness_site_is_visible() -> None:
    """The dedicated harness SharePoint site is discoverable for the test user.

    Confirms the M365 group from #14 was provisioned and the test
    user was added with at least Read access on the site. A failure
    here usually means the test user wasn't added as a member of
    the group.
    """
    token = _harness_token_or_skip()

    with httpx.Client(timeout=10.0) as client:
        response = client.get(
            f"{GRAPH_BASE}/sites",
            params={"search": EXPECTED_HARNESS_SITE_NAME},
            headers={"Authorization": f"Bearer {token}"},
        )
    response.raise_for_status()
    sites = response.json().get("value", [])

    matching = [s for s in sites if s.get("displayName") == EXPECTED_HARNESS_SITE_NAME]
    assert matching, (
        f"Harness site '{EXPECTED_HARNESS_SITE_NAME}' not found among "
        f"{[s.get('displayName') for s in sites]!r}. "
        "Did you provision the M365 group and add the test user?"
    )

    site = matching[0]
    web_url = site.get("webUrl", "")
    assert "sharepoint-mcp-harness" in web_url, f"Harness site URL looks wrong: {web_url!r}"
