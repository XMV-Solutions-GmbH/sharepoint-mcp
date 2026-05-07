<!--
SPDX-License-Identifier: MIT OR Apache-2.0
SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
SPDX-FileContributor: David Koller <david.koller@xmv.de>
-->

# Spike: msgraph-sdk-python vs raw httpx

**Date**: 2026-05-06
**Issue**: [#8](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues/8)
**Decision**: **raw `httpx`**, not `msgraph-sdk-python`.

---

## Question

`mcp-server-sharepoint` touches ~6 Microsoft Graph endpoints in v0.1: search, drive children, item content, checkout, checkin, discardCheckout (plus `/me` for sign-in verification). Both `msgraph-sdk-python` (the official Kiota-generated SDK) and a hand-written `httpx` client could implement these. Pick one.

## Method

1. Measure transitive-dependency footprint of each.
2. Implement the same operation (`sp_list`) in both styles.
3. Check how each integrates with our keyring-backed token cache (we already own the auth flow; the Graph client is a downstream consumer, not the auth orchestrator).

## Footprint

Measured by installing each package fresh into an empty venv and inspecting `site-packages/`.

| Package | Disk | Top-level packages | Notes |
|---|---|---|---|
| `msgraph-sdk` | **220 MB** | 47 | `msgraph` module alone is 182 MB of generated Kiota code; pulls `cryptography` (15 MB), `aiohttp` (6.3 MB), `azure-*`, `opentelemetry`, `multidict`, `frozenlist`, `propcache`, `yarl`, `urllib3`, `zipp` |
| `httpx` | **2.4 MB** | 10 | direct deps: `httpcore`, `h11`, `anyio`, `sniffio`, `certifi`, `idna` |

~**92× difference in disk footprint**. For a tool installed via `uvx mcp-server-sharepoint` and intended to be lightweight, this matters: every transitive dep is also attack surface, install time, and a potential source of supply-chain risk.

## Code comparison: `sp_list`

### Option A — `msgraph-sdk-python`

```python
from azure.core.credentials import AccessToken, TokenCredential
from msgraph import GraphServiceClient

class KeyringTokenCredential(TokenCredential):
    """Adapter: keyring-cached token → azure-identity's TokenCredential."""

    def __init__(self, profile: str) -> None:
        self.profile = profile

    def get_token(self, *scopes: str, **kwargs: object) -> AccessToken:
        cached = get_cached_access_token(self.profile)
        return AccessToken(token=cached.value, expires_on=cached.expires_at)


async def sp_list(profile: str, drive_id: str, item_id: str) -> list[dict]:
    client = GraphServiceClient(
        credentials=KeyringTokenCredential(profile),
        scopes=["Files.ReadWrite.All"],
    )
    page = await (
        client.drives.by_drive_id(drive_id)
        .items.by_drive_item_id(item_id)
        .children.get()
    )
    if page is None or page.value is None:
        return []
    return [
        {
            "name": item.name,
            "type": "folder" if item.folder else "file",
            "size": item.size,
            "modified": (
                item.last_modified_date_time.isoformat()
                if item.last_modified_date_time
                else None
            ),
            "web_url": item.web_url,
        }
        for item in page.value
    ]
```

### Option B — raw `httpx`

```python
import httpx

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


async def sp_list(profile: str, drive_id: str, item_id: str) -> list[dict]:
    token = get_cached_access_token(profile)
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/children"
    async with httpx.AsyncClient() as client:
        response = await client.get(url, headers={"Authorization": f"Bearer {token.value}"})
        response.raise_for_status()
        page = response.json()
    return [
        {
            "name": item["name"],
            "type": "folder" if "folder" in item else "file",
            "size": item.get("size"),
            "modified": item.get("lastModifiedDateTime"),
            "web_url": item.get("webUrl"),
        }
        for item in page.get("value", [])
    ]
```

Both are about the same length once the SDK option includes the `TokenCredential` adapter — but the httpx version has no abstraction layer between us and the wire. When something goes wrong against a real tenant, the failure is at the HTTP level we already understand, not three layers deep into a generated client.

## Auth-token integration

Our auth flow caches refresh + access tokens in the OS keyring under `SP_PROFILE`-namespaced keys. The Graph client is **downstream** of this cache; it does not own auth.

- **`msgraph-sdk`** assumes `azure.identity` `TokenCredential`-shaped sources. We can adapt (see the `KeyringTokenCredential` above), but the SDK's implicit assumptions about token caching, refresh, and scope acquisition fight ours. We end up implementing a no-op adapter that calls our cache anyway.
- **`httpx`** takes the bearer token as a header. No adapter, no implicit behaviour, no surprise about which code path refreshed the token last.

Conclusion: the SDK was designed for apps where Microsoft's auth library *is* the auth source. For an OSS tool that needs to control token persistence, refresh timing, and per-profile separation explicitly, the SDK adds friction without value.

## Other dimensions considered

- **Type safety**: SDK gives typed Graph response models. We immediately project everything to `dict`-shaped MCP tool results, so the typing is lost at the boundary. mypy's view of the SDK return types is also unreliable because Kiota generates `Optional[X] | None`-heavy signatures that we'd have to assert/cast through anyway.
- **API drift handling**: SDK auto-updates if Microsoft changes Graph. Counterpoint: the 6 endpoints we touch are stable and well-documented; raw HTTP against a stable API surface is a low-cost commitment, and any drift would be caught immediately by harness tests.
- **Search query**: `POST /search/query` has a JSON request body shape we'd hand-roll either way.
- **Chunked upload**: Both options support resumable upload sessions; httpx requires more code, SDK has a helper. Plausible reason to revisit if `sp_save` with large files turns out to need lots of upload-session orchestration. For v0.1, simple uploads are fine; chunked uploads are tracked separately as spike #24.

## Decision

**Use raw `httpx` directly. Do not depend on `msgraph-sdk-python`.**

Reasons in priority order:

1. **Footprint**: 220 MB → 2.4 MB. Smaller install, smaller attack surface, faster `uvx` cold-start.
2. **Control over auth**: we already own the token cache; the SDK fights this contract.
3. **Debuggability**: when an audit-critical save fails, we want the failure to be readable at the HTTP level, not buried in a generated client.
4. **Low endpoint count**: 6 endpoints don't justify a generic SDK.

## When the SDK would be the right call

If a future project needs to hit dozens of Graph endpoints with non-uniform shapes, wants Microsoft-supplied response typing, and is happy to defer auth to `azure.identity` end-to-end — `msgraph-sdk-python` is the correct choice. That is not us.

## Follow-ups landed by this spike

- `httpx` added as a runtime dependency in `pyproject.toml`.
- `docs/app-concept.md` open-question #1 closed.
- The `sp_list` httpx sketch above will land for real in #17.
- An end-to-end harness validation of the chosen client (real Graph call from `tests/harness/`) lands as part of #28 (harness gate).
