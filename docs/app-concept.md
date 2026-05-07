<!-- SPDX-License-Identifier: MIT OR Apache-2.0 -->

# mcp-server-sharepoint — App Concept

A Model Context Protocol server that lets AI coding agents read and edit files in SharePoint document libraries **without breaking SharePoint's version history, audit trail, or locking semantics**.

Built because the standard alternatives — rclone, WebDAV mounts, the Anthropic-hosted M365 MCP — either skip SharePoint's checkout/checkin model entirely or expose it only for search and read. Neither pattern is acceptable when your documents include ISMS records, controlled procedures, or anything that needs an audit trail.

---

## Why this exists

For organisations that run their controlled documentation in SharePoint (ISO 27001, DSGVO, due-diligence prep, contract templates), the canonical document is in SharePoint with version history, retention policies, and access control. An AI agent that "edits the file" by overwriting it with a new copy:

- creates a new version, but with no commit message,
- holds no lock, so concurrent edits race silently,
- attributes the change to whatever client did the upload (often "rclone client"), not to the human,
- breaks if SharePoint requires explicit checkout for the library.

The Microsoft Graph API exposes a clean **checkout / edit / checkin** model. This MCP wraps that model and presents it as a small set of MCP tools an agent can call.

---

## Core use cases

1. **Read-only research** — agent searches SharePoint, reads docs, cites paths, never mutates anything.
2. **Draft + promote** — agent drafts a doc locally (in the Claude Code repo's working directory), then publishes it to SharePoint as a new file.
3. **Edit existing doc with audit trail** — agent checks out a doc, edits it, checks it back in with a commit message and major/minor version selection. Lock prevents concurrent clobbering.
4. **Discard speculative edit** — checkout something, decide not to keep changes, release the lock without committing.

---

## Non-goals

- **Not a sync engine.** No bidirectional continuous sync à la rclone bisync. All edits go through explicit `open` / `save` / `release` calls.
- **Not a SharePoint admin tool.** No library creation, permission management, retention-policy editing.
- **Not a search index.** Search delegates to Microsoft Graph; no local indexing or caching of search results.
- **Not multi-protocol.** OneDrive personal works incidentally (same Graph endpoints), but Outlook / Teams / Calendar are out of scope. If those are needed, sibling MCPs.

---

## Tools exposed (MCP surface)

```text
sp_search(query, site?, folder?, file_type?, modified_after?)
    → list of (path, web_url, last_modified, author)

sp_list(path)
    → folder listing with type/size/modified for each child

sp_read(path)
    → downloads to a temp location, returns local path; no checkout, read-only

sp_open(path)
    → checkout + download to working directory; returns local path
    → fails if file is already checked out by someone else

sp_save(path, comment, version="minor"|"major")
    → upload + checkin with comment; returns new versionId
    → ETag-check; fails if file changed under us

sp_release(path)
    → discardCheckout; drops local working copy

sp_status()
    → list of currently checked-out files (path, since, local_path)

sp_history(path, limit=20)
    → list of versions with id, modified, author, comment

sp_get_version(path, version_id)
    → fetch a specific historical version (read-only)
```

Every tool maps to one or two Microsoft Graph calls. No clever caching beyond what Graph already provides.

---

## Architecture

```text
┌──────────────────┐      stdio JSON-RPC      ┌─────────────────────┐
│   Claude Code    │ ◄──────────────────────► │ mcp-server-sharepoint│
│   (or any        │                          │  (Python process,   │
│    MCP client)   │                          │   one per tenant)   │
└──────────────────┘                          └──────────┬──────────┘
                                                         │ Microsoft Graph
                                                         │ (HTTPS + OAuth)
                                                         ▼
                                              ┌────────────────────┐
                                              │   SharePoint /     │
                                              │   OneDrive (M365)  │
                                              └────────────────────┘
```

**Runtime**: single process per tenant, started via `.mcp.json` config in the consuming repo. Stateless except for token cache and "what's currently checked out" registry.

**Language**: Python. Aligns with the official MCP Python SDK and `msgraph-sdk-python` (or plain `httpx` for fewer dependencies — TBD in tech-spike). Type-checked with mypy, formatted with ruff.

**Working directory**: configurable per-instance (`WORKING_DIR` env var). Default `~/.cache/sharepoint-mcp/<tenant>/working/`. Files live there only while checked out; `sp_release` and `sp_save` clear them.

---

## Authentication

**OAuth 2.0 Device Code Flow** for interactive auth on headless Linux boxes:

1. First run surfaces a user code and verification URL to the calling MCP client (which displays them to the human).
2. User opens URL on phone or laptop, enters code, signs in with their M365 account.
3. Server polls Microsoft Identity until consent is granted; on success, access + refresh token are cached locally via OS keyring (`keyring` Python lib) — no plaintext on disk.
4. Refresh token used silently afterwards; re-prompt on full expiry (60–90 days).

### Default client: XMV-hosted multi-tenant app

The package ships with a baked-in `client_id` for an Entra app registration owned by **XMV Solutions GmbH** — multi-tenant, public client, delegated scopes only (`Files.ReadWrite.All`, `Sites.ReadWrite.All`, `User.Read`, `offline_access`). End users install via `uvx mcp-server-sharepoint` and sign in immediately. No tenant-specific app registration required; no IT-admin involvement on the consumer side.

Registered in XMV's Entra tenant on 2026-05-06:

- **`client_id`** (default `SP_CLIENT_ID`): `cb7cf68d-90d5-4841-90a7-de3a40be280b`
- **Display name**: `mcp-server-sharepoint`
- **Sign-in audience**: `AzureADMultipleOrgs`
- **Homepage**: <https://github.com/XMV-Solutions-GmbH/sharepoint-mcp>
- **Privacy URL**: <https://xmv.de/oss/sharepoint-mcp/privacy>
- **Terms URL**: <https://xmv.de/oss/sharepoint-mcp/terms>

This mirrors the pattern used by Azure CLI, GitHub CLI, and similar OSS tools: the publisher hosts a single multi-tenant public-client app registration; the app ID is not a secret. End users sign in with their own M365 credentials, every action is attributed to the human in SharePoint's audit log.

Tenant routing happens via the `common` / `organizations` endpoint at sign-in — the user's tenant is derived from the account they pick, no `SP_TENANT_ID` configuration needed by default.

### BYO override (enterprise / restrictive tenants)

Tenants with strict app-allowlisting policies can override defaults via env:

- `SP_TENANT_ID=<guid>` — pin sign-in to a specific tenant instead of `common`.
- `SP_CLIENT_ID=<guid>` — use the organization's own app registration instead of the XMV default.

If neither is set, `uvx mcp-server-sharepoint` is zero-config.

### Service-principal auth (deferred)

Client-credentials flow for unattended automation is out of scope for v0.1. MVP is interactive user auth — every action is attributed to the signed-in user, which is what compliance wants.

---

## Multi-profile pattern

For zero-config use, one MCP entry with no env vars at all is enough:

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "uvx",
      "args": ["mcp-server-sharepoint"]
    }
  }
}
```

When working across multiple SharePoint tenants from the same machine (consultancy workflow with several customers), use distinct `SP_PROFILE` values to keep token caches and working directories separated:

```json
{
  "mcpServers": {
    "sharepoint-acme": {
      "command": "uvx",
      "args": ["mcp-server-sharepoint"],
      "env": { "SP_PROFILE": "acme" }
    },
    "sharepoint-globex": {
      "command": "uvx",
      "args": ["mcp-server-sharepoint"],
      "env": { "SP_PROFILE": "globex", "SP_TENANT_ID": "<guid>" }
    }
  }
}
```

Token caches and working directories are namespaced by `SP_PROFILE` (default: `default`). Each profile holds its own refresh token; a second customer means a second `mcpServers` entry. Tools appear in the agent as `mcp__sharepoint-acme__sp_search` etc.

---

## Conflict and safety model

- **Library-level checkout requirement** is honoured: if the library forces explicit checkout, `sp_save` without prior `sp_open` returns an error.
- **Lock is real**: `sp_open` calls Graph `/checkout`, which actually locks the file in SharePoint. Other users see "checked out by [user]" until release.
- **Stale-edit detection**: `sp_save` includes the `If-Match` ETag from `sp_open`; mismatch → error, agent must re-`open` to reconcile.
- **Crash recovery**: `sp_status` is the source of truth for what's checked out. If the MCP crashed mid-edit, `sp_release` (or admin discardCheckout via web) cleans up.
- **No silent overwrites**: there is no path that writes to SharePoint without an explicit `sp_save` call from the agent, with a comment.

---

## Out-of-band concerns

- **Audit attribution**: actions appear in SharePoint audit log as the signed-in user with the `mcp-server-sharepoint` user-agent string. Useful for distinguishing AI-mediated edits from manual ones.
- **Telemetry**: none by default. Opt-in structured logging to stderr; consuming app routes it.
- **Compliance scope**: this MCP does not pretend to be a controlled processing system itself. It is a thin pass-through to Graph; SharePoint's existing controls remain authoritative.

---

## Testability

Per `ENGINEERING_PRINCIPLES.md` § 5, we maintain **three distinct test layers**, and the harness layer is the gate that must be green before any feature ticket enters "Doing".

### Unit tests (`tests/unit/`)

Pure-function logic in isolation. The Microsoft Graph client and the keyring are both **mocked**. No network, no credentials, sub-second per test. Run on every save during development, on every PR in CI.

- Tool argument parsing and validation.
- Token-cache logic with a mocked keyring backend.
- Path resolution from site-relative paths to drive/item IDs (logic, not the actual Graph call).
- Error-mapping (Graph error codes → MCP tool errors with the right shape).

### Integration tests (`tests/integration/`)

How our internal modules fit together. Mocks **at the system boundary** are acceptable: a mock HTTP server (e.g., `respx` or a small `httpx` mock) standing in for `graph.microsoft.com`. Deterministic, runnable in CI without any real Microsoft credentials.

- The MCP tool layer correctly routes calls to the auth + Graph-client layers.
- ETag round-trip (`sp_open` → `sp_save`) works against a recorded fixture.
- Crash-recovery: `sp_status` reconciles registry against a mock-Graph-state.
- Read-only mode: when `SP_ALLOW_WRITES` is unset, write tools are not registered.

### Harness tests (`tests/harness/`) — the AI-development enabler

**Real connection to a real SharePoint sandbox**, configured the same way an ISMS-relevant production library would be (checkout-required, retention enabled, audit-on). Real OAuth Device Code flow against Microsoft Identity. Real account, **least-privilege scoped to a single test site**.

- **Sandbox**: a dedicated SharePoint site `sharepoint-mcp-harness` in the XMV tenant (or a dedicated dev tenant, TBD).
- **Test account**: a service-purpose user in XMV's tenant — name TBD — granted **only** `Edit` permission on the harness site, **no** access to other XMV resources. A leaked harness refresh token must not put any production data at risk.
- **Initial setup**: the human admin (David) walks the test account through a one-time Device Code login on the agent's working machine. The resulting refresh token is cached in the agent's OS keyring under profile `harness`. The same refresh token is also stored as a GitHub Actions secret for CI.
- **What's covered**:
  - End-to-end auth: Device Code → keyring → refresh-token-loop survives a fresh process start.
  - Each tool (`sp_search`, `sp_list`, `sp_read`, `sp_open`, `sp_save`, `sp_release`, `sp_status`) hit at least once against the live API.
  - Lock semantics: `sp_open` of an already-checked-out file fails with the right error; concurrent `sp_open` attempts are exclusive.
  - ETag stale-write detection: edit-and-save after another process has changed the same file → fails as expected.
  - Audit-log attribution: the test-account user appears as actor in the SharePoint audit log for each write.
- **Run sites**: from the agent's working machine on every iteration; from CI on every PR (using the secret-stored refresh token).

### Test-environment hierarchy

| Environment | Purpose | Auth |
|---|---|---|
| Local (developer machine) | Unit + integration during edit-iteration | None / mocked |
| Local + harness sandbox | Harness during iteration on tool-level changes | Device Code → keyring (one-time human login) |
| CI on PR | Unit + integration + harness | Refresh token from GitHub Secrets |
| First-user PyPI install | Production for end users; not under our test control | Their own M365 account, their own tenant |

There is no "staging" environment in the cluster sense — the harness sandbox **is** the staging-equivalent for this project. The next step beyond harness is end-user installation.

### What is NOT in scope of the test suite

- Real-world tenants other than the harness sandbox. We don't run automated tests against XMV's production data, customer tenants, or any data we don't fully own.
- Microsoft itself (uptime, scope-policy changes) — we observe these through harness failures and respond with code/doc fixes, not by trying to mock around them.
- Cross-MCP-client compatibility (Claude Desktop, other MCP clients) — the protocol is the contract; we test against `mcp` Python SDK conformance, not specific clients.

---

## MVP scope (v0.1)

- Tools: `sp_search`, `sp_list`, `sp_read`, `sp_open`, `sp_save`, `sp_release`, `sp_status`.
- Single-tenant focused; multi-tenant via launching multiple processes (no in-process tenant switching).
- Device code auth, keyring token cache.
- Python 3.11+, packaged for `uvx`/`pipx` install.
- Test layers: unit + integration locally and in CI; **harness tests against the dedicated XMV sandbox site** with a least-privilege test user (see Testability section above).
- Read-only by default; writes opt-in via `SP_ALLOW_WRITES=true`.

Deferred to v0.2:

- `sp_history`, `sp_get_version`.
- `sp_publish` (upload a new file from local) — covers the "draft + promote" use case from § Core use cases. Not in v0.1 because it is independent of the checkout/checkin chain that v0.1 focuses on.
- Service-principal auth.
- Bulk operations (`sp_save_many`).
- OneNote / Excel-cell-level tools (likely separate MCPs).

---

## Why XMV OSS

- Generally useful for any Linux-based AI dev workflow that touches SharePoint — wider problem than the original use case that motivated it.
- Reusable across XMV's own customer engagements.
- Forcing function for our own discipline: if we're going to let agents edit ISMS-relevant docs, the mechanics had better be inspectable.
- Fits the existing XMV public-repo pattern (small, focused, Linux-headless-friendly tools).

---

## Open questions for the tech spike

> **Resolved 2026-05-06:**
>
> - `msgraph-sdk-python` vs raw `httpx` — raw `httpx`. See [docs/spikes/2026-05-06-graph-sdk-vs-httpx.md](spikes/2026-05-06-graph-sdk-vs-httpx.md). Footprint difference is ~92× (220 MB → 2.4 MB), the SDK fights our keyring-owned auth contract, and the 6 endpoints we touch don't justify a generic typed client.
> - Token storage on headless Linux — **three-tier**: `keyring` when real, `PlainFileTokenStore` (mode 0600) as default fallback, `EncryptedFileTokenStore` opt-in via `SP_TOKEN_PASSPHRASE`. See [docs/spikes/2026-05-06-keyring-vs-encrypted-file.md](spikes/2026-05-06-keyring-vs-encrypted-file.md). Auto-detect at first use; `SP_TOKEN_STORE=keyring|file|encrypted-file` to override. **No env vars needed for the typical install.**
>
> **Resolved 2026-05-07:**
>
> - Working-directory cleanup policy on crash — **persistent registry + explicit `sp_release`** for v0.1; no TTL, no automatic server-side reconciliation. See [docs/spikes/2026-05-07-working-dir-cleanup.md](spikes/2026-05-07-working-dir-cleanup.md). The registry survives crashes; `sp_status` shows what was open; the agent / human chooses to resume or release.
> - Chunked-upload threshold for `sp_save` — **single-shot `PUT /content`** for v0.1, supports up to 250 MB per Microsoft's current SharePoint cap. See [docs/spikes/2026-05-07-chunked-upload-threshold.md](spikes/2026-05-07-chunked-upload-threshold.md). Resumable upload sessions deferred to v0.2 unless a real-world file >250 MB shows up.

### Still open

All spike questions resolved for v0.1. Future spikes track separately under the v0.2 milestone when surfaced.
