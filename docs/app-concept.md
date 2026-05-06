<!-- SPDX-License-Identifier: MIT OR Apache-2.0 -->

# sharepoint-mcp — App Concept

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
│   Claude Code    │ ◄──────────────────────► │  sharepoint-mcp     │
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

1. First run prints a device code and verification URL.
2. User opens URL on phone or laptop, enters code, signs in.
3. Token is cached locally via OS keyring (`keyring` Python lib) — no plaintext on disk.
4. Refresh token used silently afterwards; re-prompt on full expiry (60–90 days).

Per-tenant config holds: `tenant_id`, `client_id` (a public-client app registration in that tenant), required scopes (`Files.ReadWrite.All`, `Sites.ReadWrite.All`).

Service-principal / client-credentials flow as a future extension for unattended automation, but **MVP is interactive user auth** — every action is attributed to the signed-in user, which is what compliance wants.

---

## Multi-tenant pattern

One MCP instance per SharePoint tenant. The consuming repo's `.mcp.json`:

```json
{
  "mcpServers": {
    "sharepoint-anqer": {
      "command": "uvx",
      "args": ["sharepoint-mcp"],
      "env": {
        "SP_TENANT_ID": "<tenant-id>",
        "SP_CLIENT_ID": "<app-registration-id>",
        "SP_PROFILE": "anqer"
      }
    }
  }
}
```

Token caches and working directories are namespaced by `SP_PROFILE`. A second customer = a second `mcpServers` entry with its own profile. Tools appear in the agent as `mcp__sharepoint-anqer__sp_search` etc.

---

## Conflict and safety model

- **Library-level checkout requirement** is honoured: if the library forces explicit checkout, `sp_save` without prior `sp_open` returns an error.
- **Lock is real**: `sp_open` calls Graph `/checkout`, which actually locks the file in SharePoint. Other users see "checked out by [user]" until release.
- **Stale-edit detection**: `sp_save` includes the `If-Match` ETag from `sp_open`; mismatch → error, agent must re-`open` to reconcile.
- **Crash recovery**: `sp_status` is the source of truth for what's checked out. If the MCP crashed mid-edit, `sp_release` (or admin discardCheckout via web) cleans up.
- **No silent overwrites**: there is no path that writes to SharePoint without an explicit `sp_save` call from the agent, with a comment.

---

## Out-of-band concerns

- **Audit attribution**: actions appear in SharePoint audit log as the signed-in user with the `sharepoint-mcp` user-agent string. Useful for distinguishing AI-mediated edits from manual ones.
- **Telemetry**: none by default. Opt-in structured logging to stderr; consuming app routes it.
- **Compliance scope**: this MCP does not pretend to be a controlled processing system itself. It is a thin pass-through to Graph; SharePoint's existing controls remain authoritative.

---

## MVP scope (v0.1)

- Tools: `sp_search`, `sp_list`, `sp_read`, `sp_open`, `sp_save`, `sp_release`, `sp_status`.
- Single-tenant focused; multi-tenant via launching multiple processes (no in-process tenant switching).
- Device code auth, keyring token cache.
- Python 3.11+, packaged for `uvx`/`pipx` install.
- Tests: integration tests against a dedicated test SharePoint site (provided via env vars in CI).

Deferred to v0.2:

- `sp_history`, `sp_get_version`.
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

1. `msgraph-sdk-python` vs raw `httpx` calls? SDK is heavy; raw is ~6 endpoints we actually need.
2. Token storage: keyring (cross-platform but requires DBus / Secret Service on Linux) vs encrypted file with a passphrase. Keyring assumes a desktop session — may not work on truly headless servers without `gnome-keyring-daemon` running.
3. Working-directory cleanup policy on crash: TTL-based, or explicit reconciliation against `sp_status`?
4. How big can a single edit be before we need chunked upload (Graph's resumable upload session)? Default cutoff is 4 MB.
