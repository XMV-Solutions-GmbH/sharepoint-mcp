<!-- SPDX-License-Identifier: MIT OR Apache-2.0 -->

# mcp-server-sharepoint

[![PyPI version](https://img.shields.io/pypi/v/mcp-server-sharepoint?color=0E7EE0)](https://pypi.org/project/mcp-server-sharepoint/)
[![Python versions](https://img.shields.io/pypi/pyversions/mcp-server-sharepoint?color=0E7EE0)](https://pypi.org/project/mcp-server-sharepoint/)
[![Downloads](https://img.shields.io/pypi/dm/mcp-server-sharepoint?label=downloads%2Fmonth&color=0E7EE0)](https://pypi.org/project/mcp-server-sharepoint/)
[![Licence](https://img.shields.io/badge/licence-MIT%20OR%20Apache--2.0-blue.svg)](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/blob/main/LICENSE)
[![CI](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/actions/workflows/ci.yml)
[![Coverage](https://img.shields.io/codecov/c/github/XMV-Solutions-GmbH/sharepoint-mcp/main)](https://codecov.io/gh/XMV-Solutions-GmbH/sharepoint-mcp)
[![Contributions welcome](https://img.shields.io/badge/contributions-welcome-brightgreen.svg)](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues)

> **In one sentence:** an [MCP](https://modelcontextprotocol.io) server that lets AI coding agents like Claude Code edit files in your SharePoint document libraries **the same way a careful human would** — with proper checkout, comments, version history, and lock conflicts — instead of overwriting things and breaking your audit trail.

## What is this for?

You have important documents in SharePoint — ISO 27001 controls, contract templates, ISMS records, runbooks, a quality manual. They have version history, audit trails, retention policies, and someone above you cares that they stay compliant.

You'd love to let AI agents help you draft and update these documents — they're great at it. But every other "AI + SharePoint" tool you've tried makes the audit log say *"a robot named rclone updated this file at 03:42 with no comment"* — which means your auditor has questions.

**`mcp-server-sharepoint` makes the audit log say what actually happened**: *"David Koller updated this file at 14:32. Comment: ‘Tightened control wording for A.5.1 per CISO review.’ Version: 3.5 → 3.6 (minor)."* Because the AI agent uses the same Microsoft-blessed checkout/checkin model a human Office user would.

Concretely, the agent gets these tools:

| Tool | What the agent does | What ends up in SharePoint |
|---|---|---|
| `sp_search`, `sp_list`, `sp_read` | finds and reads files | nothing changes |
| `sp_open` | acquires a checkout lock + downloads | "checked out by *you*" appears for everyone else |
| `sp_save` | uploads + checks in with a comment | a real new version with a real comment in the audit log |
| `sp_release` | discards a checkout | lock released, no version created |
| `sp_status` | shows what's currently checked out by this agent | nothing changes |

Every action is attributed to the human who signed in once via Microsoft's standard Device Code login. No service-account "robot" identity. No silent overwrites. No broken locks. Lock conflicts are reported as conflicts. ETag checks catch concurrent edits before they clobber.

## Installation

```bash
pip install mcp-server-sharepoint
# or, with uv (recommended):
uv tool install mcp-server-sharepoint
# or, on the fly without installing globally:
uvx mcp-server-sharepoint --help
```

Requires Python 3.11+. Works on Linux, macOS, Windows.

## Quickstart

### 1. Sign in once (out of band)

```bash
uvx mcp-server-sharepoint login
```

Output looks like:

```text
Sign in to mcp-server-sharepoint via the Device Code flow:
Open the URL in a browser and type the code.

     URL:   https://login.microsoft.com/device
     Code:  D2LKUY4AV

Waiting for sign-in...
```

Open the URL in any browser, type the code, sign in with your M365 account. Your refresh token is cached locally — see [Token storage](#token-storage). The MCP server itself never blocks for human interaction afterwards.

### 2. Wire it into Claude Code

In your project's `.mcp.json`:

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

Restart Claude Code. The agent now has `sp_search`, `sp_list`, `sp_read`, `sp_status` available — **read-only by default**.

### 3. Enable writes (when you're ready)

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "uvx",
      "args": ["mcp-server-sharepoint"],
      "env": { "SP_ALLOW_WRITES": "true" }
    }
  }
}
```

Now `sp_open`, `sp_save`, `sp_release` are also available.

### 4. Try it

```text
You:    Find our latest ISO 27001 control A.5.1 policy in SharePoint.
Agent:  [calls sp_search → finds it]
        Found "iso27001-A.5.1.md" at https://contoso.sharepoint.com/...

You:    Read it, suggest two improvements based on the new revision of the standard.
Agent:  [calls sp_read → reads file → suggests in chat]

You:    Apply them and save with a comment summarising the changes.
Agent:  [calls sp_open → modifies → sp_save with comment]
        Saved version 1.4. Comment recorded: "Tightened wording per ISO 27001:2022 to match new control objective; added cross-reference to A.5.2."
```

Each tool call gets a permission prompt in Claude Code (you can mark trusted ones as "always allow" per session). Read tools are flagged read-only; write tools are flagged destructive — you see the difference.

## What it can do, in detail

### Read tools (always available)

| Tool | Purpose |
|---|---|
| `sp_search(query, site?, folder?, file_type?, modified_after?)` | KQL-style search across SharePoint sites the user has access to. Returns hits with name, path, web URL, last-modified date, author. |
| `sp_list(url)` | List a SharePoint folder's children (files + sub-folders) with size, type, last-modified. URL is the human-readable web URL. |
| `sp_read(url)` | Download a file's content to a local temp file with the original extension preserved. **Read-only — does NOT acquire a checkout.** |
| `sp_status(verify=False)` | Show what files this agent currently has checked out, when, and where the local working copies are. With `verify=True`, additionally queries SharePoint to confirm the server-side lock state — adds `server_locked` (`true`/`false`/`null`) and `lock_holder` (display name) to each entry. Costs one Graph call per registry entry. |
| `sp_sites(query?)` | Discover SharePoint sites the user can see. `query` is a free-text site-name search; omit to list all. Useful as the agent's starting point when no site URL is known yet. |
| `sp_subsites(parent_site_url)` | List immediate sub-sites under a parent site URL. Recurse on each result's `web_url` to walk deeper. |
| `sp_followed_sites()` | List sites the user has Followed in SharePoint — a curated "my SharePoint" entry point. Not available in service-principal mode (no signed-in user). |
| `sp_drives(site_url)` | List the document libraries (drives) on a site — default Shared Documents plus Site Assets, Style Library, and any custom libraries. Most read/write tools accept URLs into any library transparently; `sp_drives` is the discovery step when the agent doesn't yet know which libraries exist. |
| `sp_trash_list(site_url)` | List items in the SharePoint site's recycle bin (id, name, size, deleted_date_time, deleted_from_location, deleted_by). Read-only. *Uses Graph beta endpoint — see note below.* |

#### Non-default libraries

URLs into **non-default document libraries** (Site Assets, Style Library, custom libraries) work transparently across `sp_list`, `sp_read`, `sp_open`, `sp_save`, `sp_publish`, etc. The resolver tries the default Shared Documents drive first; on a 404, it lists the site's drives, matches the URL's first path segment to a library name, and retries against that library. One extra Graph round-trip per first-look-up at a non-default library — acceptable cost for the convenience.

### Write tools (opt-in via `SP_ALLOW_WRITES=true`)

| Tool | Purpose |
|---|---|
| `sp_open(url)` | **Acquire a server-side checkout lock** + download the current content to a working-directory path. Other users see "checked out by *you*" until you save or release. Fails with a clear error if someone else already holds the lock. |
| `sp_save(url, comment, version="minor"\|"major")` | Upload your changes + check the file back in with an audit comment + new version. **`comment` is required and must be non-empty** — describes what changed for the audit log. ETag round-trip catches "someone else changed the file underneath us" and refuses to clobber. |
| `sp_release(url)` | Discard a pending checkout: drop the lock server-side and delete the local working copy. Use when you decide *not* to keep your edits. |
| `sp_open_many(urls)` | Bulk variant of `sp_open` — acquires checkouts on multiple files in parallel (up to 4 concurrent Graph calls per Microsoft throttling guidance). Returns one result per URL: `{path, status: "ok"\|"error", local_path?, error?}`. Per-file failures don't abort the batch. Honors `Retry-After` on 429/503. |
| `sp_save_many(operations)` | Bulk variant of `sp_save` — each op `{url, comment, version?}`. Same parallel/error-isolation semantics as `sp_open_many`. ETag round-trip applies per file. |
| `sp_trash_restore(site_url, item_id)` | Restore an item from the site's recycle bin to its original location. `item_id` comes from `sp_trash_list`. Not destructive (recovers data) but does change observable site state. *Uses Graph beta endpoint — see note below.* |

#### Recycle bin: beta endpoint

`sp_trash_list` and `sp_trash_restore` currently call Microsoft Graph's `/beta` endpoints — the site-level recycle-bin API has not yet been promoted to v1.0. Stable enough that production tools (SharePoint web UI, admin center) rely on it, but Microsoft reserves the right to change the schema. We pin to the documented beta shape and will migrate to v1.0 when it lands.

#### Large files

`sp_save` uses Microsoft Graph's [resumable upload session](https://learn.microsoft.com/en-us/graph/api/driveitem-createuploadsession) for files larger than **100 MB** (configurable via `SP_CHUNKED_UPLOAD_THRESHOLD_MB`). Files at-or-below the threshold use a single-shot `PUT /content` for a faster path. Microsoft caps single-shot at 250 MB; the resumable path supports up to 250 GB. Chunks are 5 MiB and retry on transient 5xx / connection errors with exponential backoff.

### Authentication

- **OAuth 2.0 Device Code flow** against Microsoft Identity (default). You sign in once; the refresh token is cached locally and silently renewed (~60–90 days until full re-login).
- **Bring-your-own-app or use ours.** XMV publishes a multi-tenant Entra app registration that's baked in as the default — same pattern as Azure CLI / GitHub CLI. Tenants with strict app-allowlisting can override via `SP_CLIENT_ID` and `SP_TENANT_ID` env vars.
- **Token storage** is auto-detected at first use: OS keyring (macOS Keychain / Windows Credential Locker / Linux Secret Service) when available, mode-0600 plain JSON file as fallback (same convention as `gh auth`, `aws configure`). Optional encryption with `SP_TOKEN_PASSPHRASE` for paranoid setups or CI.
- **Multi-customer / multi-tenant**: separate `SP_PROFILE` per tenant, each with its own token cache.

#### Service-principal mode (unattended automation)

For CI / scheduled jobs where no human is in the loop, run with `SP_AUTH_MODE=service-principal` (or just set `SP_CLIENT_SECRET` — auto-detected). Required env vars: `SP_CLIENT_ID`, `SP_CLIENT_SECRET`, `SP_TENANT_ID`. The app registration must have **Application** Microsoft Graph permissions (`Files.ReadWrite.All`, `Sites.ReadWrite.All`) with admin consent recorded.

Tradeoff: every action is attributed to the *application* principal in SharePoint's audit log, NOT a real user. The compliance-friendly default stays delegated user auth — only switch when no human is in the loop.

### Security model

Three layers of "don't accidentally damage anything":

1. **Your MCP client (Claude Code) prompts before each tool call by default.** Read tools are flagged read-only; write tools are flagged destructive — you see the difference at the prompt.
2. **Read-only by default at our server.** Without `SP_ALLOW_WRITES=true`, the write tools aren't even registered. The agent literally can't see them.
3. **`sp_save` requires a non-empty audit comment.** The agent has to articulate intent, and that lands in the SharePoint audit log.

The threat model is "your local OS account is trusted" — same as `~/.ssh/id_rsa`, `gh` tokens, `aws` config. The tool isn't designed to defend against host compromise; it's designed to keep audit trails honest under normal use.

## Roadmap

| Version | Status | Theme | Highlights |
|---|---|---|---|
| **v0.1** | ✅ released 2026-05-07 | Audit-preserving doc edits | The seven `sp_*` tools above, three-layer test harness, Trusted-Publisher PyPI release pipeline, branch-protected `main`. |
| **v0.2** | ✅ released 2026-05-07 | Write-side polish | `sp_publish` (upload new file), `sp_history` + `sp_get_version` (version-history access), `sp_open_many` + `sp_save_many` (bulk operations with concurrency cap), `sp_status(verify=True)` server-side reconciliation, resumable uploads for files >100 MB (auto-switch), service-principal auth for unattended automation. |
| **v0.3** | 📋 queued | Broader SharePoint surface | SharePoint Lists CRUD (custom lists, issue trackers, etc.), modern Pages read/edit, permissions inspection, sharing-link creation, multi-library access, site discovery, recycle-bin restore, delta queries for change tracking. |
| **v0.4** | 🤔 maybe | Admin functions | Site / library / permission administration, IF customer demand emerges. |
| **v1.0** | 🎯 stability lock-in | "API stable, production-tested" | After v0.x has been used in real customer environments for ~3–6 months without breaking changes. Not "more features" — a **commitment that what you depend on today still works tomorrow.** |

The full ticket-by-ticket plan lives at the [issues page](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues).

## Multi-profile pattern

For consultancy workflows with multiple SharePoint tenants, give each its own profile so the token caches don't collide:

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
      "env": { "SP_PROFILE": "globex" }
    }
  }
}
```

Sign each one in separately:

```bash
uvx mcp-server-sharepoint login --profile acme
uvx mcp-server-sharepoint login --profile globex
```

Tools appear in Claude as `mcp__sharepoint-acme__sp_search` etc. Cross-tenant accidents don't happen because the tokens are namespaced.

## BYO Entra app registration

Tenants with strict app-allowlisting can override the bundled multi-tenant default:

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "uvx",
      "args": ["mcp-server-sharepoint"],
      "env": {
        "SP_TENANT_ID": "<your-tenant-guid>",
        "SP_CLIENT_ID": "<your-app-registration-guid>"
      }
    }
  }
}
```

The app registration must be: multi-tenant or single-tenant, public client (no secret), Device Code flow allowed, with delegated permissions `Files.ReadWrite.All`, `Sites.ReadWrite.All`, `User.Read`, `offline_access`.

## Token storage

Three backends, auto-detected at first use:

| Tier | Backend | When | Setup |
|---|---|---|---|
| 1 | OS keyring | macOS Keychain / Windows Credential Locker / Linux with Secret Service | none |
| 2 | Plain file `~/.cache/sharepoint-mcp/<profile>/token.json` mode 0600 | Headless Linux default | none |
| 3 | Encrypted file (Fernet, Scrypt KDF) | When `SP_TOKEN_PASSPHRASE` is set | env var |

Force a specific backend with `SP_TOKEN_STORE=keyring|file|encrypted-file`. See [the spike doc](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/blob/main/docs/spikes/2026-05-06-keyring-vs-encrypted-file.md) for the rationale — short version: same security model as `gh auth`, `aws configure`, `npm login`.

## Troubleshooting

### "No usable credentials"

The cached token expired (refresh tokens last ~60–90 days) or never existed. Run:

```bash
uvx mcp-server-sharepoint login --profile <name>
```

### "Cannot checkout: file is already checked out by another user"

Someone else (or a previous instance of your own agent) has the file locked. Wait, or in the SharePoint web UI go to the library → file → "Discard check-out".

### "File changed under us between sp_open and sp_save"

Your agent had the file open, but someone else edited it before your save. Recover with:

```text
sp_release(url)        # drop your stale working copy + lock
sp_open(url)           # acquire fresh lock + content
# re-apply edits to the new content
sp_save(url, comment="…", version="minor")
```

### Linux: keyring fails / "Secret Service unavailable"

The plain-file backend kicks in automatically — no action needed. If you'd rather have encryption at rest:

```bash
export SP_TOKEN_PASSPHRASE='<some-strong-passphrase>'
uvx mcp-server-sharepoint login --profile <name>
```

### Recovery after a crash

```bash
uvx mcp-server-sharepoint     # restart the server
```

In the agent, ask `sp_status` — you'll see anything that was checked out before the crash. For each, either resume work (working file is still on disk; `sp_save` works as normal) or drop it (`sp_release`).

The registry survives crashes; nothing is lost.

## Development

```bash
git clone https://github.com/XMV-Solutions-GmbH/sharepoint-mcp.git
cd sharepoint-mcp
uv sync --extra dev

# Unit + integration (no real SharePoint), with coverage reporting
./tests/run_tests.sh

# Harness (real SharePoint sandbox; requires harness-profile login)
./tests/run_tests.sh harness
```

| Document | What's in it |
|---|---|
| [`docs/app-concept.md`](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/blob/main/docs/app-concept.md) | Vision, MVP scope, MCP tool surface, auth, conflict model |
| [`docs/testconcept.md`](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/blob/main/docs/testconcept.md) | Three-layer test strategy (unit / integration / harness) |
| [`docs/RELEASING.md`](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/blob/main/docs/RELEASING.md) | How releases happen |
| [`ENGINEERING_PRINCIPLES.md`](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/blob/main/ENGINEERING_PRINCIPLES.md) | Project-agnostic engineering baseline |
| [`CLAUDE.md`](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/blob/main/CLAUDE.md) | Project-specific overrides |
| [`docs/spikes/`](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/tree/main/docs/spikes) | Design-decision history (httpx vs SDK, token storage, etc.) |

## Contributing

Contributions are welcome. Please read [CONTRIBUTING.md](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/blob/main/CONTRIBUTING.md) and the [Code of Conduct](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/blob/main/CODE_OF_CONDUCT.md) first.

Bug reports and feature requests go to [GitHub Issues](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues).

## Licence

Dual-licensed under either of:

- Apache License, Version 2.0 ([LICENSE-APACHE](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/blob/main/LICENSE-APACHE) or <http://www.apache.org/licenses/LICENSE-2.0>)
- MIT licence ([LICENSE-MIT](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/blob/main/LICENSE-MIT) or <http://opensource.org/licenses/MIT>)

at your option.

Unless you explicitly state otherwise, any contribution intentionally submitted for inclusion in this project by you, as defined in the Apache-2.0 license, shall be dual licensed as above, without any additional terms or conditions.

## Contact

- **Organisation**: XMV Solutions GmbH
- **Email**: <oss@xmv.de>
- **Website**: <https://xmv.de/en/oss/>
- **GitHub**: [@XMV-Solutions-GmbH](https://github.com/XMV-Solutions-GmbH)
