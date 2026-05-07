<!-- SPDX-License-Identifier: MIT OR Apache-2.0 -->

# mcp-server-sharepoint

[![Licence](https://img.shields.io/badge/licence-MIT%20OR%20Apache--2.0-blue.svg)](LICENSE)
[![CI](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/actions/workflows/ci.yml)
[![contributions welcome](https://img.shields.io/badge/contributions-welcome-brightgreen.svg?style=flat)](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues)

A **Model Context Protocol** server for SharePoint document libraries. Lets AI coding agents (Claude Code, Claude Desktop, any MCP client) read and edit files on SharePoint **without breaking version history, audit trail, or locking semantics** — by wrapping Microsoft Graph's native checkout / edit / checkin model as MCP tools.

> **Status:** v0.1 alpha. All MVP tools work end-to-end against real SharePoint. Not yet on PyPI. Full background in [docs/app-concept.md](docs/app-concept.md).

---

## Why

The standard alternatives — `rclone`, WebDAV mounts, the Anthropic-hosted M365 MCP — either skip SharePoint's checkout/checkin model entirely or expose it only for search and read. That's not acceptable for documents under retention or audit constraints (ISMS records, controlled procedures, contract templates).

`mcp-server-sharepoint` keeps the audit trail intact: every edit goes through an explicit `sp_open` → `sp_save` (or `sp_release`) cycle, attributed to the signed-in user, with a commit message and an honest version bump. Lock conflicts are reported as conflicts. ETag-based stale-write detection prevents silent overwrites.

Full rationale in [docs/app-concept.md](docs/app-concept.md).

---

## MCP tools

| Tool | What it does | Annotations |
|---|---|---|
| `sp_search(query, …)` | KQL-style search across visible SharePoint sites | read-only, idempotent |
| `sp_list(url)` | List children of a folder | read-only, idempotent |
| `sp_read(url)` | Download file content to a local temp path | read-only, idempotent |
| `sp_status()` | Show files currently checked out by this profile | read-only, idempotent |
| `sp_open(url)` | **Acquire checkout lock** + download to working dir | non-destructive lock |
| `sp_save(url, comment, version="minor"\|"major")` | Upload + checkin with audit comment + ETag | destructive |
| `sp_release(url)` | discardCheckout + drop local working copy | destructive |

**Read-only by default.** The four read tools are always registered. The three write tools (`sp_open`/`sp_save`/`sp_release`) are only registered when `SP_ALLOW_WRITES=true` (or `1`/`yes`/`on`) is set — belt-and-suspenders to your MCP client's per-call permission prompts.

Each tool is registered with proper [MCP `ToolAnnotations`](https://modelcontextprotocol.io/specification/server/tools#tool-annotations) so Claude Code's permission UI renders the right confirmation prompt for each.

---

## Quickstart

### 1. Install

```bash
# Once published on PyPI:
uvx mcp-server-sharepoint --help

# For now, from a local checkout:
git clone https://github.com/XMV-Solutions-GmbH/sharepoint-mcp.git
cd sharepoint-mcp
uv sync
```

### 2. Sign in once (out of band)

```bash
uv run mcp-server-sharepoint login
```

You'll see something like:

```text
Sign in to mcp-server-sharepoint via the Device Code flow:
Open the URL in a browser and type the code.

     URL:   https://login.microsoft.com/device
     Code:  D2LKUY4AV

Waiting for sign-in...
```

Open the URL in any browser, type the code, sign in with your M365 account. The refresh token gets cached locally — see [Token storage](#token-storage) for where exactly. The MCP server itself never blocks for human interaction.

### 3. Wire into your MCP client

For **Claude Code**, add to `.mcp.json`:

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

Restart Claude Code. The agent now has `sp_search`, `sp_list`, `sp_read`, `sp_status` available. To enable write tools too:

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

### 4. Use it

In Claude Code, ask the agent something like:

```text
Find the latest version of our ISO 27001 control A.5.1 policy in
SharePoint, read it, suggest improvements, and save the updated
version with a comment summarising the changes.
```

The agent will: `sp_search` → `sp_open` → modify → `sp_save`. Each tool call gets a permission prompt in Claude Code (read tools too — Claude Code prompts on first use of any external tool by default). After a couple of approvals you can switch to "always allow" per tool.

---

## Multi-profile / multi-customer

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
uv run mcp-server-sharepoint login --profile acme
uv run mcp-server-sharepoint login --profile globex
```

Tools appear in Claude as `mcp__sharepoint-acme__sp_search` etc. Cross-tenant accidents don't happen because the tokens are namespaced.

---

## BYO Entra app registration

The package ships with a baked-in `client_id` for an XMV-Solutions-published multi-tenant Entra app (`cb7cf68d-90d5-4841-90a7-de3a40be280b`). End users normally don't need to think about this — same pattern as Azure CLI, GitHub CLI.

If your tenant has strict app-allowlisting and refuses unknown publishers, register your own and override:

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

The app must be:

- Multi-tenant or single-tenant for your tenant
- Public client (no secret), Device Code flow allowed
- Delegated permissions: `Files.ReadWrite.All`, `Sites.ReadWrite.All`, `User.Read`, `offline_access`

---

## Token storage

Three backends, auto-detected at first use:

| Tier | Backend | When | Setup |
|---|---|---|---|
| 1 | OS keyring | macOS Keychain / Windows Credential Locker / Linux with Secret Service | none |
| 2 | Plain file `~/.cache/sharepoint-mcp/<profile>/token.json` mode 0600 | Headless Linux default | none |
| 3 | Encrypted file (Fernet, Scrypt KDF) | When `SP_TOKEN_PASSPHRASE` is set | env var |

Force a specific backend with `SP_TOKEN_STORE=keyring|file|encrypted-file`. See [the spike doc](docs/spikes/2026-05-06-keyring-vs-encrypted-file.md) for the rationale — short version: same security model as `gh auth`, `aws configure`, `npm login`. Encrypted-file is opt-in for CI / paranoid setups.

---

## Security model

This is the part to read carefully if you're putting the MCP near ISMS-relevant content.

**What this tool does:**

- Runs locally as a child process of your MCP client (Claude Code etc.).
- Talks directly to `login.microsoftonline.com` and `graph.microsoft.com` over HTTPS.
- Uses your delegated OAuth token. **Every action is attributed to your M365 account in SharePoint's audit log.**

**What it does not do:**

- No telemetry. No XMV-controlled servers in the loop. No data leaves your machine except the direct calls to Microsoft.
- No service-principal / client-credentials flow in v0.1 — interactive user auth only. Compliance-friendly: no "rclone client" entries in audit logs.
- No sharing-link generation, no permission changes, no library / site administration.

**Three layers of protection against accidental damage:**

1. **MCP-client permission prompts.** Claude Code prompts before each tool call by default. Even if you approve "always allow" on a write tool, it knows the tool's `destructiveHint` is `True` and can present that.
2. **Read-only by default at the server.** Without `SP_ALLOW_WRITES=true`, the write tools aren't registered. Claude can't accidentally call `sp_save` if it never sees `sp_save`.
3. **`sp_save` always requires a non-empty comment.** The audit log gets meaningful "what changed" annotations, not blank entries. The agent has to articulate intent.

**Threat model:** if your local user account is compromised, the attacker has your tokens (same as your SSH keys, your gh CLI tokens, your aws CLI tokens). The tool is not designed to defend against host compromise. It is designed to keep audit trails honest and locks correct under normal use.

---

## Troubleshooting

### "No usable credentials" / `AuthRequiredError`

The cached token expired (refresh tokens last ~60–90 days) or never existed. Run:

```bash
uv run mcp-server-sharepoint login --profile <name>
```

Then retry whatever tool call failed.

### "Cannot checkout: file is already checked out by another user"

Someone else (or a previous instance of your own agent) has the file locked. Wait for them to release, or in the SharePoint web UI go to the library → file → "Discard check-out".

### "File changed under us between sp_open and sp_save"

Your agent had the file open, but someone else edited it before your save. Recover with:

```text
sp_release(url)        # drop your stale working copy + lock
sp_open(url)           # acquire fresh lock + content
# re-apply your edits to the new content
sp_save(url, comment="…", version="minor")
```

### Linux: keyring fails / "Secret Service unavailable"

That's fine. The plain-file backend kicks in automatically — no action needed. If you'd rather have encryption at rest:

```bash
export SP_TOKEN_PASSPHRASE='<some-strong-passphrase>'
uv run mcp-server-sharepoint login --profile <name>
```

Then keep `SP_TOKEN_PASSPHRASE` exported (in `~/.bashrc`, `direnv`, etc.) for subsequent runs to decrypt.

### Recovery after a crash

```bash
uv run mcp-server-sharepoint     # restart the server
```

In the agent, ask `sp_status` — you'll see anything that was checked out before the crash. For each:

- Resume work: working file is still on disk; `sp_save` works as normal.
- Drop it: call `sp_release`.

The registry survives crashes; nothing is lost.

---

## Development

```bash
git clone https://github.com/XMV-Solutions-GmbH/sharepoint-mcp.git
cd sharepoint-mcp
uv sync --extra dev

# Unit + integration (no real SharePoint)
./tests/run_tests.sh

# Harness (real SharePoint sandbox; requires harness-profile login)
./tests/run_tests.sh harness
```

Project layout, testing strategy, and engineering principles:

- [docs/app-concept.md](docs/app-concept.md) — vision, MVP scope, MCP tool surface, auth, conflict model
- [docs/testconcept.md](docs/testconcept.md) — three-layer test strategy (unit / integration / harness)
- [ENGINEERING_PRINCIPLES.md](ENGINEERING_PRINCIPLES.md) — project-agnostic engineering baseline
- [CLAUDE.md](CLAUDE.md) — project-specific overrides
- [docs/spikes/](docs/spikes/) — design-decision history (httpx vs SDK, token storage, etc.)

---

## Contributing

Contributions are welcome. Please read [CONTRIBUTING.md](CONTRIBUTING.md) and the [Code of Conduct](CODE_OF_CONDUCT.md) first.

---

## Licence

Dual-licensed under either of:

- Apache License, Version 2.0 ([LICENSE-APACHE](LICENSE-APACHE) or <http://www.apache.org/licenses/LICENSE-2.0>)
- MIT license ([LICENSE-MIT](LICENSE-MIT) or <http://opensource.org/licenses/MIT>)

at your option.

Unless you explicitly state otherwise, any contribution intentionally submitted for inclusion in this project by you, as defined in the Apache-2.0 license, shall be dual licensed as above, without any additional terms or conditions.

---

## Contact

- **Organisation**: XMV Solutions GmbH
- **Email**: <oss@xmv.de>
- **Website**: <https://xmv.de/en/oss/>
- **GitHub**: [@XMV-Solutions-GmbH](https://github.com/XMV-Solutions-GmbH)
