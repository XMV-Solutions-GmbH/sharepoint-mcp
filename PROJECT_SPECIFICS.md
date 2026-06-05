<!--
SPDX-License-Identifier: MIT OR Apache-2.0
SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
SPDX-FileContributor: David Koller <david.koller@xmv.de>
-->

# PROJECT_SPECIFICS.md — `mcp-server-sharepoint`

Project-specific content for `mcp-server-sharepoint`. Read after `AGENTS.md` per its reading order. Everything in here is specific to this repo; the generic agent rules live in `AGENTS.md` + `ENGINEERING_PRINCIPLES.md` + `PROJECT_MANAGEMENT_PRINCIPLES.md`.

## What this project is

`mcp-server-sharepoint` — a Model Context Protocol server that wraps Microsoft Graph's SharePoint **checkout / edit / checkin** model so AI coding agents can edit SharePoint documents without breaking version history, audit trail, or locking.

Full vision and tool surface in [docs/app-concept.md](docs/app-concept.md). Read it before changing anything that touches the public MCP tool surface.

## Project-specific docs

| Doc | Purpose |
|---|---|
| [docs/app-concept.md](docs/app-concept.md) | Vision, MVP scope, MCP tool surface, auth model, conflict/safety semantics, open tech-spike questions |
| [docs/testconcept.md](docs/testconcept.md) | Test-harness strategy for AI-assisted development |
| [docs/howto-oss.md](docs/howto-oss.md) | OSS-template setup notes inherited from the template; trim once they no longer apply |
| [README.md](README.md) | Quickstart for end users (install via `uvx`, configure `.mcp.json`) |
| [CONTRIBUTING.md](CONTRIBUTING.md) | Contribution flow |
| [SECURITY.md](SECURITY.md) | Vulnerability disclosure |
| [CHANGELOG.md](CHANGELOG.md) | Keep-a-changelog history |

## Tracker

**Authoritative tracker: GitHub Issues + the repo-bound GitHub Project** at <https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues>. See `ENGINEERING_PRINCIPLES.md` § 2.

- Labels:
  - `type:feat` / `type:fix` / `type:chore` / `type:docs` / `type:test`
  - `area:auth` / `area:tools` / `area:ci` / `area:packaging` / `area:docs`
  - `priority:p0` / `p1` / `p2`
  - `agent:claude` when an AI agent is the executor.
- Issue body convention: **Context** · **Acceptance criteria** (checkbox list) · **Out of scope** · **Links**.
- Milestones map to releases: `v0.1.0 — MVP`, `v0.2.0`, etc.

The legacy `docs/todo.md` is a frozen artefact from the OSS template; do not extend it. New work goes into Issues.

## Tech stack

- **Python 3.11+**, packaged for `uvx` / `pipx` install.
- **Raw `httpx`** against Microsoft Graph (the `msgraph-sdk-python` vs `httpx` spike resolved in favour of raw `httpx` — ~92× smaller footprint, no fight with the keyring-owned auth contract; see [docs/spikes/2026-05-06-graph-sdk-vs-httpx.md](docs/spikes/2026-05-06-graph-sdk-vs-httpx.md)).
- **MCP Python SDK** (FastMCP) for the protocol layer.
- **Auth**: OAuth 2.0 Device Code flow; three-tier token store (keyring → plain file mode 0600 → encrypted file via `SP_TOKEN_PASSPHRASE`), auto-detected, override with `SP_TOKEN_STORE`.
- **Tests**: pytest + a dedicated SharePoint test tenant. CI runs in GitHub Actions; secrets injected from repo secrets.
- **Lint/format**: ruff, mypy.

## Licence + attribution (this project)

This repo is **OSS**, dual-licensed. The proprietary-template SPDX variant used by sister repos does **not** apply here. Per `ENGINEERING_PRINCIPLES.md` §§ 11–12:

- **Licence**: dual-licensed **MIT OR Apache-2.0** — see [LICENSE-MIT](LICENSE-MIT), [LICENSE-APACHE](LICENSE-APACHE).
- **Copyright holder**: XMV Solutions GmbH.
- **SPDX licence identifier** for every new source file: `MIT OR Apache-2.0`.

The header forms (per language comment style), the `git config user.name` / `user.email` contributor rule, and the "never list an AI as a contributor / never add AI commit trailers" rules are all in `ENGINEERING_PRINCIPLES.md` §§ 11–12 — this repo follows them with the `MIT OR Apache-2.0` identifier above.

## Project-specific overrides of the engineering baseline

- **PR workflow already triggered (per § 13).** As soon as the package is published to PyPI / installable via `uvx`, the project has external users. Treat `main` as deployable trunk from that moment: feature branches + PRs, branch protection on `main`, CI green required for merge. Until the first published release, direct commits to `main` are acceptable for chores and docs.
- **Test environment (per § 5).** A dedicated SharePoint test site is required for integration/harness tests. Credentials live in GitHub Actions secrets for CI and in a developer-local `.env` (git-ignored) for iterative work. The tenant/site is documented in `docs/testconcept.md`.
- **No proprietary headers (per §§ 11–12).** Every header uses `SPDX-License-Identifier: MIT OR Apache-2.0`.
- **Harness token renewal (per § 5).** Microsoft refresh tokens rotate every ~60–90 days, so the `SHAREPOINT_HARNESS_TOKEN_JSON` repo secret is a recurring monthly maintenance chore that must be refreshed before CI's harness job starts failing. Run `./scripts/renew-harness-token.sh` — single-command Device Code login + smoke test + `gh secret set`. See [docs/testconcept.md § Renewing the harness token](docs/testconcept.md).

## Environments + URLs

- **GitHub**: <https://github.com/XMV-Solutions-GmbH/sharepoint-mcp>
- **Harness sandbox site**: `sharepoint-mcp-harness` in the XMV Solutions tenant — <https://xmvsolutions.sharepoint.com/sites/sharepoint-mcp-harness>.
- **Harness test user**: `d.koller@xmv.de` — real M365 user, **Edit** permission on the harness site only, no admin roles, no other tenant access. v0.1 uses delegated user auth (no service principal) so audit-log entries stay attributed to a real human; a leaked harness token is bounded to that user on that site.
- **CI harness secret**: `SHAREPOINT_HARNESS_TOKEN_JSON` (base64-encoded `~/.cache/sharepoint-mcp/harness/token.json`).
- **Default Entra app** (multi-tenant public client, owned by XMV): `client_id` `cb7cf68d-90d5-4841-90a7-de3a40be280b`, display name `mcp-server-sharepoint`, audience `AzureADMultipleOrgs`. Overridable via `SP_CLIENT_ID` / `SP_TENANT_ID`; zero-config when unset.
- There is no "staging" cluster environment — the harness sandbox **is** the staging-equivalent. The next step beyond harness is end-user PyPI install.

## MCP-specific design invariants

These are project-defining behaviours of the MCP server itself. They bind every existing tool and every future one; violating them is a release blocker. Full rationale in [docs/app-concept.md](docs/app-concept.md).

### 1. How tools are presented — nomenclature is load-bearing

Tool names follow `sp_<category>_<noun>_<verb>`. The first segment after `sp_` is always one of six categories — `{auth, site, drive, list, share, search}` — encoding which SharePoint entity the tool operates on, so an LLM picks the right tool from the **name alone** without reading every docstring (decisive when the catalog is 30+ tools and context budget is limited). Tool groups are filterable at startup via `SP_TOOL_GROUPS` (comma-separated subset of the six; `auth` always registered; unknown names = loud non-zero exit). On startup the server emits one stderr banner advertising version + active groups + writes flag, e.g. `mcp-server-sharepoint 0.7.0 — groups=[drive,list,site,search,share,auth] writes=true`.

### 2. How code/auth is emitted on sign-in — Device Code flow

Auth is **OAuth 2.0 Device Code Flow** for headless Linux. The `sp_auth_begin(profile)` tool starts the flow and returns the **user code + verification URL** to the calling MCP client, which displays them to the human; `sp_auth_status(profile)` reports `{signed_in | pending | none}`. The server polls Microsoft Identity until consent is granted, then caches access + refresh tokens locally via the three-tier token store. Multi-profile via `SP_PROFILE`; tools then appear as `mcp__sharepoint-<profile>__sp_*`. Every action is attributed to the signed-in human in SharePoint's audit log.

### 3. How content is exchanged — no base64, ever (local filesystem paths)

Tools **never** accept base64-encoded content as a parameter and **never** return it as a value. Production use showed even Claude-class models silently lose `=` padding when content round-trips through JSON — the failure is silent corruption, not a clean error. Binary content is exchanged via **local filesystem paths**: download tools (`sp_drive_file_read`) write to a temp path and return it; upload tools read from a path the LLM provides. The LLM uses its own filesystem tools on that path. There is no exception (the removed `sp_download_binary` base64 tool was replaced by this temp-file pattern). Recursive parent-folder creation is uniform across every write tool.

### 4. MCP-server test-harness — real Graph against a real sandbox

The harness layer (`tests/harness/`) is the § 5 **gate**: no v0.1 feature ticket lands without a corresponding harness test or a documented justification. It runs the real code against the **real** Microsoft Graph + the real `sharepoint-mcp-harness` SharePoint sandbox over real Device Code auth as the least-privilege user `d.koller@xmv.de`. Mock-shape discipline: mocks must match **captured** real-server response shapes (capture via `curl`/`httpx` against the sandbox first; captured shape wins over docs). Run locally on each iteration and in CI on every PR via `SHAREPOINT_HARNESS_TOKEN_JSON`; `./tests/run_tests.sh harness` / `all` drive it.

### 5. Behavioural / MCP-install test — an LLM with only the tool suite

The behavioural harness (`tests/harness/behavioural/`) verifies what the pytest harness cannot: that an **LLM-driven agent that knows only the MCP tool suite** (not the dev environment or operator briefing) picks the right tools in the right order against the full catalog. It spawns a cloud Claude agent, points it at this MCP server with the harness sandbox's refresh token, and gives it scripted real-world tasks (e.g. create folder trees + upload + rename/move; bulk metadata edits; a multi-step directory reorg). Each run is scored on tool-selection accuracy (did it pick `sp_drive_file_move` rather than `upload`+`delete`, which loses version history?), step count, successful end-state match, and confusion incidents. This is the layer that catches "agent picked `sp_upload_new_file` when the task was to edit an existing file" — the exact failure that motivated the v0.7.0 nomenclature restructuring. Run before every breaking release.

## Glossary

- **Checkout / checkin** — SharePoint's lock-edit-commit model exposed by Microsoft Graph; the core thing this MCP wraps so version history, audit trail, and locking survive AI edits.
- **`sp_<category>_<noun>_<verb>`** — the tool-naming scheme; category ∈ `{auth, site, drive, list, share, search}`.
- **Harness sandbox** — the dedicated `sharepoint-mcp-harness` SharePoint site used for live harness tests; the staging-equivalent for this project.
- **Behavioural harness** — the cloud-agent test layer that scores whether an LLM with only the tool suite drives the MCP correctly.
- **Device Code flow** — the interactive OAuth path for headless Linux; surfaces a user code + verification URL to the human.
- **`SP_TOOL_GROUPS` / `SP_ALLOW_WRITES` / `SP_PROFILE`** — startup env vars controlling, respectively, which tool categories register, whether mutating tools register, and token-cache/working-dir namespacing.
