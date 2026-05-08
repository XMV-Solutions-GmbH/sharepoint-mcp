<!-- SPDX-License-Identifier: MIT OR Apache-2.0 -->
# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

No entries.

## [v0.4.0] — 2026-05-08

### Added — integrated MCP-tool login flow

- **`sp_login_begin(profile?, force?)`** — non-blocking. Initiates Microsoft Identity Device Code login as an MCP tool; returns within ~1s with `user_code` + `verification_url`. Background polling task writes the token on success. Idempotent unless `force=True`. (#75)
- **`sp_login_status(profile?)`** — three states the agent can act on directly: `signed_in` (valid token on disk, regardless of how it got there), `pending` (Device Code flow in progress), `none` (agent should call `sp_login_begin`). Critically: `signed_in` is determined by an active probe of the token cache, so a user who logged in via CLI days ago shows up correctly. (#76)
- Both tools' MCP descriptions include UX guidance for relaying the verification URL + user code to the user — code FIRST in its own code block (no labels), URL SECOND as a plain auto-link below. Optimises mobile copy → click → paste workflow.
- README has a new "Login from an MCP client" section showing the two-call agent pattern. CLI `login` / `logout` remain documented as the manual fallback path.

### Switched

- The four `sharepoint_mcp/auth/*.py` modules are now thin shims over [`mcp-microsoft-graph-auth`](https://pypi.org/project/mcp-microsoft-graph-auth/) (>=0.1.1). Backend implementations (Device Code primitives, TokenStore backends, service-principal client-credentials grant, `LoginSessionRegistry`) live in the shared library now. SharePoint-specific defaults (multi-tenant client_id, scopes, env-var conventions) stay here. (#74)
- `EncryptedFileTokenStore` now raises `NoUsableTokenStoreError` at construction time when `SP_TOKEN_PASSPHRASE` is empty (eager validation), rather than on first use — catches misconfigured CI faster.

### Limitations

- Pending login sessions live in the MCP server process. A server restart mid-flow drops them; the agent must call `sp_login_begin` again. Documented in README under "Login from an MCP client".
- Cross-process file lock on the token cache (sharepoint-mcp issue #77) is **deferred to a follow-up**. Concurrent CLI + tool-flow login on the same profile is not actively coordinated; in practice each path uses an atomic write so one wins cleanly, but the typed `concurrent_login_attempt` error path is a future improvement.

## [v0.3.0] — 2026-05-07

### Added — broader SharePoint Graph coverage

- **`sp_sites(query?)`, `sp_subsites(parent_site_url)`, `sp_followed_sites()`** — site discovery (#49). Find SharePoint sites without hardcoded URLs.
- **`sp_drives(site_url)`** + transparent multi-library support across every read/write tool — Site Assets, Style Library, custom libraries (#48).
- **`sp_trash_list(site_url)`** — list items in the SharePoint site recycle bin (#50, list-only). Restore deferred to follow-up #64 — Microsoft Graph doesn't currently expose a `restore` action at site scope.
- **`sp_lists(site_url)`, `sp_list_columns`, `sp_list_items`, `sp_get_item`** (read) and **`sp_create_item`, `sp_update_item`, `sp_delete_item`** (write) — full CRUD on SharePoint Lists (#44).
- **`sp_permissions(url)`** — read-only inspection of who has access to a file/folder/site (#46). Normalised across user / group / sharing-link / siteUser / application principals.
- **`sp_share_list(url)`** (read), **`sp_share_create(url, type, scope, expires?, password?)`** + **`sp_share_revoke(url, link_id)`** (write) — sharing-link management with conservative defaults (`view` + `organization`); explicit opt-in required for `anonymous` (#47).
- **`sp_pages_list(site_url)`, `sp_page_read(page_url)`, `sp_page_update(page_url, title?, description?, thumbnail_web_url?)`** — modern SharePoint Pages with raw canvasLayout JSON (#45). Canvas-layout edits intentionally deferred — needs a clearer agent UX.
- **`sp_changes(scope_url, since?)`** — Microsoft Graph delta queries for incremental change tracking (#51). Cursor-based; opaque to callers.

### Changed

- `CheckoutRegistry.add` / `.remove` already serialise via `threading.Lock`; multi-library support added a `resolve_drive_item_full` helper that transparently retries against the matched library on a default-drive 404. One extra `/drives` lookup per fallback; happy path unchanged.
- `sp_permissions` grantee output gains a `link_web_url` field (used by `sp_share_list` to surface the actual share URL on existing links). Backward-compatible additive.

### Documentation

- README now covers all 24 read tools + 11 write tools across drives, lists, sites, pages, sharing, permissions, recycle-bin, and delta. Roadmap updated.

## [v0.2.0] — 2026-05-07

### Added

- **`sp_publish(local_path, target_folder_url, name=None)`** (write) — upload a brand-new local file to a SharePoint folder for the "draft + promote" workflow. Refuses to overwrite — use `sp_open` + `sp_save` for edits with proper audit history. (#31)
- **`sp_history(url, limit=20)`** (read) — list a SharePoint file's version history (newest first). (#42)
- **`sp_get_version(url, version_id)`** (read) — download a specific historical version to a local temp file. (#42)
- **`sp_status(verify=False)`** (read) — `verify=True` queries SharePoint to confirm server-side lock state and surfaces `server_locked` + `lock_holder` per registry entry. Default unchanged. (#37)
- **`sp_open_many(urls)` / `sp_save_many(operations)`** (write) — bulk variants with concurrency cap of 4, Retry-After-aware backoff on 429/503, per-item failure isolation, input-order preservation. (#41)
- **Resumable uploads** — `sp_save` automatically switches to Microsoft Graph's resumable upload session for files larger than `SP_CHUNKED_UPLOAD_THRESHOLD_MB` (default 100 MB). 5 MiB chunks with retry on transient errors. (#38)
- **Service-principal / client-credentials auth** — activated by `SP_AUTH_MODE=service-principal` or auto-detected when `SP_CLIENT_SECRET` is set. For unattended automation; trades human attribution in the audit log for the ability to run unattended. (#40)

### Changed

- `CheckoutRegistry.add` / `.remove` now hold a process-wide `threading.Lock` so bulk operations don't race when adding multiple entries concurrently.
- `Development Status` classifier moved from "2 - Pre-Alpha" to "3 - Alpha".

### Documentation

- README updated with bulk operations, large-file handling, service-principal mode docs, and the new tool entries.

## [v0.1.0] — 2026-05-07

### Added

- **Authentication**: OAuth 2.0 Device Code Flow against Microsoft Identity, silent refresh-token loop, three-tier token persistence (OS keyring / plain file mode 0600 / passphrase-encrypted file). Multi-profile support via `SP_PROFILE`. BYO Entra app registration via `SP_CLIENT_ID` / `SP_TENANT_ID` for tenants with strict app-allowlisting.
- **Read tools** (always registered): `sp_search`, `sp_list`, `sp_read`, `sp_status`.
- **Write tools** (registered when `SP_ALLOW_WRITES=true`): `sp_open`, `sp_save`, `sp_release`. Mandatory non-empty audit comment on every save. ETag-based stale-write detection. Per-process checkout registry persisted across crashes.
- **MCP tool annotations** correctly applied to every tool (`readOnlyHint`, `destructiveHint`, etc.) so MCP clients render appropriate permission prompts.
- **CLI**: `mcp-server-sharepoint login [--profile NAME]`, `mcp-server-sharepoint logout [--profile NAME]`, `mcp-server-sharepoint` (default — start the MCP server on stdio).
- **Test harness**: three layers (unit / integration / harness) with the harness layer running against a real SharePoint sandbox in CI via the `SHAREPOINT_HARNESS_TOKEN_JSON` repo secret.
- **Documentation**: README with quickstart + security model + troubleshooting; engineering principles + project conventions; testconcept; spike decisions for major design choices.

### Project layout

- Python package skeleton: `pyproject.toml` (hatchling, dual-license MIT OR Apache-2.0), `src/sharepoint_mcp/` layout.
- Tooling: ruff (lint + format, line-length 100, target py311), mypy (strict), pytest 8+ with auto-markers per layer.
- CI: lint + test + harness jobs in GitHub Actions on every push to `main` and intra-repo PR.

### Initial project structure

- Engineering principles, app concept, license (dual MIT OR Apache-2.0), contributing guidelines, security policy.

[Unreleased]: https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/compare/v0.1.0...HEAD

[v0.1.0]: https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/releases/tag/v0.1.0
