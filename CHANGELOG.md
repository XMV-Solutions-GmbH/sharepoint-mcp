<!-- SPDX-License-Identifier: MIT OR Apache-2.0 -->
# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed (breaking)

- **`sp_list` renamed to `sp_list_folder`**. The previous name was one
  letter away from `sp_lists` (SharePoint Lists CRUD) and agents under
  context pressure routinely confused the two — they have entirely
  different semantics (`sp_list_folder` lists drive items in a
  document-library folder; `sp_lists` lists SharePoint List collections
  like Issue Trackers, Tasks, Custom Lists). Migration: rename the tool
  call. Behaviour, signature, and return shape are unchanged.

### Removed

- **`sp_page_update`** — removed. Microsoft Graph's modern Pages API exposes
  reads of the full canvas layout but the metadata-only write surface
  (`title` / `description` / `thumbnail`) was a half-tool — agents reached
  for it expecting full edits and routinely tried to use it to change page
  content, which it can't do. Canvas-layout writes need a clearer agent UX
  before they're safe; until then, modern Pages have to be edited via the
  SharePoint web UI. `sp_pages_list` and `sp_page_read` remain.

- **`sp_subsites`** — removed. Sub-sites are a legacy SharePoint construct;
  modern tenants use flat site architectures with Hub Sites. The tool added
  a third navigation entry point on top of `sp_sites` and `sp_followed_sites`
  for a rarely-needed traversal pattern. Migration: list all visible sites
  with `sp_sites(query)` (which already includes sub-sites in its results)
  or read site metadata directly if you need the parent/child relationship.

- **`sp_upload_new_file`** — removed entirely (closes
  [#99](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues/99)).
  The b64-inline API was an attractive nuisance: agents reached for it even for
  files already on disk, inlining file content as base64 into the tool call
  (wastes context tokens, risks transcription corruption on long strings).
  Migration: write content to a local file (or a `tempfile`) and call
  `sp_publish(local_path, target_folder_url)`. `sp_publish` reads the file
  directly, uses zero tokens of file content in agent context, and supports
  files of any size.

### Added

- **`sp_delete_file(site_url, path)`** — soft-delete a file or folder to the
  site recycle bin via `DELETE /drives/{id}/items/{id}` (Graph returns 204).
  The item is recoverable from the SharePoint recycle bin. Returns
  `{deleted: true, path}`. Gated by `SP_ALLOW_WRITES=true`. Implements
  [#92](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues/92).

- **`sp_move_file(site_url, source_path, destination_path)`** — move or rename
  a file/folder via `PATCH /drives/{id}/items/{id}` with a new
  `parentReference` and/or `name`. The destination is the full path after the
  move (last segment = new name; preceding segments = existing destination
  folder). Supports cross-folder move, in-place rename, and combined
  move-and-rename in a single call. Returns `{moved: true, source, destination,
  web_url}`. Gated by `SP_ALLOW_WRITES=true`. Implements
  [#95](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues/95).

- **`sp_copy_file(site_url, source_path, destination_path)`** — copy a file to
  a new path via `POST /drives/{id}/items/{id}/copy`. The Graph copy endpoint
  is asynchronous (returns 202 Accepted + `Location` header); this tool polls
  the operation-status URL until the copy completes or a configurable timeout
  (default 60 s) elapses. Also handles synchronous 200/201 responses (test
  tenants) and 303 CDN-redirect responses (both on the initial POST and during
  polling). Returns `{copied: true, source, destination, web_url}`. Gated by
  `SP_ALLOW_WRITES=true`. Implements
  [#96](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues/96).

## [v0.6.1] — 2026-05-12

Bug-fix release. No new features, no breaking changes.

### Fixed

- **`sp_create_folder`: missing `:` in nested-path Graph URL** (closes [#90](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues/90)). The endpoint for creating a folder under an existing parent used `drive/root:/{parent}/children` instead of the correct `drive/root:/{parent}:/children`. Microsoft Graph's path-based addressing syntax requires the trailing `:` to close the path context before appending the relationship segment. The root-level case (`drive/root/children`) was unaffected. The incorrect URL was accepted by the mock layer in unit tests (mocks match whatever URL the code produces) — the bug was only visible against the real Graph API.

### Engineering

- Added `tests/harness/test_create_folder_and_upload.py` — 6 harness tests against the real SharePoint sandbox (absent in v0.6.0). These tests would have caught the URL bug immediately. Going forward, harness tests are part of the implementation of any tool that calls Graph, not a follow-up. Fixed the 3 corresponding unit-test mock URLs (`root:/{parent}/children` → `root:/{parent}:/children`).

## [v0.6.0] — 2026-05-12

Two new write tools that complete the creation surface — previously the
checkout/checkin lifecycle (`sp_open`/`sp_save`) only worked on *existing* items.

### Added

- **`sp_create_folder(site_url, path)`** — create a folder hierarchy in the
  site's default document library. `path` is relative to the document library
  root (e.g. `"2026/Q2/Reports"`). A leading `"Shared Documents/"` prefix is
  stripped for convenience. Intermediate folders that don't yet exist are
  created in one call (recursive mkdir semantics). Existing folders are skipped
  without error (idempotent). Returns `{created, already_existed, web_url}`.
  Gated by `SP_ALLOW_WRITES=true`. Implements [#86](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues/86).

- **`sp_upload_new_file(site_url, path, content)`** — upload a new file from
  inline base64-encoded content. `path` includes the filename and is relative
  to the document library root. Content-Type is inferred from the file extension
  with `application/octet-stream` as fallback. Refuses if the target already
  exists (raises `FileAlreadyExistsError` directing the user to `sp_open`/`sp_save`
  instead). Capped at 4 MB decoded; for larger files use `sp_publish`.
  Returns `{item_id, etag, web_url, size}`. Gated by `SP_ALLOW_WRITES=true`.
  Implements [#87](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues/87).

### Design note

These two tools are complementary to, not in conflict with, the checkout/checkin
lifecycle. Checkout/checkin (`sp_open`/`sp_save`) prevents concurrent overwrite on
*existing* items. Creation (`sp_create_folder`, `sp_upload_new_file`, `sp_publish`)
has no prior version to conflict with — no checkout needed by definition. Typical
agent workflow for bootstrapping new content: `sp_create_folder` → `sp_upload_new_file`
(or `sp_publish` for files already on disk). Typical workflow for editing existing
content: `sp_open` → edit → `sp_save`.

### Engineering

- 560 unit tests (was 529; +31). New test files:
  `tests/unit/tools/test_create_folder.py` (15 tests) and
  `tests/unit/tools/test_upload_new_file.py` (16 tests). Coverage: happy path,
  deep nested creation, partial pre-existence, full idempotency, name collision
  with file, library-prefix stripping, path normalisation, content-type inference,
  base64 decode validation, size limit, bearer propagation, all input-validation
  branches.

## [v0.5.0] — 2026-05-12

**Breaking change** to the consent-env-var contract — same pattern as `outlook-mcp` v0.4.0. Operators upgrading from v0.4.x must update their `.mcp.json` to set `SP_ALLOW_WRITES` to exactly `"true"` or `"false"`; legacy truthy values (`1`, `yes`, `on`) and unset / empty are now rejected at startup. Plus the OAuth consent screen now reflects the operator's actual decision — with `SP_ALLOW_WRITES=false` the prompt requests only `Files.Read.All` + `Sites.Read.All` instead of the ReadWrite variants.

### Changed (breaking)

- **`SP_ALLOW_WRITES` must be set to exactly `"true"` or `"false"`** (case-insensitive, trimmed). Any other value — including unset / empty / legacy `1`/`yes`/`on` — causes the server (and the CLI `login` subcommand) to refuse to start with a formatted onboarding-help message printed to stderr. The motivation matches the outlook-mcp issue #37 user-side rationale: operators silently landing in read-only mode without realising writes were a separately-opt-in feature was the dominant onboarding failure mode in v0.4.x.
- **OAuth scopes now respect the consent decision.** With `SP_ALLOW_WRITES=false`, `resolve_scopes()` requests `Files.Read.All` + `Sites.Read.All` (read-only variants) instead of the ReadWrite ones. With `=true`, the ReadWrite variants replace them. The consent screen on a fresh login reads accordingly. Previously the OAuth scopes were ALWAYS ReadWrite regardless of the `SP_ALLOW_WRITES` decision, which was inconsistent with the "read-only by default" tool-surface story.
- **Server start is no longer silently read-only** when consent is unset. Previously the server fell through to read-only mode with an INFO log; operators commonly missed the log and assumed writes were broken. The new error message is itself the documentation.

### Added

- **`sharepoint_mcp.auth.flow.SharepointConsentNotConfiguredError`** — new exception class raised by the strict consent parser. Re-exported from `auth.flow.__all__` so downstream tooling can catch it.
- **`sharepoint_mcp.auth.flow.validate_consent_config()`** — returns `writes_enabled` (True/False) or raises. Single source of truth; called from `_build_server()` at module import and from `cli.main()` before the login flow.
- **`sharepoint_mcp.auth.flow.resolve_scopes()`** — runtime-resolving scope tuple. Returns the read-only base scopes when `SP_ALLOW_WRITES=false`, ReadWrite variants when `=true`. Resolved at call time so `monkeypatch.setenv` flips behaviour without re-imports.

### Engineering

- 529 unit tests (was 506; +23 new). New tests cover: strict env-var parser, scope resolution for both flags, server-build refusal on unset / legacy-truthy values, CLI gating, scope-tuple shape against backwards-compat aliases.
- New harness test (`tests/harness/test_consent_gate.py`) — 6 end-to-end checks against the real harness profile + token-store. Confirms scope resolution doesn't break the cached harness token, `_build_server()` exposes / withholds write tools per env-var decision, and the error message contains the actionable strings (`SP_ALLOW_WRITES`, `"true"`, `"false"`, `.mcp.json`).

### Migration from v0.4.x

Add the explicit decision to your `.mcp.json` env section:

```jsonc
{
  "mcpServers": {
    "sharepoint": {
      "command": "uvx",
      "args": ["mcp-server-sharepoint"],
      "env": {
        "SP_ALLOW_WRITES": "false"   // read-only mode (no checkout/save/share/edit tools)
        // or
        // "SP_ALLOW_WRITES": "true"   // enables sp_open, sp_save, sp_release, sp_publish, sp_*_item, sp_share_*, sp_page_update
      }
    }
  }
}
```

If you were already setting `SP_ALLOW_WRITES=true` in v0.4.x, no change is needed (that exact value remains the strict form). If you relied on legacy `1`/`yes`/`on`, change to `true`.

**OAuth re-consent:** the first time the server starts under v0.5 with `SP_ALLOW_WRITES=false`, the next login will request narrower scopes (`Files.Read.All` instead of `Files.ReadWrite.All`). Existing cached tokens from v0.4 keep working — Graph accepts the broader-scope token even when the client requests only narrower scopes on the next refresh.

## [v0.4.1] — 2026-05-11

Bug-fix release. No new features, no breaking changes.

### Fixed

- **`sp_list` / `sp_read` now work on URLs with localized document library names** (e.g. German `Freigegebene Dokumente`, Italian `Documenti condivisi`, Spanish `Documentos compartidos`). Previously these URLs round-tripped from `sp_search` results but failed with 404 when passed back into `sp_list` or `sp_read`, because the path-based Graph resolver only knew the English library name. `resolve_drive_item_full` gains a new first-fallback step: on primary 404, **strip the first path segment and retry against the default drive's root** before consulting the library-name list. This handles the "URL prefixed with the localized default-library display name" case (the dominant cause of #79) without locale enumeration. The existing library-name-search fallback continues to handle non-default custom libraries (`SiteAssets`, etc.). Closes [#79](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues/79).

### Added (internal)

- **`sharepoint_mcp.tools._common.resolve_drive_item_by_share_url(client, web_url, headers)`** — new helper that wraps Microsoft Graph's `/shares/{u!base64}/driveItem` endpoint. Not used by `sp_list` / `sp_read` directly (that endpoint requires sharing-link access, which not all service-principal-style auth configurations have — empirically verified by the harness returning 403 on bare site-membership URLs), but exposed for future tools where the caller has a real shared-link URL. Encoding follows the Graph spec: urlsafe base64 of the UTF-8 URL bytes, stripped of `=` padding, prefixed with `u!`.

### Engineering

- 520 unit tests (was 506; +14 new — 4 for the `resolve_drive_item_by_share_url` helper, 2 for the new strip-first-segment fallback path in `resolve_drive_item_full`, plus regression tests in `test_read.py` and `test_list_folder.py` covering the round-trip on German URLs). Existing tests that 404 the primary path now mock the strip-retry as 404 too so the fall-through to library-search is exercised.

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
