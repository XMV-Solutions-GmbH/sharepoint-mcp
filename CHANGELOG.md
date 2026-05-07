<!-- SPDX-License-Identifier: MIT OR Apache-2.0 -->
# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

No entries — track the next-version work in the v0.3 ticket queue.

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
