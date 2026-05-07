<!-- SPDX-License-Identifier: MIT OR Apache-2.0 -->
# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

*(no entries — most recent commits land under v0.1.0 once that release is cut)*

## [v0.1.0] — pending

### Added

- **Authentication**: OAuth 2.0 Device Code Flow against Microsoft Identity, silent refresh-token loop, three-tier token persistence (OS keyring / plain file mode 0600 / passphrase-encrypted file). Multi-profile support via `SP_PROFILE`. BYO Entra app registration via `SP_CLIENT_ID` / `SP_TENANT_ID` for tenants with strict app-allowlisting.
- **Read tools** (always registered): `sp_search`, `sp_list`, `sp_read`, `sp_status`.
- **Write tools** (registered when `SP_ALLOW_WRITES=true`): `sp_open`, `sp_save`, `sp_release`. Mandatory non-empty audit comment on every save. ETag-based stale-write detection. Per-process checkout registry persisted across crashes.
- **MCP tool annotations** correctly applied to every tool (`readOnlyHint`, `destructiveHint`, etc.) so MCP clients render appropriate permission prompts.
- **CLI**: `sharepoint-mcp login [--profile NAME]`, `sharepoint-mcp logout [--profile NAME]`, `sharepoint-mcp` (default — start the MCP server on stdio).
- **Test harness**: three layers (unit / integration / harness) with the harness layer running against a real SharePoint sandbox in CI via the `SHAREPOINT_HARNESS_TOKEN_JSON` repo secret.
- **Documentation**: README with quickstart + security model + troubleshooting; engineering principles + project conventions; testconcept; spike decisions for major design choices.

### Project layout

- Python package skeleton: `pyproject.toml` (hatchling, dual-license MIT OR Apache-2.0), `src/sharepoint_mcp/` layout.
- Tooling: ruff (lint + format, line-length 100, target py311), mypy (strict), pytest 8+ with auto-markers per layer.
- CI: lint + test + harness jobs in GitHub Actions on every push to `main` and intra-repo PR.

### Initial project structure

- Engineering principles, app concept, license (dual MIT OR Apache-2.0), contributing guidelines, security policy.

[Unreleased]: https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/compare/v0.1.0...HEAD
