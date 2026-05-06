<!-- SPDX-License-Identifier: MIT OR Apache-2.0 -->

# sharepoint-mcp

[![Licence](https://img.shields.io/badge/licence-MIT%20OR%20Apache--2.0-blue.svg)](LICENSE)
[![CI](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/actions/workflows/ci.yml)
[![contributions welcome](https://img.shields.io/badge/contributions-welcome-brightgreen.svg?style=flat)](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues)

A **Model Context Protocol** server for SharePoint document libraries. Lets AI coding agents read and edit files on SharePoint **without breaking version history, audit trail, or locking semantics** — by wrapping Microsoft Graph's native checkout / edit / checkin model as MCP tools.

> **Status:** pre-alpha. Concept frozen, MVP in planning. See [docs/app-concept.md](docs/app-concept.md) and the [issue tracker](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues).

---

## Why

The standard alternatives — `rclone`, WebDAV mounts, the Anthropic-hosted M365 MCP — either skip SharePoint's checkout/checkin model entirely or expose it only for search and read. That's not acceptable for documents that live under retention or audit constraints (ISMS records, controlled procedures, contract templates). `sharepoint-mcp` keeps the audit trail intact: every edit goes through an explicit `open` → `save` (or `release`) cycle, attributed to the signed-in user, with a commit message and an honest version bump.

Full rationale, tool surface, auth model, and conflict semantics in [docs/app-concept.md](docs/app-concept.md).

---

## MCP tools (planned)

```text
sp_search    sp_list      sp_read      sp_open
sp_save      sp_release   sp_status    sp_history    sp_get_version
```

Each maps to one or two Microsoft Graph calls. No clever caching beyond what Graph provides.

---

## Quickstart (planned, post-MVP)

Once the first release is on PyPI:

```bash
# Install via uvx (no global Python install required)
uvx sharepoint-mcp --help
```

Wire it into your MCP client (e.g. Claude Code) via `.mcp.json`:

```json
{
  "mcpServers": {
    "sharepoint": {
      "command": "uvx",
      "args": ["sharepoint-mcp"],
      "env": {
        "SP_TENANT_ID": "<tenant-id>",
        "SP_CLIENT_ID": "<app-registration-id>",
        "SP_PROFILE": "default"
      }
    }
  }
}
```

First run prompts a Device Code login; subsequent runs use a cached refresh token.

---

## Documentation

| Document | Description |
| -------- | ----------- |
| [App Concept](docs/app-concept.md) | Vision, MVP scope, MCP tool surface, auth, conflict model |
| [Test Concept](docs/testconcept.md) | Test-harness strategy for AI-assisted development |
| [Engineering Principles](ENGINEERING_PRINCIPLES.md) | Project-agnostic baseline (language, status workflow, source control, licensing) |
| [Project Conventions](CLAUDE.md) | sharepoint-mcp-specific overrides on top of the principles |
| [Contributing](CONTRIBUTING.md) | How to contribute |
| [Security Policy](SECURITY.md) | How to report vulnerabilities |

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
