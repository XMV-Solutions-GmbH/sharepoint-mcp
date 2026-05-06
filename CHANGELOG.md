<!-- SPDX-License-Identifier: MIT OR Apache-2.0 -->
# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Python package skeleton: `pyproject.toml` (hatchling, dual-license), `src/sharepoint_mcp/` layout with empty MCP server stub on top of `mcp.server.fastmcp.FastMCP`, argparse CLI shim with `--help` / `--version`.
- Three-layer test scaffolding (`tests/{unit,integration,harness}/`) with auto-marker conftest hook, `tests/run_tests.sh` dispatcher, four passing unit smoke tests.
- Tooling: ruff (lint + format, line-length 100, target py311), mypy (strict), pytest with markers per layer.
- CI updated: replace bats step with `uv sync --extra dev`, ruff check + format, mypy, pytest unit + integration default.
- Initial project structure (engineering principles, app concept, license, contributing).

[Unreleased]: https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/compare/v0.1.0...HEAD
