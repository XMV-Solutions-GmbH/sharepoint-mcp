#!/usr/bin/env bash
# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
#
# Run the behavioural cloud-agent harness for a single scenario.
#
# Architecture: see tests/harness/behavioural/runner.py module docstring.
# Spec:         see docs/app-concept.md § Testability § Behavioural harness.
#
# Usage:
#   scripts/run-behavioural-harness.sh <scenario-name>
#
# Example:
#   scripts/run-behavioural-harness.sh anqer-reorg
#
# Prerequisites (the script aborts with a clear message if any are missing):
#   - `claude` CLI on $PATH (Claude Code installed, logged in).
#   - `uv` on $PATH.
#   - A harness profile token usable by the local server, populated via
#     `uv run mcp-server-sharepoint login --profile harness` at least once.
#   - `SP_ALLOW_WRITES=true` will be injected for the agent's MCP env.
#
# Cost note: this consumes Claude API tokens (your account) AND mutates
# the real harness SharePoint sandbox. Don't run it in tight loops.

set -euo pipefail

SCENARIO="${1:-}"
if [[ -z "$SCENARIO" ]]; then
  echo "usage: $0 <scenario-name>" >&2
  echo "available scenarios:" >&2
  ls "$(dirname "$0")/../tests/harness/behavioural/scenarios/"*.md \
    2>/dev/null | xargs -n1 basename | sed 's/\.md$//' | sed 's/^/  - /' >&2
  exit 2
fi

REPO_ROOT="$(cd "$(dirname "$0")/.." && pwd)"

# Prerequisites
for bin in claude uv; do
  if ! command -v "$bin" >/dev/null 2>&1; then
    echo "ERROR: required binary not on PATH: $bin" >&2
    exit 2
  fi
done

# Forward to the Python runner. The runner does the real work (seed,
# spawn claude, parse stream-json, verify final state, cleanup) so the
# orchestration logic stays unit-testable.
cd "$REPO_ROOT"
exec uv run python -m tests.harness.behavioural.runner \
  --scenario "$SCENARIO" \
  --repo-root "$REPO_ROOT" \
  --profile "${SP_HARNESS_PROFILE:-harness}"
