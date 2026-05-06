#!/usr/bin/env bash
# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
#
# Test runner for sharepoint-mcp. Dispatches to the three test layers
# defined in ENGINEERING_PRINCIPLES.md § 5.
#
# Usage:
#   ./tests/run_tests.sh             # default: unit + integration
#   ./tests/run_tests.sh unit        # only unit
#   ./tests/run_tests.sh integration # only integration (boundary mocks)
#   ./tests/run_tests.sh harness     # only harness (real Microsoft Graph)
#   ./tests/run_tests.sh all         # unit + integration + harness
#
# Harness tests require harness credentials installed locally (or in
# CI as the SHAREPOINT_HARNESS_REFRESH_TOKEN secret) — see
# docs/testconcept.md.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

cd "${REPO_ROOT}"

target="${1:-default}"

run_layer() {
    local layer="$1"
    local path="${SCRIPT_DIR}/${layer}"
    if [[ ! -d "${path}" ]]; then
        echo "ERROR: layer directory not found: ${path}" >&2
        return 1
    fi
    echo ">>> uv run pytest tests/${layer}"
    uv run pytest -m "${layer}" "${path}"
}

case "${target}" in
    unit)
        run_layer unit
        ;;
    integration)
        run_layer integration
        ;;
    harness)
        run_layer harness
        ;;
    all)
        run_layer unit
        run_layer integration
        run_layer harness
        ;;
    default)
        run_layer unit
        run_layer integration
        ;;
    *)
        echo "Unknown target: ${target}" >&2
        echo "Usage: $0 [unit|integration|harness|all|default]" >&2
        exit 2
        ;;
esac
