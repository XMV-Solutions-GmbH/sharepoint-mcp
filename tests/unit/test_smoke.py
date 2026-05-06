# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Smoke tests — verify the package imports and basic wiring works."""

from __future__ import annotations


def test_package_importable() -> None:
    import sharepoint_mcp

    assert sharepoint_mcp.__version__


def test_cli_importable() -> None:
    from sharepoint_mcp import cli

    assert callable(cli.main)


def test_server_module_importable() -> None:
    from sharepoint_mcp import server

    assert callable(server.run)
    assert server.mcp.name == "sharepoint-mcp"


def test_cli_version_exits_cleanly() -> None:
    import pytest

    from sharepoint_mcp.cli import main

    with pytest.raises(SystemExit) as excinfo:
        main(["--version"])
    assert excinfo.value.code == 0
