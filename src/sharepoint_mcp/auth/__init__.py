# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Authentication layer.

OAuth 2.0 Device Code flow against Microsoft Identity, with token
persistence via two interchangeable backends (OS keyring + encrypted
file). Token refresh is silent; first-run device-code login surfaces
to the MCP client via a structured error response.

Public API lands here once the components stabilise. Until then, see
the individual submodules.
"""
