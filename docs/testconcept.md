<!-- SPDX-License-Identifier: MIT OR Apache-2.0 -->

# Test Concept

This is the operationalised version of `ENGINEERING_PRINCIPLES.md` § 5 for `mcp-server-sharepoint`. Read the principles first; this document is the project-specific instantiation.

---

## Three test layers

| Layer | Where | What it verifies | External world | Speed |
|---|---|---|---|---|
| **Unit** | `tests/unit/` | Pure-function logic in isolation | All externals mocked (Microsoft Graph via `respx`, OS keyring via in-memory fake) | sub-second per test |
| **Integration** | `tests/integration/` | Cross-module wiring with boundary mocks | Boundary mocks at HTTP layer (`respx` against `graph.microsoft.com`) and at the keyring/token-store layer | <1 s per test |
| **Harness** | `tests/harness/` | Our code against the **real** Microsoft Graph + a real SharePoint sandbox | Real network, real Graph endpoints, real least-privilege user `d.koller@xmv.de` | seconds per test (network bound) |

The harness layer is the **gate** per `ENGINEERING_PRINCIPLES.md` § 5: no v0.1 feature ticket lands without a corresponding harness test or a documented justification for why one isn't possible.

---

## What runs where

- `./tests/run_tests.sh` (default = `unit + integration`) — runs in CI on every PR. No SharePoint credentials needed.
- `./tests/run_tests.sh harness` — requires `harness` profile token cache (run `uv run mcp-server-sharepoint login --profile harness` once). Runs from the developer machine and in CI via the `SHAREPOINT_HARNESS_TOKEN_JSON` repo secret. The token rotates every ~60-90 days; renew with `./scripts/renew-harness-token.sh` (see "Renewing the harness token" below).
- `./tests/run_tests.sh all` — unit + integration + harness in one shot.

The `tests/conftest.py` auto-marks tests by their parent directory so `pytest -m unit` / `-m integration` / `-m harness` filter correctly without each test having to apply the marker by hand.

---

## Harness sandbox

**Site**: `sharepoint-mcp-harness` in the XMV Solutions tenant. URL: <https://xmvsolutions.sharepoint.com/sites/sharepoint-mcp-harness>.

**Test user**: `d.koller@xmv.de` — a real M365 user with the **smallest license that includes SharePoint** (E5 Developer in this case). Member of the M365 group that backs the harness site, with **Edit** permission. **No** admin roles, no access to other XMV tenant resources beyond what membership in the harness group implies.

**Why a real user, not a service principal**: v0.1 only supports delegated user auth (no client-credentials flow). A leaked harness refresh token would let an attacker act as `d.koller@xmv.de` against the harness site only — the blast radius is bounded by the user's permissions.

**Seed data**: the harness Documents library contains synthetic test files (`README.md`, `policies/iso27001-control-A.5.1.md`, `drafts/onboarding-draft.md`) sufficient for sp_search / sp_list / sp_read / sp_open / sp_save / sp_release flows. Tests that mutate state clean up after themselves via the `_cleanup.py` fixture.

---

## Authentication for tests

**Local development**:

```bash
uv run mcp-server-sharepoint login --profile harness
```

Refresh token cached at `~/.cache/sharepoint-mcp/harness/token.json` (mode 0600) on the developer's machine. Survives reboots, lasts ~60–90 days until Microsoft expires the refresh token.

**CI**:

The harness CI job receives the cached token (refresh token + access token, base64-encoded) via the `SHAREPOINT_HARNESS_TOKEN_JSON` secret. The job materialises the token cache at the start, runs `./tests/run_tests.sh harness`, and discards the runner.

### Renewing the harness token

Microsoft Identity rotates refresh tokens every ~60-90 days, so the harness secret is a recurring monthly maintenance chore. The flow is automated via:

```bash
./scripts/renew-harness-token.sh
```

Walks through Microsoft's Device Code login for the `d.koller@xmv.de` harness account, runs a `/me` smoke test to confirm the new token is valid, base64-encodes the cached `~/.cache/sharepoint-mcp/harness/token.json`, and uploads it to the repo as `SHAREPOINT_HARNESS_TOKEN_JSON` via `gh secret set`. CI's next harness run picks it up automatically — no other manual steps.

There's no automatic refresh-the-secret mechanism. That would require either client-credentials (which we don't use here for compliance reasons — `sp_*` tools default to delegated user auth so audit-log entries stay attributed to a real human) or a long-lived service-principal seed (also off the table for the harness sandbox).

---

## Mock-shape validation against real Graph (the discipline)

**Mocks must match real-server response shapes.** Documentation alone is not enough — Microsoft's docs are mostly accurate but occasionally drift. Mocks based on docs can silently agree with code based on the same docs while production behaviour diverges.

**The validated workflow:**

1. **Capture real responses** before writing mocks for a new tool: run the tool's underlying Graph endpoint against the harness sandbox via `curl` or a one-off `httpx` call, save the response payload.
2. **Base the mocks on captured payloads.** If the captured shape diverges from what docs imply, the captured shape wins.
3. **Write code to handle the captured shape.**
4. **Harness test confirms** the live behaviour matches what unit tests assume.

**For v0.1 specifically:** initial mock drafts were inferred from Microsoft Graph documentation, then validated against captured responses on 2026-05-07. Findings: most response shapes were operationally correct (extra fields in real responses don't break parsing); one mismatch caught (`/shares` endpoint requires sharing-link, not site-membership — fixed by switching to site→drive lookup). The `sp_list` rewrite is the visible artefact of that validation pass.

**Where the mock-vs-reality gap is still untested**: error-path responses that are hard to provoke deterministically against the real server, namely:

- **412 Precondition Failed** on `PUT /content` (only triggered by another user changing the file between sp_open and sp_save) — currently mock-only.
- **423 Locked** on `POST /checkout` (only triggered by two parallel sessions of the same user) — currently mock-only.

These are documented as known gaps. Both follow well-defined HTTP status semantics (RFC 7232 / RFC 4918) and Microsoft's documented error shapes. If a real-world failure surfaces an unexpected response shape, capture it and tighten the mocks.

---

## Test naming

| Prefix / location | Purpose |
|---|---|
| `tests/unit/test_*.py` | Per-module unit tests, mocks at any layer needed |
| `tests/unit/tools/test_<tool>.py` | One file per `sp_*` tool |
| `tests/unit/test_server.py` | FastMCP registration logic (read-only-default gating) |
| `tests/integration/test_*.py` | Cross-module flows with boundary mocks |
| `tests/harness/test_<tool>.py` | One file per tool against real Graph |
| `tests/harness/test_write_lifecycle.py` | open → save / open → release end-to-end |
| `tests/harness/test_auth_smoke.py` | First-line auth chain (post-login proof of life) |

---

## What's NOT in scope of the test suite

- Real-world tenants other than the harness sandbox. We don't run automated tests against XMV's production data, customer tenants, or anything we don't fully own.
- Microsoft itself (Microsoft Graph uptime, scope-policy changes) — observed via harness failures, responded to with code/doc fixes, not by trying to mock around them.
- Cross-MCP-client compatibility (Claude Desktop, other MCP clients) — the protocol is the contract; we test against `mcp` Python SDK conformance, not specific clients.
- Performance / load testing — manual / future. v0.1 is correctness first.
