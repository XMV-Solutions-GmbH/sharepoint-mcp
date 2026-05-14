<!-- SPDX-License-Identifier: MIT OR Apache-2.0 -->
# Proposal: Run harness tests as a CI gate on every PR

- **Status:** Implemented
- **Authors:** `David Koller <dko-ek@xmv.de>`
- **Date drafted:** 2026-05-14
- **Date accepted:** 2026-05-14
- **Tracking issue:** n/a (decided during initial project setup)

## Context

`ENGINEERING_PRINCIPLES.md` § 5 names "whether to run harness tests in CI" as a project-specific trade-off. The costs and benefits are described there. This record documents the decision for `sharepoint-mcp`.

The project uses a real Microsoft SharePoint sandbox (`xmvsolutions.sharepoint.com/sites/sharepoint-mcp-harness`) as the harness environment. Every tool in the MCP server makes authenticated Microsoft Graph API calls. The AI agent does active development on this repo and holds harness credentials locally.

## Decision

Harness tests run in CI on every pull request as a required status check.

The CI pipeline has a dedicated `harness` job that authenticates via a cached OAuth token stored as a GitHub Actions secret, then runs `pytest tests/harness` against the real SharePoint sandbox. A PR cannot merge if the harness job fails.

## Alternatives considered

### Alternative A: Local-only harness (no CI harness job)

The agent runs harness tests locally during development; CI only covers unit + integration.

This was **not** chosen because:

- The primary motivation for this project is AI-assisted development. Bugs caused by wrong assumptions about Graph API behaviour (wrong URL shape, unexpected 303 response, 404 on a differently-named library) only surface against the real API. Without CI harness, those bugs can slip through to `main` if the agent forgets to run harness locally, or if the agent's local token is stale.
- Every PR to this project ships tool implementations that call Graph. There is no PR category where skipping real-API verification makes sense.
- The sandbox credentials are stable (long-lived refresh token, test-only account scoped to the sandbox site). The maintenance burden is low.

### Alternative B: Separate harness CI schedule (nightly, not per-PR)

Run harness on a schedule rather than per PR.

This was **not** chosen because it decouples the harness result from the specific commit it tests, which makes failures harder to attribute and allows broken commits to merge.

## Consequences

**Positive:**

- Every PR is verified against the real Graph API before merge. The 303-CDN-redirect bug (PR #97) was caught by the CI harness job, not by unit tests.
- Contributors without local harness credentials still get their PRs verified end-to-end.

**Negative:**

- CI run time is longer (~2 min for the harness job vs. ~30 s for unit tests).
- PRs can fail due to Graph API availability, not project-code bugs. When this happens, the failure is transient and a re-run resolves it.
- The refresh token stored in GitHub Actions secrets must be rotated when it expires. Current token lifetime: ~90 days from last use.

**Neutral but worth knowing:**

- The harness account (`sharepoint-mcp-harness@xmvsolutions.onmicrosoft.com` or equivalent) is scoped exclusively to the sandbox site. A compromised CI secret gives access only to test data, not production.
- The harness job is skipped on forks (GitHub does not expose repository secrets to fork PRs). External contributors should note that their PRs will not have a harness result; the maintainer re-runs harness after review.

## Implementation notes

- Initial CI harness setup: part of the `v0.1.0` project bootstrap.
- Harness token stored as `SP_HARNESS_TOKEN` (or equivalent) in GitHub repository secrets.
- Discovered and fixed during this project: Graph's copy endpoint returns `303 See Other` (not just `202 Accepted`) — unit tests with `respx` mocks could not catch this; harness did. See PR #97.
- Error-path harness tests added in PR #98 following `ENGINEERING_PRINCIPLES.md` § 5 update requiring error-path coverage.

## References

- `ENGINEERING_PRINCIPLES.md` § 5 — "Harness tests in CI: a project-specific trade-off"
- `ENGINEERING_PRINCIPLES.md` § 16 — "Project-specific decisions as permanent records"
- PR #97 — first PR where the CI harness job caught a real Graph API contract violation (303 response)
