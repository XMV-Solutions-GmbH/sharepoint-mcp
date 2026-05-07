<!--
SPDX-License-Identifier: MIT OR Apache-2.0
SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
SPDX-FileContributor: David Koller <david.koller@xmv.de>
-->

# Spike: working-directory cleanup policy on crash

**Date**: 2026-05-07
**Issue**: [#23](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues/23)
**Decision**: **Persistent registry + explicit `sp_release` for v0.1; no TTL, no automatic server-side reconciliation. Server-side reconciliation in v0.2.**

---

## Question

If the MCP process crashes mid-checkout (kernel panic, OOM-killer, `kill -9`, ssh-disconnect), what becomes of the local working copy and the server-side checkout lock? Two general approaches:

1. **TTL-based**: each registry entry carries an expiry; `sp_status` (or a periodic sweeper) drops or releases entries past their TTL.
2. **Explicit reconciliation**: `sp_status` cross-checks every entry against SharePoint's actual lock state on every call, and surfaces drift to the agent.

## What v0.1 actually needs

The crash scenarios in practice:

| Failure | Server lock | Local file | Registry entry |
|---|---|---|---|
| MCP process crashes mid-`sp_open` (after lock, before write) | held | maybe partial | maybe missing |
| MCP process crashes between `sp_open` and `sp_save` | held | present | present |
| Whole machine reboots between `sp_open` and `sp_save` | held | present | present |
| User decides not to commit edits | held until `sp_release` | present | present |

The thing that consistently survives a crash is the **server-side lock** plus the **local registry file** (it's persisted to disk on every `add`). Local working files survive normal reboots; only true disk-loss removes them.

The thing that handles all four cases without us writing recovery logic is **`sp_release`**: a registry entry says "we believe we have this checked out", `sp_release` calls `discardCheckout` on the server-side and cleans up locally. It's already idempotent ("path not in registry" returns silently).

## Decision

For v0.1:

- **No TTL.** Entries live until `sp_release` or `sp_save` removes them. Registry growth is bounded by how many files an agent has open at once, which in practice is ≤10. Disk cost is negligible (one ~500-byte JSON entry per).
- **No automatic server-side reconciliation in `sp_status`.** The `server_locked` field that the original concept hinted at is omitted; `sp_status` returns the local view only. We document the limitation in the tool description.
- **Trust the registry as source of truth** for "what we have open"; let `sp_save`'s ETag round-trip catch any divergence at the moment that actually matters (the moment of write).

The agent / human recovery path after a crash:

1. Restart the MCP process.
2. Run `sp_status` — see what was checked out before the crash.
3. For each entry: either resume work (the local working copy is still on disk; `sp_save` works as normal), or release it (`sp_release`).
4. Done.

That's a clean recovery story without server-side reconciliation, periodic sweeps, or TTL bookkeeping. It puts the human / agent in the loop, which is appropriate for ISMS-relevant edits anyway.

## Deferred to v0.2

- **`sp_status` server-side reconciliation**: optional `verify=True` flag that queries Graph for each entry to confirm the lock is still held. Useful when the registry has drifted from server state (e.g., admin manually discarded a lock via the SharePoint web UI).
- **TTL with auto-release**: only attractive if a use-case emerges where checkouts pile up unintentionally (e.g., a daemon agent that crashes regularly). Not seen in v0.1 use.
- **Crash-safe upload**: the current `sp_save` writes to working file with `Path.write_bytes`, which is non-atomic. A truly paranoid version would write to a sibling temp + rename. Defer until someone reports a torn write.

## What this rules out

- A "garbage collector" thread / cron that quietly releases checkouts. Hidden side-effects on shared infrastructure are exactly the kind of thing the audit-trail-preserving design exists to avoid. If a checkout is released, the agent should know.
- Best-effort "guess what the user wanted": if the registry is inconsistent (e.g., entry exists but local file is gone), `sp_save` already raises `FileNotFoundError`. The agent surfaces that to the human; no silent recovery.

## Follow-up

Track the v0.2 enhancements as separate tickets when they become relevant. For v0.1, the existing `sp_status` + `sp_release` + `sp_save`-with-ETag combination is the documented recovery path.
