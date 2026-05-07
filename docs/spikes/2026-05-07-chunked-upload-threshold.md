<!--
SPDX-License-Identifier: MIT OR Apache-2.0
SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
SPDX-FileContributor: David Koller <david.koller@xmv.de>
-->

# Spike: chunked-upload threshold for sp_save

**Date**: 2026-05-07
**Issue**: [#24](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues/24)
**Decision**: **Single-shot `PUT /content` for v0.1 — supports up to 250 MB per file. Resumable upload sessions deferred to v0.2.**

---

## Question

Microsoft Graph offers two upload mechanisms for `driveItem` content:

1. **Simple upload** — `PUT /drives/{id}/items/{id}/content` with the file body in the request. One round-trip. Hard size cap (per Graph docs).
2. **Resumable upload session** — `POST /createUploadSession` returns an upload-URL, then chunks are uploaded as `PUT` requests against that URL. Supports very large files; survives transient errors.

Which one does v0.1 use, and at what size do we switch?

## Microsoft's actual limits (verified 2026-05-07 against current docs)

- **Simple `PUT /content` on driveItem**: up to **250 MB** for SharePoint and OneDrive. Files larger than this fail with 413.
- **Resumable upload session**: up to **250 GB** (effectively unlimited for our use case).
- Microsoft *recommends* using resumable sessions for files larger than **4 MB** as a defensive choice (better behaviour on flaky networks), but it's not enforced.

(Historical note: older Graph versions had a 4 MB cap on simple PUT for some endpoints. The current 250 MB ceiling for SharePoint driveItems is documented at <https://learn.microsoft.com/en-us/graph/api/driveitem-put-content>.)

## What v0.1 actually saves

The audience for `sharepoint-mcp` is AI agents editing ISMS-relevant policy documents, contracts, and procedure docs. Realistic file-size distribution:

- Markdown / plain text: typically <100 KB.
- DOCX policy docs: typically <2 MB. Largest seen in practice: ~10 MB (with embedded screenshots).
- PPTX with embedded media: up to ~50 MB. Rare.
- Anything ≥100 MB: should not be in a SharePoint policy library at all (probably belongs in a binary store / DAM).

**95th percentile of in-scope edits is under 50 MB.** 250 MB is a comfortable ceiling.

## Decision

**Single-shot `PUT /content` for v0.1, with a clear error if Microsoft refuses (which it would only do at 250 MB+).** The current `save()` implementation already does this — we keep it as-is.

Reasons:

1. **The use case doesn't need it.** No ISMS document the agent is plausibly editing approaches the 250 MB cap.
2. **Simpler error semantics.** Single-PUT either succeeds or fails atomically; resumable sessions add a state machine (session created → chunks uploaded → committed) that interacts unpleasantly with our ETag-based stale-write detection.
3. **Faster path for the 99% case.** One round-trip beats N+2 for the small files we care about.

If a user hits the 250 MB cap, Microsoft returns a clear 413 Payload Too Large; we propagate it as an `httpx.HTTPStatusError`. The error message is enough for the agent to surface "this file is too large for the simple-upload path" to the human.

## When to revisit (v0.2 trigger)

Add resumable upload session support when **any** of these is true:

- A user reports a real-world file >250 MB they need to edit via this tool.
- We add OneDrive personal support where larger personal files (videos, image archives) might land in scope.
- We see network-flake failures on multi-MB uploads in harness or production logs (resumable sessions can resume after a transient error; single-PUT can't).

Until then, the simple PUT is the right call.

## What this rules out

- Speculatively implementing chunked uploads "to be safe". The complexity of a state-machine path that we never exercise is a worse maintenance burden than a clear 413 error from Microsoft for the rare file that's too large.
- A configurable threshold env var (`SP_CHUNKED_UPLOAD_THRESHOLD_MB`). Adds knobs without need; the threshold isn't a thing v0.1 has.

## Follow-up

Tracked for v0.2 if a user actually hits the limit. Until then, no code changes.
