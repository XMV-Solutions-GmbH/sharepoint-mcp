# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""sp_drive_file_checkin — upload a local working copy, checkin, return new version id.

Three Graph calls per save:

1. `PUT /drives/{id}/items/{id}/content` with `If-Match: <etag>` —
   uploads the new content. Server returns 412 if the ETag doesn't
   match (someone else changed the file under us); we surface that
   as `StaleWriteError`.
2. `POST /drives/{id}/items/{id}/checkin` with `comment` + optional
   `checkInAs: "published"` for a major-version bump.
3. `GET /drives/{id}/items/{id}/versions?$top=1&$orderby=...` to
   read back the version id we created.

On success, the registry entry is removed (file is no longer locked
on the server) and the local working copy is deleted (caller already
saw the file via sp_drive_file_checkout's return path; keeping the working file
around invites confusion about what's authoritative).

`comment` is required and must be non-empty — the same audit-trail
discipline that sp_drive_file_checkin exists to preserve.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Literal

import httpx

from sharepoint_mcp.auth import get_token
from sharepoint_mcp.checkout_registry import CheckoutRegistry
from sharepoint_mcp.tools._common import GRAPH_BASE
from sharepoint_mcp.tools._upload import (
    StaleUploadSessionError,
    chunked_upload_threshold_bytes,
    upload_resumable,
)

VersionLevel = Literal["minor", "major"]


class NotCheckedOutError(RuntimeError):
    """The path isn't in the local checkout registry — call sp_drive_file_checkout first."""


class StaleWriteError(RuntimeError):
    """ETag mismatch — file changed between checkout and checkin.

    Caller should re-`sp_drive_file_checkout` to reconcile, then re-apply edits.
    """


def save(
    url: str,
    *,
    comment: str,
    version: VersionLevel = "minor",
    profile: str = "default",
    http: httpx.Client | None = None,
) -> dict[str, Any]:
    """Upload working copy, checkin, return version metadata.

    `comment` is required and must be non-empty (audit trail).
    `version="major"` creates a published version; `"minor"` (the
    default) creates a draft.

    Returns: `{"version_id": "<str>", "etag": "<str>", "web_url": "<str>"}`.

    Raises:
        ValueError: empty url, empty comment, invalid version level.
        NotCheckedOutError: no registry entry for `url` — sp_drive_file_checkout
            wasn't called first.
        StaleWriteError: ETag mismatch (412) — the file changed
            underneath the open lock; re-sp_drive_file_checkout required.
        FileNotFoundError: working-copy file is missing on disk.
        httpx.HTTPStatusError: any other non-2xx Graph response.
    """
    if not url or not url.strip():
        raise ValueError("sp_drive_file_checkin requires a non-empty url")
    if not comment or not comment.strip():
        raise ValueError(
            "sp_drive_file_checkin requires a non-empty comment for the audit trail"
        )
    if version not in ("minor", "major"):
        raise ValueError(f"version must be 'minor' or 'major', got {version!r}")

    registry = CheckoutRegistry(profile=profile)
    entry = registry.get(url)
    if entry is None:
        raise NotCheckedOutError(
            f"sp_drive_file_checkin called without a prior sp_drive_file_checkout for {url!r}. "
            "Call sp_drive_file_checkout first to acquire the checkout lock.",
        )

    local_file = Path(entry.local_path)
    if not local_file.exists():
        raise FileNotFoundError(
            f"Working copy missing at {entry.local_path!r}; "
            "the local file was deleted between sp_drive_file_checkout and sp_drive_file_checkin.",
        )

    token = get_token(profile)
    auth_headers = {"Authorization": f"Bearer {token}"}
    client = http if http is not None else httpx.Client(timeout=120.0, follow_redirects=True)
    try:
        # 1) Upload content. Single-shot PUT for files under the chunked
        #    threshold (default 100 MB; configurable); resumable upload
        #    session for larger files (Graph caps single-shot at 250 MB).
        threshold = chunked_upload_threshold_bytes()
        file_size = local_file.stat().st_size
        if file_size > threshold:
            try:
                upload = upload_resumable(
                    client,
                    drive_id=entry.drive_id,
                    item_id=entry.item_id,
                    local_file=local_file,
                    etag=entry.etag,
                    auth_headers=auth_headers,
                )
            except StaleUploadSessionError as exc:
                raise StaleWriteError(
                    f"File changed under us between sp_drive_file_checkout and "
                    f"sp_drive_file_checkin for {url!r}. Call "
                    "sp_drive_file_checkout_discard then sp_drive_file_checkout "
                    "again to reconcile.",
                ) from exc
        else:
            put_response = client.put(
                f"{GRAPH_BASE}/drives/{entry.drive_id}/items/{entry.item_id}/content",
                headers={**auth_headers, "If-Match": entry.etag},
                content=local_file.read_bytes(),
            )
            if put_response.status_code == 412:
                raise StaleWriteError(
                    f"File changed under us between sp_drive_file_checkout and "
                    f"sp_drive_file_checkin for {url!r}. Call "
                    "sp_drive_file_checkout_discard then sp_drive_file_checkout "
                    "again to reconcile.",
                )
            put_response.raise_for_status()
            upload = put_response.json()

        # 2) Checkin (releases the lock)
        checkin_body: dict[str, Any] = {"comment": comment}
        if version == "major":
            checkin_body["checkInAs"] = "published"
        checkin_response = client.post(
            f"{GRAPH_BASE}/drives/{entry.drive_id}/items/{entry.item_id}/checkin",
            headers={**auth_headers, "Content-Type": "application/json"},
            json=checkin_body,
        )
        checkin_response.raise_for_status()

        # 3) Fetch the version id of the version we just created
        versions_response = client.get(
            f"{GRAPH_BASE}/drives/{entry.drive_id}/items/{entry.item_id}/versions",
            headers=auth_headers,
            params={"$top": 1, "$orderby": "lastModifiedDateTime desc"},
        )
        versions_response.raise_for_status()
        versions = versions_response.json().get("value", [])
        version_id = str(versions[0]["id"]) if versions else ""

        # Cleanup: registry + local working file
        registry.remove(url)
        try:
            local_file.unlink()
        except FileNotFoundError:
            pass

        return {
            "version_id": version_id,
            "etag": str(upload.get("eTag") or ""),
            "web_url": str(upload.get("webUrl") or url),
        }
    finally:
        if http is None:
            client.close()
