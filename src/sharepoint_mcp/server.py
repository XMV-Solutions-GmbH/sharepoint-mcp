# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""MCP server: registers the `sp_*` tools with FastMCP and runs on stdio.

v0.7.0 restructured the surface around a `sp_<category>_<noun>_<verb>`
nomenclature. See [docs/app-concept.md](../../docs/app-concept.md)
§ Tools exposed and § Tool design principles for the full surface and
the binding invariants (no base64, recursive parent creation, the
SP_TOOL_GROUPS filter).

Tool annotations are part of our security story: if we lie here, the
client can't make sensible safety decisions. Read-only by default —
write tools register only when `SP_ALLOW_WRITES=true`.

Tool group filtering: `SP_TOOL_GROUPS=drive,search` restricts the
registered surface. Default (unset) registers all groups. `auth` is
always registered regardless of filter — it's the prerequisite for
every other call.
"""

from __future__ import annotations

import os
import sys
from typing import Any

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from sharepoint_mcp import __version__
from sharepoint_mcp.auth.flow import (
    ALLOW_WRITES_ENV as _AUTH_FLOW_ALLOW_WRITES_ENV,
)
from sharepoint_mcp.auth.flow import (
    SharepointConsentNotConfiguredError,
    validate_consent_config,
)
from sharepoint_mcp.auth.login_tools import login_begin as _do_login_begin
from sharepoint_mcp.auth.login_tools import login_status as _do_login_status
from sharepoint_mcp.tools.bulk import open_many as _do_open_many
from sharepoint_mcp.tools.bulk import save_many as _do_save_many
from sharepoint_mcp.tools.changes import changes as _do_changes
from sharepoint_mcp.tools.copy_file import copy_file as _do_copy_file
from sharepoint_mcp.tools.create_folder import create_folder as _do_create_folder
from sharepoint_mcp.tools.delete_file import delete_file as _do_delete_file
from sharepoint_mcp.tools.file_metadata import file_metadata as _do_file_metadata
from sharepoint_mcp.tools.get_version import get_version as _do_get_version
from sharepoint_mcp.tools.history import history as _do_history
from sharepoint_mcp.tools.list_folder import list_folder as _do_list
from sharepoint_mcp.tools.lists import create_item as _do_create_item
from sharepoint_mcp.tools.lists import delete_item as _do_delete_item
from sharepoint_mcp.tools.lists import get_item as _do_get_item
from sharepoint_mcp.tools.lists import list_columns as _do_list_columns
from sharepoint_mcp.tools.lists import list_items as _do_list_items
from sharepoint_mcp.tools.lists import lists as _do_lists
from sharepoint_mcp.tools.lists import update_item as _do_update_item
from sharepoint_mcp.tools.move_file import move_file as _do_move_file
from sharepoint_mcp.tools.open_file import open_file as _do_open
from sharepoint_mcp.tools.pages import page_read as _do_page_read
from sharepoint_mcp.tools.pages import pages_list as _do_pages_list
from sharepoint_mcp.tools.permissions import permissions as _do_permissions
from sharepoint_mcp.tools.publish import publish as _do_publish
from sharepoint_mcp.tools.read import read_file as _do_read
from sharepoint_mcp.tools.release import release as _do_release
from sharepoint_mcp.tools.save import save as _do_save
from sharepoint_mcp.tools.search import search as _do_search
from sharepoint_mcp.tools.sharing import share_create as _do_share_create
from sharepoint_mcp.tools.sharing import share_list as _do_share_list
from sharepoint_mcp.tools.sharing import share_revoke as _do_share_revoke
from sharepoint_mcp.tools.sites import drives as _do_drives
from sharepoint_mcp.tools.sites import followed_sites as _do_followed_sites
from sharepoint_mcp.tools.sites import sites as _do_sites
from sharepoint_mcp.tools.status import status as _do_status
from sharepoint_mcp.tools.trash import trash_list as _do_trash_list

PROFILE_ENV = "SP_PROFILE"
DEFAULT_PROFILE = "default"
TOOL_GROUPS_ENV = "SP_TOOL_GROUPS"
# Re-exported for backwards-compat with v0.4.x importers.
ALLOW_WRITES_ENV = _AUTH_FLOW_ALLOW_WRITES_ENV

# The six categories from docs/app-concept.md § Tools exposed. `auth`
# is always registered regardless of the filter — it's the bootstrap
# for every other call.
ALL_TOOL_GROUPS: tuple[str, ...] = ("auth", "site", "drive", "list", "share", "search")


class SharepointToolGroupsError(ValueError):
    """Raised on startup when SP_TOOL_GROUPS contains an unknown group name."""


def _get_profile() -> str:
    return os.environ.get(PROFILE_ENV, DEFAULT_PROFILE)


def writes_enabled() -> bool:
    """True iff `SP_ALLOW_WRITES` is set to exactly `"true"`.

    Strict parser since v0.5 — raises
    `SharepointConsentNotConfiguredError` if the env var is unset,
    empty, or has a value other than `true` or `false`.
    """
    return validate_consent_config()


def parse_tool_groups(value: str | None) -> set[str]:
    """Parse `SP_TOOL_GROUPS` env value into the set of enabled groups.

    Default (None / empty string) = all groups. `auth` is added back
    unconditionally even if omitted — the agent can't bootstrap any
    other call without it.

    Raises `SharepointToolGroupsError` with a clear message if the
    value lists a group not in `ALL_TOOL_GROUPS`. Loud failure is
    deliberate: silent-skip on a typo would hide tools the operator
    expected to be available.
    """
    if value is None or not value.strip():
        return set(ALL_TOOL_GROUPS)
    requested = {seg.strip().lower() for seg in value.split(",") if seg.strip()}
    unknown = requested - set(ALL_TOOL_GROUPS)
    if unknown:
        raise SharepointToolGroupsError(
            f"SP_TOOL_GROUPS contains unknown group(s): {sorted(unknown)}. "
            f"Valid groups: {list(ALL_TOOL_GROUPS)}."
        )
    requested.add("auth")
    return requested


# ── auth ──────────────────────────────────────────────────────────────────


def register_auth_tools(mcp_instance: FastMCP) -> None:
    """Register `sp_auth_*` — login lifecycle. Always on regardless of SP_TOOL_GROUPS."""

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Begin SharePoint Sign-In",
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=False,
            openWorldHint=True,
        ),
        description=(
            "Initiate Microsoft Identity Device Code login for `profile` "
            "(default 'default'). Non-blocking: returns within ~1s with "
            "user_code + verification_url. A background task continues to "
            "poll Microsoft Identity and writes the token on success — "
            "the agent should poll sp_auth_status until status is "
            "'signed_in', or until the user completes / cancels the flow.\n\n"
            "Idempotency: if a pending session already exists for this "
            "profile, the existing session is returned unchanged unless "
            "force=True (which cancels and restarts).\n\n"
            "**UX guidance for surfacing the result to the user**: render "
            "user_code FIRST in its own one-line code block (no labels, "
            "no whitespace) and verification_url SECOND on its own line as "
            "a plain auto-link (NOT in a code block). The user copies the "
            "code first, then taps the link, then pastes into the page "
            "that opens — this minimises app-switching on mobile / smartphone "
            "MCP clients. URL inside a code block suppresses link rendering; "
            "code with labels makes copy-paste include the label noise."
        ),
    )
    async def sp_auth_begin(
        profile: str | None = None,
        force: bool = False,
    ) -> dict[str, Any]:
        return await _do_login_begin(
            profile=profile if profile is not None else _get_profile(),
            force=force,
        )

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="SharePoint Sign-In Status",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "Return the current sign-in state for `profile` (default 'default'). "
            "Three states the agent can act on directly:\n\n"
            "- 'signed_in' — valid token on disk (regardless of how it got there: "
            "  CLI login, prior tool-flow, or just refreshed silently). "
            "  signed_in_user_upn populated. The agent can proceed.\n"
            "- 'pending' — Device Code flow in progress. user_code, "
            "  verification_url, time_remaining_s populated.\n"
            "- 'none' — no token, no flow. Agent should call sp_auth_begin.\n\n"
            "Recently-terminal sessions (`expired` / `failed` / `cancelled`) "
            "surface their error via the `error` field instead of falling back "
            "to 'none' — so the agent can render a specific failure message.\n\n"
            "**UX guidance when status='pending'**: render user_code FIRST in "
            "its own one-line code block (no labels), verification_url SECOND "
            "as a plain auto-link below. User copies the code, taps the link, "
            "pastes into the page that opens — same pattern as sp_auth_begin."
        ),
    )
    async def sp_auth_status(profile: str | None = None) -> dict[str, Any]:
        return await _do_login_status(
            profile=profile if profile is not None else _get_profile(),
        )


# ── site ──────────────────────────────────────────────────────────────────


def register_site_tools(mcp_instance: FastMCP) -> None:
    """Register `sp_site_*` — site / library / page / recycle-bin discovery (read-only)."""

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Search SharePoint Sites",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "Discover SharePoint sites the signed-in user has access to. "
            "`query` is a free-text site-name search (e.g. 'finance'); "
            "omit / leave empty to list all visible sites. Returns each "
            "site's id, name, web_url, description, last_modified. "
            "Read-only. Use as the entry point when the agent doesn't "
            "yet know which site URL to drill into."
        ),
    )
    def sp_site_list(query: str | None = None) -> list[dict[str, Any]]:
        return _do_sites(query, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="List Followed SharePoint Sites",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "List sites the signed-in user has marked as Followed in "
            "SharePoint. Useful 'my SharePoint' entry point for an agent "
            "starting from the user's curated list rather than guessing "
            "site URLs. Not available in service-principal auth mode "
            "(no signed-in user) — falls back to a clear error there. "
            "Read-only."
        ),
    )
    def sp_site_followed_list() -> list[dict[str, Any]]:
        return _do_followed_sites(profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="List SharePoint Document Libraries",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "List all document libraries (drives) on a SharePoint site — "
            "default Shared Documents plus Site Assets, Style Library, "
            "and any custom libraries. `site_url` is the site's web URL. "
            "Returns each drive's id, name, web_url, description, "
            "drive_type, and quota. Most `sp_drive_*` tools accept URLs "
            "into any library transparently — sp_site_drive_list is the "
            "discovery step when the agent doesn't know which libraries "
            "exist yet."
        ),
    )
    def sp_site_drive_list(site_url: str) -> list[dict[str, Any]]:
        return _do_drives(site_url, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="List SharePoint Pages",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "List all modern SharePoint Pages (Site Pages) on a site. "
            "Returns each page's id, name (filename), title, web_url, "
            "description, page_layout, thumbnail_web_url, last_modified, "
            "last_modified_by. Read-only."
        ),
    )
    def sp_site_page_list(site_url: str) -> list[dict[str, Any]]:
        return _do_pages_list(site_url, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Read SharePoint Page",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "Fetch a single SharePoint Page including its canvasLayout "
            "(sections, columns, web parts) as JSON. page_url shape: "
            "https://<host>/sites/<name>/SitePages/<page>.aspx. Read-only."
        ),
    )
    def sp_site_page_read(page_url: str) -> dict[str, Any]:
        return _do_page_read(page_url, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="List SharePoint Recycle Bin",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "List items in the SharePoint site's recycle bin. Returns "
            "each item's id, name, size, deleted_date_time, "
            "deleted_from_location (original folder), and deleted_by "
            "(display name). Read-only. NOTE: Microsoft Graph does not "
            "expose a restore action at site scope; items currently have "
            "to be restored via the SharePoint web UI. This tool uses "
            "Graph's /beta endpoint — the site-level recycle-bin API has "
            "not yet been promoted to v1.0. Schema may shift; we'll "
            "migrate when v1.0 lands."
        ),
    )
    def sp_site_trash_list(site_url: str, limit: int = 200) -> list[dict[str, Any]]:
        return _do_trash_list(site_url, limit=limit, profile=_get_profile())


# ── drive ─────────────────────────────────────────────────────────────────


def register_drive_read_tools(mcp_instance: FastMCP) -> None:
    """Register read-only `sp_drive_*` tools (files / folders in document libraries)."""

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="List SharePoint Folder",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "List the immediate children of a SharePoint or OneDrive folder. "
            "`url` is the folder's human-readable web URL (e.g. from a previous "
            "sp_search_query hit's web_url, or the SharePoint web UI). Returns each "
            "child with name, type ('folder' or 'file'), size, last-modified date, "
            "and webUrl. Read-only — does not modify SharePoint state."
        ),
    )
    def sp_drive_folder_list(url: str, limit: int = 100) -> list[dict[str, Any]]:
        return _do_list(url, limit=limit, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Read SharePoint File",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "Download a SharePoint file's content to a local temp file. Returns the "
            "absolute path of the temp file with the original extension preserved. "
            "Read-only — does NOT acquire a checkout/lock; use sp_drive_file_checkout "
            "for that. `url` is the file's human-readable web URL (e.g. from "
            "sp_search_query hits). The LLM consumes the file via filesystem tools "
            "(Read, Bash) — no base64 round-trip."
        ),
    )
    def sp_drive_file_read(url: str) -> str:
        return _do_read(url, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="List SharePoint File Version History",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "List a SharePoint file's version history. Returns up to `limit` "
            "versions newest-first, each with id (use with sp_drive_file_version_get), "
            "last_modified, last_modified_by (display name or email), and size. "
            "Read-only. NOTE: per-version comments aren't currently exposed via "
            "Microsoft Graph v1.0 — they land in SharePoint's web UI version "
            "history but not in this response shape."
        ),
    )
    def sp_drive_file_history(url: str, limit: int = 20) -> list[dict[str, Any]]:
        return _do_history(url, limit=limit, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Read SharePoint File Version",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "Download a specific historical version of a SharePoint file to a "
            "local temp file. Returns the absolute path. Use sp_drive_file_history "
            "first to find the version_id you want. Read-only — does NOT acquire a "
            "checkout, does NOT modify SharePoint state."
        ),
    )
    def sp_drive_file_version_get(url: str, version_id: str) -> str:
        return _do_get_version(url, version_id, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Track SharePoint Drive Changes",
            readOnlyHint=True,
            idempotentHint=False,
            openWorldHint=False,
        ),
        description=(
            "Return items in a SharePoint site's default drive that "
            "changed since the optional `since` cursor — Microsoft "
            "Graph delta query. First call (since=None) returns the "
            "full item set + an initial cursor. Subsequent calls with "
            "the cursor return only created/modified/deleted items "
            "since that cursor. Result: {items: [...], cursor: str}. "
            "The cursor is opaque — store it (the agent typically "
            "puts it in conversation memory or a scratchpad) and pass "
            "it back as `since`. A stale cursor surfaces as a 410 "
            "Gone error; drop it and call again with since=None for "
            "a full re-sync. Read-only."
        ),
    )
    def sp_drive_change_track(scope_url: str, since: str | None = None) -> dict[str, Any]:
        return _do_changes(scope_url, since=since, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="List Checked-Out Drive Files",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "List the files this MCP profile currently has checked out (acquired via "
            "sp_drive_file_checkout). Returns each entry's original path, when "
            "checkout happened, and the local working-copy path. Read-only. "
            "With verify=True, additionally queries SharePoint to confirm the "
            "server-side lock state (server_locked + lock_holder fields); costs "
            "one Graph call per registry entry. Default verify=False is sub-second, "
            "registry-only — sp_drive_file_checkin's ETag round-trip catches "
            "divergence at write time."
        ),
    )
    def sp_drive_checkout_list(verify: bool = False) -> list[dict[str, Any]]:
        return _do_status(profile=_get_profile(), verify=verify)


def register_drive_write_tools(mcp_instance: FastMCP) -> None:
    """Register write `sp_drive_*` tools. Gated by `SP_ALLOW_WRITES=true`."""

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Create SharePoint Folder",
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "Create a folder hierarchy in the site's default document library. "
            "`site_url` is the SharePoint site URL. `path` is the folder path "
            "to create, relative to the document library root — e.g. "
            "'2026/Q2/Reports'. A leading 'Shared Documents/' prefix is stripped "
            "automatically for convenience. Intermediate folders that don't exist "
            "yet are created in one call (recursive mkdir semantics). Existing "
            "folders are silently skipped, making the operation idempotent. "
            "Returns {created, already_existed, web_url}."
        ),
    )
    def sp_drive_folder_create(site_url: str, path: str) -> dict[str, Any]:
        return _do_create_folder(site_url, path, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Upload New SharePoint File",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=False,
        ),
        description=(
            "Upload a brand-new local file as a new document in a SharePoint folder. "
            "Use for the 'draft + promote' workflow: agent drafts locally, then "
            "publishes to SharePoint as a fresh file. REFUSES if the target name "
            "already exists in the folder — use sp_drive_file_checkout + "
            "sp_drive_file_checkin to edit existing files (gives proper audit comment "
            "+ version history). Recursively creates any missing parent folders along "
            "target_folder_url. `name` defaults to the local file's basename; override "
            "to publish under a different filename. Returns the new driveItem's webUrl, "
            "etag, size, last_modified. **No base64** — reads bytes from local_path."
        ),
    )
    def sp_drive_file_upload(
        local_path: str,
        target_folder_url: str,
        name: str | None = None,
    ) -> dict[str, Any]:
        return _do_publish(
            local_path,
            target_folder_url,
            name=name,
            profile=_get_profile(),
        )

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Delete SharePoint Drive File",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "Delete a file or folder in a SharePoint document library. "
            "SharePoint sends it to the site recycle bin — recoverable for "
            "~93 days via sp_site_trash_list. Does NOT hard-delete. "
            "site_url: https://<host>/sites/<name>. "
            "path: drive-relative path, e.g. '2026/Q2/report.md'."
        ),
    )
    def sp_drive_file_delete(site_url: str, path: str) -> dict[str, Any]:
        return _do_delete_file(site_url, path, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Move / Rename SharePoint Drive File",
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=False,
            openWorldHint=False,
        ),
        description=(
            "Move or rename a file or folder in a SharePoint document library. "
            "destination_path is the full path of the item after the move "
            "(not the destination folder). Last segment = new name; preceding "
            "segments = existing destination folder. "
            "Combine move + rename in one call by changing both folder and name. "
            "site_url: https://<host>/sites/<name>. "
            "Paths are drive-relative, e.g. 'Archive/2026/report.md'."
        ),
    )
    def sp_drive_file_move(
        site_url: str, source_path: str, destination_path: str
    ) -> dict[str, Any]:
        return _do_move_file(site_url, source_path, destination_path, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Copy SharePoint Drive File",
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=False,
            openWorldHint=False,
        ),
        description=(
            "Copy a file in a SharePoint document library to a new path. "
            "destination_path is the full path of the copy after creation "
            "(not the destination folder). Last segment = name of the copy; "
            "preceding segments = existing destination folder. "
            "The Graph copy operation is asynchronous; this tool polls until "
            "completed (up to 60 s). "
            "site_url: https://<host>/sites/<name>. "
            "Paths are drive-relative, e.g. 'Projects/ACME/contract.docx'."
        ),
    )
    def sp_drive_file_copy(
        site_url: str, source_path: str, destination_path: str
    ) -> dict[str, Any]:
        return _do_copy_file(site_url, source_path, destination_path, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Read/Write SharePoint File Metadata",
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=False,
            openWorldHint=False,
        ),
        description=(
            "Read or update the custom SharePoint column values (metadata) attached "
            "to a document-library file. "
            "Read mode (fields omitted): returns a flat dict of all column "
            "values for the file's list item — system fields (Modified, Author, "
            "etc.) plus any custom library columns. "
            "Write mode (fields provided): PATCHes the supplied column key→value "
            "pairs and returns the full updated field state. Only keys present in "
            "`fields` are touched; other columns are unchanged. Use internal column "
            "names (e.g. 'Department', '_Status') — the same keys from read mode."
        ),
    )
    def sp_drive_file_metadata(
        url: str,
        fields: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        return _do_file_metadata(url, fields=fields, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Checkout SharePoint File",
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=False,
            openWorldHint=False,
        ),
        description=(
            "Acquire a server-side checkout lock on a SharePoint file and download "
            "its current content to a local working-copy path. Other users see the "
            "file as 'checked out by you' until you call sp_drive_file_checkin or "
            "sp_drive_file_checkout_discard. Returns the local working-copy path. "
            "Fails with a clear error if the file is already checked out by "
            "another user."
        ),
    )
    def sp_drive_file_checkout(url: str) -> str:
        return _do_open(url, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Checkin SharePoint File",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=False,
        ),
        description=(
            "Upload the local working copy and checkin the file with a comment, "
            "creating a new version (minor by default, or major if version='major'). "
            "Releases the server-side checkout lock. `comment` is REQUIRED and goes "
            "into the SharePoint audit trail — describe what changed. Detects "
            "stale-write conflicts (file changed by someone else between "
            "sp_drive_file_checkout and sp_drive_file_checkin) via ETag round-trip "
            "and raises a clear error so the agent can re-open and reconcile. "
            "Returns the new version's id, etag, and webUrl."
        ),
    )
    def sp_drive_file_checkin(url: str, comment: str, version: str = "minor") -> dict[str, Any]:
        if version not in ("minor", "major"):
            raise ValueError(f"version must be 'minor' or 'major', got {version!r}")
        return _do_save(
            url,
            comment=comment,
            version=version,  # type: ignore[arg-type]
            profile=_get_profile(),
        )

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Discard SharePoint Checkout",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "Discard a pending checkout without saving any local changes. Releases "
            "the server-side lock, deletes the local working-copy file, and "
            "removes the registry entry. Idempotent: silently no-ops when nothing "
            "is checked out for the given url. Use this when you decide not to "
            "keep edits made after sp_drive_file_checkout."
        ),
    )
    def sp_drive_file_checkout_discard(url: str) -> None:
        _do_release(url, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Bulk Checkout SharePoint Files",
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=False,
            openWorldHint=False,
        ),
        description=(
            "Acquire server-side checkout locks on multiple SharePoint files in "
            "parallel (up to 4 concurrent Graph calls per Microsoft throttling "
            "guidance). Returns one result per input URL in the same order: "
            "{path, status='ok', local_path} on success, "
            "{path, status='error', error} on failure. Per-file failures do NOT "
            "abort the rest — caller decides whether to continue or rollback "
            "(via sp_drive_file_checkout_discard on the successful entries). Use "
            "when an agent has to edit a known set of files and wants the "
            "round-trip latency amortised across them."
        ),
    )
    def sp_drive_file_checkout_bulk(urls: list[str]) -> list[dict[str, Any]]:
        return _do_open_many(urls, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Bulk Checkin SharePoint Files",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=False,
        ),
        description=(
            "Upload local working copies and checkin multiple SharePoint files in "
            "parallel (up to 4 concurrent Graph calls). Each `operations` entry: "
            "{url, comment, version='minor'|'major' (default 'minor')}. `comment` "
            "is required for each — goes into the audit trail per file. Returns "
            "one result per input op in the same order: {path, status='ok', "
            "version_id, etag, web_url} on success, {path, status='error', error} "
            "on failure. Per-file failures do NOT abort the rest. ETag round-trip "
            "for stale-write detection still applies per file."
        ),
    )
    def sp_drive_file_checkin_bulk(operations: list[dict[str, Any]]) -> list[dict[str, Any]]:
        # Validation of per-op shape happens inside save_many.
        return _do_save_many(
            operations,  # type: ignore[arg-type]
            profile=_get_profile(),
        )


# ── list ──────────────────────────────────────────────────────────────────


def register_list_read_tools(mcp_instance: FastMCP) -> None:
    """Register read-only `sp_list_*` tools (SharePoint Lists)."""

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="List SharePoint Lists",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "List all SharePoint Lists on a site (Issue Trackers, Tasks, "
            "Custom Lists, etc.). Returns each list's id, name, "
            "display_name, web_url, description, created_date_time, "
            "last_modified_date_time, and template (e.g. 'genericList', "
            "'documentLibrary', 'tasks'). Read-only."
        ),
    )
    def sp_list_list(site_url: str) -> list[dict[str, Any]]:
        return _do_lists(site_url, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Get SharePoint List Schema",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "Return the column definitions (schema) of a SharePoint List. "
            "Each column: id, display_name, name (internal), description, "
            "required, hidden, read_only, indexed, type (text/choice/number/"
            "boolean/datetime/person/lookup/calculated/hyperlink/currency). "
            "list_url shape: https://<host>/sites/<name>/Lists/<list-name>. "
            "Read-only."
        ),
    )
    def sp_list_column_list(list_url: str) -> list[dict[str, Any]]:
        return _do_list_columns(list_url, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="List SharePoint List Items",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "List items in a SharePoint List with their full fields "
            "expanded. `filter` is an optional OData $filter expression "
            "(e.g. \"fields/Status eq 'Open'\"). `top` caps results "
            "(default 100, max 5000 per Graph). list_url shape: "
            "https://<host>/sites/<name>/Lists/<list-name>. Read-only."
        ),
    )
    def sp_list_item_list(
        list_url: str,
        filter: str | None = None,
        top: int = 100,
    ) -> list[dict[str, Any]]:
        return _do_list_items(list_url, filter=filter, top=top, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Get SharePoint List Item",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "Fetch a single SharePoint List item by id with all expanded "
            "fields. Returns id, created_date_time, last_modified_date_time, "
            "created_by, last_modified_by, web_url, fields (dict). "
            "list_url shape: https://<host>/sites/<name>/Lists/<list-name>. "
            "Read-only."
        ),
    )
    def sp_list_item_get(list_url: str, item_id: str) -> dict[str, Any]:
        return _do_get_item(list_url, item_id, profile=_get_profile())


def register_list_write_tools(mcp_instance: FastMCP) -> None:
    """Register write `sp_list_*` tools. Gated by `SP_ALLOW_WRITES=true`."""

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Create SharePoint List Item",
            readOnlyHint=False,
            destructiveHint=False,
            idempotentHint=False,
            openWorldHint=False,
        ),
        description=(
            "Create a new item in a SharePoint List. `fields` is a dict of "
            "column-internal-name -> value pairs that match the list's "
            "schema (use sp_list_column_list to inspect). Returns the new "
            "item with its server-assigned id. list_url shape: "
            "https://<host>/sites/<name>/Lists/<list-name>."
        ),
    )
    def sp_list_item_create(list_url: str, fields: dict[str, Any]) -> dict[str, Any]:
        return _do_create_item(list_url, fields, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Update SharePoint List Item",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "Patch fields on an existing SharePoint List item. Only the "
            "keys present in `fields` are changed; the rest stay as-is. "
            "Returns the updated fields dict. list_url shape: "
            "https://<host>/sites/<name>/Lists/<list-name>. `item_id` "
            "comes from sp_list_item_list."
        ),
    )
    def sp_list_item_update(list_url: str, item_id: str, fields: dict[str, Any]) -> dict[str, Any]:
        return _do_update_item(list_url, item_id, fields, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Delete SharePoint List Item",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "Delete a SharePoint List item. SharePoint sends it to the "
            "site recycle bin (default behaviour for DELETE on listItem) — "
            "use sp_site_trash_list to find it for ~93 days afterwards. "
            "list_url shape: https://<host>/sites/<name>/Lists/<list-name>. "
            "Operates on **List items only** — for files in document libraries "
            "use sp_drive_file_delete."
        ),
    )
    def sp_list_item_delete(list_url: str, item_id: str) -> None:
        _do_delete_item(list_url, item_id, profile=_get_profile())


# ── share ─────────────────────────────────────────────────────────────────


def register_share_read_tools(mcp_instance: FastMCP) -> None:
    """Register read-only `sp_share_*` tools (sharing links + permissions)."""

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="List SharePoint Sharing Links",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "List existing **sharing links** on a SharePoint file or folder. "
            "Each entry: id (use with sp_share_link_revoke), web_url (the share "
            "URL), type (view/edit/embed/blocksDownload), scope "
            "(organization/anonymous/users), roles, expiration, has_password. "
            "Read-only — does not create or revoke. Use sp_share_link_create to "
            "make a new link, sp_share_link_revoke to remove one. "
            "SCOPE: only sharing-link permissions. For ALL access grants "
            "(direct user/group assignments, inherited site permissions, "
            "plus sharing links) use sp_share_permission_list instead."
        ),
    )
    def sp_share_link_list(url: str) -> list[dict[str, Any]]:
        return _do_share_list(url, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="List SharePoint Permissions",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "List who has access to a SharePoint file, folder, or site. "
            "Pass a site URL for site-level permissions or any item URL "
            "(file or folder) for that item's permissions. Returns each "
            "permission entry with id, roles (read/write/owner), grantee "
            "({type, display_name, email, link_type, link_scope}), and "
            "inherited flag. Read-only — does not modify any permission "
            "state. Use this to answer 'who can see/edit this?' before "
            "suggesting changes or sharing links. "
            "SCOPE: all permission grants — direct user/group assignments, "
            "inherited site permissions, AND sharing links. To list only "
            "sharing links (and get their `id` for sp_share_link_revoke), use "
            "sp_share_link_list instead."
        ),
    )
    def sp_share_permission_list(url: str) -> list[dict[str, Any]]:
        return _do_permissions(url, profile=_get_profile())


def register_share_write_tools(mcp_instance: FastMCP) -> None:
    """Register write `sp_share_*` tools. Gated by `SP_ALLOW_WRITES=true`."""

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Create SharePoint Sharing Link",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=False,
        ),
        description=(
            "Create a sharing link on a SharePoint file or folder. Returns "
            "{id, web_url, type, scope, ...}. **Defaults are conservative**: "
            "type='view' (read-only), scope='organization' (signed-in users in "
            "your tenant only). To create a public link the agent must "
            "explicitly pass scope='anonymous' — the most common ISMS-audit "
            "finding. type='edit' grants WRITE to anyone with the URL within "
            "scope; combine with scope='anonymous' only on explicit user "
            "request. Optional: `expires` (ISO 8601 datetime) and `password` "
            "(only meaningful for anonymous; tenant may disable). "
            "Marked destructive in MCP annotations because the link creates "
            "a discoverable access path that persists until revoked."
        ),
    )
    def sp_share_link_create(
        url: str,
        type: str = "view",
        scope: str = "organization",
        expires: str | None = None,
        password: str | None = None,
    ) -> dict[str, Any]:
        return _do_share_create(
            url,
            type=type,
            scope=scope,
            expires=expires,
            password=password,
            profile=_get_profile(),
        )

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Revoke SharePoint Sharing Link",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "Revoke (delete) a sharing-link permission. After this call the "
            "share URL stops working. `link_id` is the permission id from "
            "sp_share_link_create or sp_share_link_list. Idempotent: re-revoking "
            "an already-revoked link is a 404 from Graph (we propagate)."
        ),
    )
    def sp_share_link_revoke(url: str, link_id: str) -> None:
        _do_share_revoke(url, link_id, profile=_get_profile())


# ── search ────────────────────────────────────────────────────────────────


def register_search_tools(mcp_instance: FastMCP) -> None:
    """Register `sp_search_*` (currently driveItem-only)."""

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Search SharePoint",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "Search SharePoint document libraries the signed-in user has access to. "
            "Returns matching files with name, path, webUrl, last-modified date, and "
            "author. Read-only. Filter args: site (URL), folder (path), file_type "
            "(extension like 'docx'), modified_after (ISO date). "
            "SCOPE: currently driveItem-only (files in document libraries). "
            "Searching List items or sites by content is not yet supported — use "
            "sp_list_item_list with an OData filter for List-item lookup, "
            "sp_site_list for site discovery."
        ),
    )
    def sp_search_query(
        query: str,
        site: str | None = None,
        folder: str | None = None,
        file_type: str | None = None,
        modified_after: str | None = None,
        limit: int = 25,
    ) -> list[dict[str, Any]]:
        return _do_search(
            query,
            site=site,
            folder=folder,
            file_type=file_type,
            modified_after=modified_after,
            limit=limit,
            profile=_get_profile(),
        )


# ── build ─────────────────────────────────────────────────────────────────


def _emit_startup_banner(groups: set[str], writes: bool) -> None:
    """Emit one stderr line on startup announcing version + config.

    Lets MCP-client log windows show which version is actually running
    without round-tripping through the MCP protocol. Documented as a
    contract in app-concept.md § Tool design principles.
    """
    sorted_groups = ",".join(sorted(groups))
    sys.stderr.write(
        f"mcp-server-sharepoint {__version__} — "
        f"groups=[{sorted_groups}] "
        f"writes={'true' if writes else 'false'}\n"
    )
    sys.stderr.flush()


def _build_server() -> FastMCP:
    """Build and return a FastMCP server with the right tools registered.

    Two startup-time validations, both designed to fail loudly:

    - `SP_ALLOW_WRITES` must be exactly `true` or `false`; anything
      else raises `SharepointConsentNotConfiguredError` with onboarding
      help.
    - `SP_TOOL_GROUPS`, if set, must contain only known group names;
      a typo raises `SharepointToolGroupsError`.

    Both exceptions propagate so the operator sees them on stderr at
    process-start time, not mid-protocol-handshake.
    """
    writes = validate_consent_config()
    groups = parse_tool_groups(os.environ.get(TOOL_GROUPS_ENV))

    server = FastMCP("mcp-server-sharepoint")
    register_auth_tools(server)  # always on

    if "site" in groups:
        register_site_tools(server)
    if "drive" in groups:
        register_drive_read_tools(server)
        if writes:
            register_drive_write_tools(server)
    if "list" in groups:
        register_list_read_tools(server)
        if writes:
            register_list_write_tools(server)
    if "share" in groups:
        register_share_read_tools(server)
        if writes:
            register_share_write_tools(server)
    if "search" in groups:
        register_search_tools(server)

    _emit_startup_banner(groups, writes)
    return server


# Build at module-import time so MCP-client launchers (uvx, etc.)
# get the validation errors immediately on startup rather than
# mid-protocol-handshake.
try:
    mcp: FastMCP = _build_server()
except (SharepointConsentNotConfiguredError, SharepointToolGroupsError) as err:
    sys.stderr.write(str(err) + "\n")
    sys.stderr.flush()
    raise


def run() -> None:
    """Start the MCP server on stdio.

    Blocks until stdin closes.
    """
    mcp.run()
