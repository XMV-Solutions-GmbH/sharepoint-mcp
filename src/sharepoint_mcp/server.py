# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""MCP server: registers the `sp_*` tools with FastMCP and runs on stdio.

Each tool is wrapped with explicit `ToolAnnotations` so MCP clients
(notably Claude Code's permission system) can render the right
prompt — read-only tools get a different treatment from destructive
ones. The annotations are part of our security story: if we lie
here, the client can't make sensible safety decisions.

**Read-only by default.** Write tools (sp_open, sp_save, sp_release)
are only registered when `SP_ALLOW_WRITES=true` (or =1 / =yes / =on)
is set in the environment. This is belt-and-suspenders to Claude
Code's per-call permission prompts: if you don't even register the
write tools, the agent literally can't call them, regardless of
whether the user accidentally clicks "Always allow" on the wrong
prompt.
"""

from __future__ import annotations

import logging
import os
import sys
from typing import Any

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from sharepoint_mcp.auth.login_tools import login_begin as _do_login_begin
from sharepoint_mcp.auth.login_tools import login_status as _do_login_status
from sharepoint_mcp.tools.bulk import open_many as _do_open_many
from sharepoint_mcp.tools.bulk import save_many as _do_save_many
from sharepoint_mcp.tools.changes import changes as _do_changes
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
from sharepoint_mcp.tools.open_file import open_file as _do_open
from sharepoint_mcp.tools.pages import page_read as _do_page_read
from sharepoint_mcp.tools.pages import page_update as _do_page_update
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
from sharepoint_mcp.tools.sites import subsites as _do_subsites
from sharepoint_mcp.tools.status import status as _do_status
from sharepoint_mcp.tools.trash import trash_list as _do_trash_list

PROFILE_ENV = "SP_PROFILE"
DEFAULT_PROFILE = "default"
ALLOW_WRITES_ENV = "SP_ALLOW_WRITES"
_TRUE_VALUES = frozenset({"1", "true", "yes", "on"})


def _get_profile() -> str:
    return os.environ.get(PROFILE_ENV, DEFAULT_PROFILE)


def writes_enabled() -> bool:
    """True iff `SP_ALLOW_WRITES` is set to a recognised truthy value.

    Default (unset / empty / anything else): writes are NOT enabled,
    matching the read-only-default policy.
    """
    return os.environ.get(ALLOW_WRITES_ENV, "").strip().lower() in _TRUE_VALUES


def register_login_tools(mcp_instance: FastMCP) -> None:
    """Register the integrated-login MCP tools (sp_login_begin, sp_login_status).

    These are always available — they're the prerequisite for using
    everything else, so gating them by SP_ALLOW_WRITES would be a
    chicken-and-egg trap.
    """

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
            "the agent should poll sp_login_status until status is "
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
    async def sp_login_begin(
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
            "- 'none' — no token, no flow. Agent should call sp_login_begin.\n\n"
            "Recently-terminal sessions (`expired` / `failed` / `cancelled`) "
            "surface their error via the `error` field instead of falling back "
            "to 'none' — so the agent can render a specific failure message.\n\n"
            "**UX guidance when status='pending'**: render user_code FIRST in "
            "its own one-line code block (no labels), verification_url SECOND "
            "as a plain auto-link below. User copies the code, taps the link, "
            "pastes into the page that opens — same pattern as sp_login_begin."
        ),
    )
    async def sp_login_status(profile: str | None = None) -> dict[str, Any]:
        return await _do_login_status(
            profile=profile if profile is not None else _get_profile(),
        )


def register_read_tools(mcp_instance: FastMCP) -> None:
    """Register the unconditionally-available read tools on `mcp_instance`."""

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Search SharePoint",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "Search the SharePoint document libraries the signed-in user has access to. "
            "Returns matching files with name, path, webUrl, last-modified date, and "
            "author. Read-only — does not modify any SharePoint state. "
            "Filter args: site (URL), folder (path), file_type (extension like 'docx'), "
            "modified_after (ISO date)."
        ),
    )
    def sp_search(
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
            "sp_search hit's web_url, or the SharePoint web UI). Returns each "
            "child with name, type ('folder' or 'file'), size, last-modified date, "
            "and webUrl. Read-only — does not modify SharePoint state."
        ),
    )
    def sp_list(url: str, limit: int = 100) -> list[dict[str, Any]]:
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
            "Read-only — does NOT acquire a checkout/lock; use sp_open for that. "
            "`url` is the file's human-readable web URL (e.g. from sp_search hits)."
        ),
    )
    def sp_read(url: str) -> str:
        return _do_read(url, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="List Checked-Out Files",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "List the files this MCP profile currently has checked out (acquired via "
            "sp_open). Returns each entry's original path, when checkout happened, "
            "and the local working-copy path. Read-only. With verify=True, "
            "additionally queries SharePoint to confirm the server-side lock state "
            "(server_locked + lock_holder fields); costs one Graph call per "
            "registry entry. Default verify=False is sub-second, registry-only — "
            "sp_save's ETag round-trip catches divergence at write time."
        ),
    )
    def sp_status(verify: bool = False) -> list[dict[str, Any]]:
        return _do_status(profile=_get_profile(), verify=verify)

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="List SharePoint File Version History",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "List a SharePoint file's version history. Returns up to `limit` "
            "versions newest-first, each with id (use with sp_get_version), "
            "last_modified, last_modified_by (display name or email), and size. "
            "Read-only. NOTE: per-version comments aren't currently exposed via "
            "Microsoft Graph v1.0 — they land in SharePoint's web UI version "
            "history but not in this response shape."
        ),
    )
    def sp_history(url: str, limit: int = 20) -> list[dict[str, Any]]:
        return _do_history(url, limit=limit, profile=_get_profile())

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
    def sp_sites(query: str | None = None) -> list[dict[str, Any]]:
        return _do_sites(query, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="List SharePoint Sub-Sites",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "List immediate sub-sites under a parent SharePoint site. "
            "`parent_site_url` is the parent's web URL "
            "(e.g. https://contoso.sharepoint.com/sites/parent). "
            "Returns direct children only — recurse on each result's "
            "web_url to walk deeper. Read-only."
        ),
    )
    def sp_subsites(parent_site_url: str) -> list[dict[str, Any]]:
        return _do_subsites(parent_site_url, profile=_get_profile())

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
    def sp_lists(site_url: str) -> list[dict[str, Any]]:
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
    def sp_list_columns(list_url: str) -> list[dict[str, Any]]:
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
    def sp_list_items(
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
    def sp_get_item(list_url: str, item_id: str) -> dict[str, Any]:
        return _do_get_item(list_url, item_id, profile=_get_profile())

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
    def sp_changes(scope_url: str, since: str | None = None) -> dict[str, Any]:
        return _do_changes(scope_url, since=since, profile=_get_profile())

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
    def sp_pages_list(site_url: str) -> list[dict[str, Any]]:
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
    def sp_page_read(page_url: str) -> dict[str, Any]:
        return _do_page_read(page_url, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="List SharePoint Sharing Links",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "List existing sharing links on a SharePoint file or folder. "
            "Each entry: id (use with sp_share_revoke), web_url (the share "
            "URL), type (view/edit/embed/blocksDownload), scope (organization"
            "/anonymous/users), roles, expiration, has_password. "
            "Read-only — does not create or revoke. Use sp_share_create to "
            "make a new link, sp_share_revoke to remove one."
        ),
    )
    def sp_share_list(url: str) -> list[dict[str, Any]]:
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
            "suggesting changes or sharing links."
        ),
    )
    def sp_permissions(url: str) -> list[dict[str, Any]]:
        return _do_permissions(url, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="List SharePoint Recycle Bin",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "List items in the SharePoint site's recycle bin. Returns "
            "each item's id (use with sp_trash_restore), name, size, "
            "deleted_date_time, deleted_from_location (original folder), "
            "and deleted_by (display name). Read-only. NOTE: this tool "
            "currently uses Microsoft Graph's /beta endpoint — the "
            "site-level recycle-bin API has not yet been promoted to "
            "v1.0. Schema may shift; we'll migrate when v1.0 lands."
        ),
    )
    def sp_trash_list(site_url: str, limit: int = 200) -> list[dict[str, Any]]:
        return _do_trash_list(site_url, limit=limit, profile=_get_profile())

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
            "drive_type, and quota. Most read/write tools accept URLs "
            "into any library transparently — sp_drives is the discovery "
            "step when the agent doesn't know which libraries exist yet."
        ),
    )
    def sp_drives(site_url: str) -> list[dict[str, Any]]:
        return _do_drives(site_url, profile=_get_profile())

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
    def sp_followed_sites() -> list[dict[str, Any]]:
        return _do_followed_sites(profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Read SharePoint File Version",
            readOnlyHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "Download a specific historical version of a SharePoint file to a "
            "local temp file. Returns the absolute path. Use sp_history first to "
            "find the version_id you want. Read-only — does NOT acquire a "
            "checkout, does NOT modify SharePoint state."
        ),
    )
    def sp_get_version(url: str, version_id: str) -> str:
        return _do_get_version(url, version_id, profile=_get_profile())


def register_write_tools(mcp_instance: FastMCP) -> None:
    """Register the gated write tools on `mcp_instance`.

    Only invoked when `SP_ALLOW_WRITES` is truthy. The functions
    themselves are real implementations; the gating is purely about
    whether the agent is offered them at tools/list time.
    """

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
            "file as 'checked out by you' until you call sp_save or sp_release. "
            "Returns the local working-copy path. Fails with a clear error if the "
            "file is already checked out by another user."
        ),
    )
    def sp_open(url: str) -> str:
        return _do_open(url, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Save and Checkin SharePoint File",
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
            "stale-write conflicts (file changed by someone else between sp_open "
            "and sp_save) via ETag round-trip and raises a clear error so the "
            "agent can re-open and reconcile. Returns the new version's id, etag, "
            "and webUrl."
        ),
    )
    def sp_save(url: str, comment: str, version: str = "minor") -> dict[str, Any]:
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
            "keep edits made after sp_open."
        ),
    )
    def sp_release(url: str) -> None:
        _do_release(url, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Publish New SharePoint File",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=False,
            openWorldHint=False,
        ),
        description=(
            "Upload a brand-new local file as a new document in a SharePoint folder. "
            "Use for the 'draft + promote' workflow: agent drafts locally, then "
            "publishes to SharePoint as a fresh file. REFUSES if the target path "
            "already exists — use sp_open + sp_save to edit existing files (gives "
            "proper audit comment + version history). `name` defaults to the local "
            "file's basename; override to publish under a different filename. "
            "Returns the new driveItem's webUrl, etag, size, last_modified."
        ),
    )
    def sp_publish(
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
            "(via sp_release on the successful entries). Use when an agent has "
            "to edit a known set of files and wants the round-trip latency "
            "amortised across them."
        ),
    )
    def sp_open_many(urls: list[str]) -> list[dict[str, Any]]:
        return _do_open_many(urls, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Bulk Save and Checkin SharePoint Files",
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
    def sp_save_many(operations: list[dict[str, Any]]) -> list[dict[str, Any]]:
        # Validation of per-op shape happens inside save_many.
        return _do_save_many(
            operations,  # type: ignore[arg-type]
            profile=_get_profile(),
        )

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
            "schema (use sp_list_columns to inspect). Returns the new "
            "item with its server-assigned id. list_url shape: "
            "https://<host>/sites/<name>/Lists/<list-name>."
        ),
    )
    def sp_create_item(list_url: str, fields: dict[str, Any]) -> dict[str, Any]:
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
            "comes from sp_list_items."
        ),
    )
    def sp_update_item(list_url: str, item_id: str, fields: dict[str, Any]) -> dict[str, Any]:
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
            "use sp_trash_list to find it for ~93 days afterwards. "
            "list_url shape: https://<host>/sites/<name>/Lists/<list-name>."
        ),
    )
    def sp_delete_item(list_url: str, item_id: str) -> None:
        _do_delete_item(list_url, item_id, profile=_get_profile())

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
    def sp_share_create(
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
            "sp_share_create or sp_share_list. Idempotent: re-revoking an "
            "already-revoked link is a 404 from Graph (we propagate)."
        ),
    )
    def sp_share_revoke(url: str, link_id: str) -> None:
        _do_share_revoke(url, link_id, profile=_get_profile())

    @mcp_instance.tool(
        annotations=ToolAnnotations(
            title="Update SharePoint Page",
            readOnlyHint=False,
            destructiveHint=True,
            idempotentHint=True,
            openWorldHint=False,
        ),
        description=(
            "Update a SharePoint Page's metadata. Pass the fields you "
            "want to change as kwargs: title, description, "
            "thumbnail_web_url. Pass None (default) to leave a field "
            "unchanged. At least one field is required. Canvas-layout "
            "(web-parts) edits are NOT supported in v0.3 — round-tripping "
            "the deep nested JSON safely needs more design work; tracked "
            "as a follow-up."
        ),
    )
    def sp_page_update(
        page_url: str,
        title: str | None = None,
        description: str | None = None,
        thumbnail_web_url: str | None = None,
    ) -> dict[str, Any]:
        return _do_page_update(
            page_url,
            title=title,
            description=description,
            thumbnail_web_url=thumbnail_web_url,
            profile=_get_profile(),
        )


def _build_server() -> FastMCP:
    """Build and return a FastMCP server with the right tools registered."""
    server = FastMCP("mcp-server-sharepoint")
    register_login_tools(server)
    register_read_tools(server)
    if writes_enabled():
        register_write_tools(server)
    else:
        # One-line note on stderr so users running uvx interactively
        # see why writes are absent. Quiet by default to avoid noise
        # in MCP-client-launched contexts (Claude Code captures stderr
        # but doesn't surface it loudly).
        logging.getLogger("sharepoint-mcp").info(
            "SP_ALLOW_WRITES not set — read-only mode (sp_open / sp_save / sp_release "
            "not registered). Set SP_ALLOW_WRITES=true to enable writes.",
        )
    return server


mcp: FastMCP = _build_server()


def run() -> None:
    """Start the MCP server on stdio.

    Blocks until stdin closes.
    """
    mcp.run()


# Suppress the "imported but unused" hint for the sys import — it's
# kept for future stderr-printing use that we may need cross-module.
_ = sys
