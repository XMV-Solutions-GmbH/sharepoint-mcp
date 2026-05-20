<!-- SPDX-License-Identifier: MIT OR Apache-2.0 -->
# Scenario: directory reorganisation

Reproduces the failure mode from the Anqer production incident
(2026-05-20) that motivated v0.7.0. Pre-restructure, the LLM picked
`sp_upload_new_file` (base64 path, single-level mkdir) for an
edit-and-save task and lost version history.

## User prompt

```
We need to reorganise the harness sandbox's `Shared Documents/scratch/`
folder. Concretely:

1. Create three new folders: `Rituals/sprint-planning`, `Rituals/retros`,
   `Rituals/archive`.
2. Move the existing `old-README.md` in the root of `scratch/` into
   `Rituals/archive/`, renaming it to `archived-README.md` in the process.
3. Upload the local file `/tmp/in/seed-sprint.md` into both
   `Rituals/sprint-planning/` and `Rituals/retros/`.
4. Delete the obsolete files `scratch/obsolete-1.md`, `scratch/obsolete-2.md`,
   `scratch/obsolete-3.md`.
5. List the final contents of `scratch/Rituals/` and report back.

Use the SharePoint MCP. Confirm each step succeeds before moving on.
```

## Expected tool sequence (minimum set)

The LLM should call:

- `sp_drive_folder_create` (×1, recursive — creates all three Rituals/*
  folders) — **not** three separate calls
- `sp_drive_file_move` (×1) — for the rename+move of old-README.md.
  **Not** `sp_drive_file_upload` + `sp_drive_file_delete` (that combo
  loses version history)
- `sp_drive_file_upload` (×2) — one per Rituals/sprint-planning and
  Rituals/retros
- `sp_drive_file_delete` (×3) — one per obsolete file
- `sp_drive_folder_list` (×1) — final listing

**Total: 8 mutating calls + 1 read.** Anything significantly above this
indicates tool-selection confusion.

## Fixture: initial state

`scenarios/anqer-reorg.fixture.yaml` (TODO) seeds the harness sandbox with:

```yaml
scratch/:
  - old-README.md           # content: "v1 readme"
  - obsolete-1.md           # content: "delete me 1"
  - obsolete-2.md           # content: "delete me 2"
  - obsolete-3.md           # content: "delete me 3"
```

## Fixture: expected final state

```yaml
scratch/:
  Rituals/:
    sprint-planning/:
      - seed-sprint.md
    retros/:
      - seed-sprint.md
    archive/:
      - archived-README.md  # content: "v1 readme"  (unchanged from old-README)
```

## Failure modes worth scoring

- **`sp_drive_file_upload` instead of `sp_drive_file_move`** for the rename
  step → loses version history. Hard fail.
- **Multiple `sp_drive_folder_create` calls** when one recursive call would
  do → "didn't read the docstring" signal. Soft fail.
- **Tries `sp_list_item_delete`** for a drive file → category confusion
  (List vs drive). Hard fail. v0.7.0 names are designed to prevent this.
- **Base64 anywhere in the transcript** → `sp_download_binary` was
  removed in v0.7.0; if it shows up, something's wrong with the build.
- **Step count > 15** → likely a retry loop.
