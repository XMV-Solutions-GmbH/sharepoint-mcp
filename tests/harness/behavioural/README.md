<!-- SPDX-License-Identifier: MIT OR Apache-2.0 -->
# Behavioural harness — does an LLM pick the right tool?

The pytest harness layer (`tests/harness/test_*.py`) calls each `sp_*`
tool directly against the real Microsoft Graph + a real test tenant.
That's a *functional* harness: it verifies the wire is right.

This directory is the *behavioural* harness: it boots the MCP server
in front of a real Claude agent, gives the agent a scripted user task,
and observes whether the agent picks the right tools. It catches the
failure mode that motivated the v0.7.0 rename — where an agent picked
`sp_upload_new_file` for an edit-and-save task and lost version history.

## How to run a scenario

```bash
# From the repo root:
scripts/run-behavioural-harness.sh anqer-reorg
```

Exit code `0` = pass, non-zero = fail. A detailed report is printed to
stdout: which `sp_*` tools were called, which were expected-but-missing,
which were prohibited-but-called, and any divergence between the
sandbox's final state and the fixture's expectations.

The wrapper checks that `claude` and `uv` are on `$PATH` and forwards
to the Python runner (`runner.py`). The runner spawns
`claude --print --output-format=stream-json` with a temporary
`.mcp.json` that points at the **local source tree's** MCP server (via
`uv run --project <repo>`), NOT at `uvx mcp-server-sharepoint` —
otherwise we'd be testing the published PyPI version instead of the
branch under development.

## Cost reality

Each run consumes real Claude API tokens (your account) AND mutates the
real harness SharePoint sandbox. Don't loop on it. The shell script
fails loudly if `claude` or `uv` are missing so accidental CI loops
without those installed exit fast.

## Per-scenario layout

Three pieces per scenario:

1. **`scenarios/<name>.md`** — the user task in natural language plus
   the human-readable scoring rationale. The runner extracts the
   prompt from the first fenced code block under `## User prompt`.
2. **`scenarios/<name>.fixture.yaml`** — machine-readable fixture: the
   sandbox state to seed, the expected final state to verify, and the
   tool-call expectations. See `anqer-reorg.fixture.yaml` for the
   canonical shape.
3. **`runner.py`** — the orchestrator. One Python module, deliberately
   not a pytest test, so unit tests can cover the parsing/scoring
   logic in regular CI without paying the live-run cost.

## Scoring rules

Per run, the runner records four signals:

- **Expected tools at least once**: every entry in
  `expected_tools_at_least_once` must appear at least once in the
  transcript. Missing → hard fail.
- **Prohibited tools**: any entry in `prohibited_tools` appearing in
  the transcript → hard fail. (E.g. `sp_list_item_delete` for a drive
  task signals category confusion that v0.7.0's nomenclature was
  designed to prevent.)
- **Final-state diff**: walking `expected_final_state.present` and
  `.absent` against the live sandbox after the agent finishes. Any
  divergence → hard fail.
- **Step count**: if total `sp_*` calls exceed `max_tool_calls`, that's
  recorded as a *note* but not a hard fail on its own — it's a soft
  signal of retry loops or tool-selection confusion.

## Unit-test coverage

The deterministic logic (prompt extraction, stream-json parsing, tool-name
normalisation, scoring, fixture loading, mcp-config writing) is covered
by `tests/unit/harness/test_behavioural_runner.py`. Those tests run in
regular CI and are cheap.

The live-run path (seed-the-sandbox, spawn-claude, verify-final-state,
cleanup) is exercised by `scripts/run-behavioural-harness.sh` manually
or via a dedicated nightly workflow if/when one is added.

## Adding a new scenario

1. Pick a name. Write `scenarios/<name>.md` with a `## User prompt`
   section containing one fenced block.
2. Write `scenarios/<name>.fixture.yaml` — same shape as
   `anqer-reorg.fixture.yaml`. Keep `scratch_root` distinct per
   scenario so parallel runs don't collide.
3. Run `scripts/run-behavioural-harness.sh <name>` and iterate on the
   prompt + scoring rules until passes are reproducible.
4. Add a row to a (future) scenario index in this README.

## What's *not* in scope

- Running this on every PR. Cost + flakiness against the live tenant
  rule that out. Nightly at most.
- Mocking out Claude or Graph. The whole point is the real end-to-end
  loop; mocks would defeat it.
- Multiple LLM providers. We test against the Claude Code client
  specifically because that's what users consume the MCP through. If
  other clients are added, sibling scenarios under a `clients/` subdir
  would be the place.
