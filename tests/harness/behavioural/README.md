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

## What it looks like

Per scenario, three pieces:

1. **`scenarios/<name>.md`** — the user task in natural language,
   plus a "expected tool sequence" section listing the tools the LLM
   *should* pick (minimum set; the LLM may use additional reads).
2. **`scenarios/<name>.fixture.yaml`** — the initial state in the
   harness tenant (folders + files) and the final expected state.
   The runner seeds + asserts.
3. The runner (TODO: `scripts/run-behavioural-harness.sh`) boots
   `claude --print` with `--mcp-config` pointing at a config that
   uses the local branch's MCP server (not the published PyPI one),
   pipes the scenario prompt, captures the transcript, parses tool
   calls from it, asserts.

## Scoring

Each run produces:

- **Tool-selection accuracy**: did the agent pick the documented
  expected tools? Extra reads are fine; wrong-tool selections fail.
- **Step count**: how many Graph calls did it take? Compare against
  the minimum for the scenario.
- **Final-state check**: did the harness sandbox end in the expected
  state (folders, files, content, metadata)?
- **Confusion incidents**: did the agent try a non-existent tool,
  hit a "wrong category" error, or get stuck in a retry loop?

## Why it's not gating v0.7.0

The pytest harness covers *every* tool against real Graph. That's
strong. The behavioural harness is the *next* test layer, and
implementing the runner needs scaffolding (a fresh Claude Code
process, MCP config that points at the editable local install, a
JSONL transcript parser) that's a meaningful effort by itself.

For v0.7.0, this directory ships with:

- The scenario file format and one concrete scenario (`anqer-reorg.md`)
- Scoring criteria documented
- A tracking issue for the runner implementation

See the v0.7.0 master ticket for the follow-up reference.
