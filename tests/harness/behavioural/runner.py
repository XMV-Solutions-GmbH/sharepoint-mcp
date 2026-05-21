# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Behavioural-harness runner — does an LLM, given the full sp_* catalog,
pick the right tools for a realistic SharePoint task?

The pytest harness (`tests/harness/test_*.py`) calls each tool directly
against real Graph. That's a *functional* harness: each tool works
in isolation. This module is the *behavioural* layer on top: it boots
the MCP server in front of a real Claude agent via `claude --print
--output-format=stream-json`, gives the agent a scripted prompt, and
inspects the tool-use events in the resulting transcript.

Architecture:

    scenarios/<name>.md            → user prompt + scoring criteria (human-readable)
    scenarios/<name>.fixture.yaml  → seed state + expected final state + scoring rules

The runner:

    1. Reads both files.
    2. Seeds the harness sandbox via direct sp_* impl calls (sp_drive_file_upload).
    3. Writes a temp .mcp.json that points at the **local source tree's**
       MCP server, not at PyPI's. This is the whole point — we're
       validating the in-development version.
    4. Spawns `claude --print --verbose --output-format=stream-json
       --mcp-config <temp> --permission-mode=bypassPermissions
       "<prompt>"` and pipes its stdout into a streaming parser.
    5. Extracts every `tool_use` event with name starting with `sp_`
       (or `mcp__sharepoint__sp_*` if the MCP-client wraps them).
    6. Verifies the final state of the sandbox by walking the
       `expected_final_state.present` and `.absent` paths.
    7. Scores: hard fail if a prohibited tool was called, an expected
       tool was missed, or final state diverges. Soft fail if step
       count exceeds `max_tool_calls`.
    8. Cleans up: deletes the scratch subtree and any local fixture
       files. Best-effort; runs in a `finally` block so a crashed
       agent run doesn't poison subsequent scenarios.

The runner is intentionally **not a pytest test**. It runs end-to-end
against real Graph, real Claude API (your account), and a real
harness sandbox — costs real money per invocation. The CI integration,
if any, runs it nightly via a dedicated workflow with explicit budget
guards.

Unit tests for the deterministic parts (transcript parsing, fixture
loading, scoring) live in `tests/unit/harness/test_behavioural_runner.py`
and run in regular CI.
"""

from __future__ import annotations

import argparse
import dataclasses
import json
import re
import subprocess
import sys
import tempfile
from collections.abc import Iterable
from pathlib import Path
from typing import Any

import yaml


@dataclasses.dataclass(frozen=True)
class Scenario:
    """Loaded scenario: prompt from the .md, scoring rules from the .yaml."""

    name: str
    prompt: str
    site_url: str
    scratch_root: str
    local_files: dict[str, str]
    initial_state: dict[str, Any]
    expected_final_state: dict[str, Any]
    expected_tools_at_least_once: list[str]
    prohibited_tools: list[str]
    max_tool_calls: int
    # Text-content scoring (used by auth-UX-style scenarios where the
    # observable isn't "what file ended up in the sandbox" but "how did the
    # agent format its reply to the user"). Each entry is a regex matched
    # against the concatenated assistant-text content blocks from the
    # transcript. `must_match` → hard fail if missing; `must_not_match` →
    # hard fail if present.
    assistant_text_must_match: list[str]
    assistant_text_must_not_match: list[str]
    # When True, skip the sandbox seed/cleanup/state-verify path entirely.
    # For scenarios where the test is purely about agent narration shape.
    skip_sandbox: bool


@dataclasses.dataclass(frozen=True)
class ScoreResult:
    """Outcome of a single behavioural run."""

    passed: bool
    tool_calls: list[str]
    expected_missing: list[str]
    prohibited_called: list[str]
    final_state_diff: list[str]
    text_missing_patterns: list[str]
    text_forbidden_patterns: list[str]
    notes: list[str]


# ── scenario loading ──────────────────────────────────────────────────────


def _extract_prompt_from_markdown(md: str) -> str:
    """Pull the user-facing prompt out of `## User prompt` → first fenced block."""
    match = re.search(
        r"##\s+User\s+prompt\s*\n+```(?:\w+)?\n(.*?)```",
        md,
        re.DOTALL | re.IGNORECASE,
    )
    if not match:
        raise ValueError(
            "Scenario .md must contain a '## User prompt' section followed by a "
            "fenced code block with the prompt body.",
        )
    return match.group(1).strip()


def load_scenario(scenario_dir: Path, name: str) -> Scenario:
    """Load `<scenario_dir>/<name>.md` and `.fixture.yaml` into a Scenario."""
    md_path = scenario_dir / f"{name}.md"
    yaml_path = scenario_dir / f"{name}.fixture.yaml"
    if not md_path.exists():
        raise FileNotFoundError(f"scenario .md not found: {md_path}")
    if not yaml_path.exists():
        raise FileNotFoundError(f"scenario fixture not found: {yaml_path}")

    fixture = yaml.safe_load(yaml_path.read_text(encoding="utf-8"))
    prompt = _extract_prompt_from_markdown(md_path.read_text(encoding="utf-8"))

    return Scenario(
        name=name,
        prompt=prompt,
        site_url=str(fixture.get("site_url") or ""),
        scratch_root=str(fixture.get("scratch_root") or ""),
        local_files={str(k): str(v) for k, v in (fixture.get("local_files") or {}).items()},
        initial_state=fixture.get("initial_state") or {},
        expected_final_state=fixture.get("expected_final_state") or {},
        expected_tools_at_least_once=list(fixture.get("expected_tools_at_least_once") or []),
        prohibited_tools=list(fixture.get("prohibited_tools") or []),
        max_tool_calls=int(fixture.get("max_tool_calls") or 50),
        assistant_text_must_match=list(fixture.get("assistant_text_must_match") or []),
        assistant_text_must_not_match=list(fixture.get("assistant_text_must_not_match") or []),
        skip_sandbox=bool(fixture.get("skip_sandbox") or False),
    )


# ── stream-json parsing ───────────────────────────────────────────────────


# `claude --print --output-format=stream-json` emits one JSON object per line.
# We only care about `tool_use` events in assistant messages. The tool name
# may appear bare (`sp_drive_file_upload`) or namespaced by the MCP-client
# (`mcp__sharepoint__sp_drive_file_upload`); we accept both and normalise.

_MCP_TOOL_PREFIX_RE = re.compile(r"^mcp__[a-zA-Z0-9_-]+__(sp_[a-zA-Z0-9_]+)$")


def _normalise_tool_name(name: str) -> str:
    """Strip an `mcp__<server>__` prefix if present.

    Claude Code wraps MCP tool names like `mcp__sharepoint__sp_drive_file_read`.
    The fixture's `expected_tools_at_least_once` list uses the bare names,
    so we normalise before comparison.
    """
    match = _MCP_TOOL_PREFIX_RE.match(name)
    return match.group(1) if match else name


def parse_tool_calls(stream_json_lines: Iterable[str]) -> list[str]:
    """Pull every `sp_*` tool-use name out of a stream-json transcript.

    Returns the call sequence in the order the agent issued them. Duplicates
    are kept — call count matters for the `max_tool_calls` cap.
    """
    calls: list[str] = []
    for raw_line in stream_json_lines:
        line = raw_line.strip()
        if not line:
            continue
        try:
            event = json.loads(line)
        except json.JSONDecodeError:
            continue
        if event.get("type") != "assistant":
            continue
        content = event.get("message", {}).get("content") or []
        for block in content:
            if not isinstance(block, dict):
                continue
            if block.get("type") != "tool_use":
                continue
            name = _normalise_tool_name(str(block.get("name") or ""))
            if name.startswith("sp_"):
                calls.append(name)
    return calls


def parse_assistant_text(stream_json_lines: Iterable[str]) -> str:
    """Concatenate every assistant-text content block from the transcript.

    Used by text-shape scoring: did the agent emit a fenced code block?
    A bare URL? The whole stream is joined into one string so multi-message
    replies still get matched against the patterns as a unit.
    """
    chunks: list[str] = []
    for raw_line in stream_json_lines:
        line = raw_line.strip()
        if not line:
            continue
        try:
            event = json.loads(line)
        except json.JSONDecodeError:
            continue
        if event.get("type") != "assistant":
            continue
        content = event.get("message", {}).get("content") or []
        for block in content:
            if not isinstance(block, dict):
                continue
            if block.get("type") == "text":
                chunks.append(str(block.get("text") or ""))
    return "\n".join(chunks)


# ── seeding + cleanup ─────────────────────────────────────────────────────


def seed_sandbox(scenario: Scenario, *, profile: str) -> None:
    """Place the scenario's initial state into the harness sandbox.

    Uses `sp_drive_file_upload` (which already creates missing parents
    recursively) so seeding goes through the same code path the agent
    exercises. Local files referenced by the prompt are written to disk.
    """
    # Local-disk seeds first (synchronous, can't fail Graph-side).
    for local_path, content in scenario.local_files.items():
        p = Path(local_path)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(content, encoding="utf-8")

    # Imports deferred so unit tests can load this module without the full
    # sharepoint_mcp install pulling in Graph helpers.
    from sharepoint_mcp.tools.publish import publish

    scratch_url = f"{scenario.site_url}/Shared Documents/{scenario.scratch_root}"
    for entry in scenario.initial_state.get("files", []):
        # Each seed file is uploaded as a fresh temp file; sp_drive_file_upload
        # reads from disk, so we materialise content there first.
        rel = str(entry["path"])
        content = str(entry.get("content", ""))
        target_folder = scratch_url
        filename = rel
        if "/" in rel:
            parent, _, filename = rel.rpartition("/")
            target_folder = f"{scratch_url}/{parent}"
        with tempfile.NamedTemporaryFile(
            mode="w",
            suffix=f"-{filename}",
            delete=False,
            encoding="utf-8",
        ) as tf:
            tf.write(content)
            tmp_path = tf.name
        publish(tmp_path, target_folder, name=filename, profile=profile)


def cleanup_sandbox(scenario: Scenario, *, profile: str) -> None:
    """Delete the scratch subtree. Best-effort — swallows missing-folder 404s."""
    try:
        from sharepoint_mcp.tools.delete_file import delete_file
    except ImportError:
        return
    try:
        delete_file(scenario.site_url, scenario.scratch_root, profile=profile)
    except Exception:
        pass


# ── final-state verification ──────────────────────────────────────────────


def verify_final_state(scenario: Scenario, *, profile: str) -> list[str]:
    """Walk expected present/absent paths against the live sandbox.

    Returns a list of diff descriptions (empty list = all good).
    """
    from sharepoint_mcp.tools.list_folder import list_folder

    diffs: list[str] = []
    scratch_url = f"{scenario.site_url}/Shared Documents/{scenario.scratch_root}"
    present = scenario.expected_final_state.get("present") or []
    absent = scenario.expected_final_state.get("absent") or []

    # We build a flat set of all paths the runner can see by recursively
    # listing the scratch root. Cheap for harness sandbox scale; expensive
    # for production trees but that's not what this runs against.
    found: set[str] = set()

    def _walk(folder_url: str, prefix: str) -> None:
        try:
            children = list_folder(folder_url, profile=profile)
        except Exception as exc:
            diffs.append(f"could not list {folder_url}: {exc}")
            return
        for child in children:
            name = str(child.get("name") or "")
            kind = str(child.get("type") or "")
            rel = f"{prefix}/{name}" if prefix else name
            found.add(rel)
            if kind == "folder":
                _walk(f"{folder_url}/{name}", rel)

    _walk(scratch_url, "")

    for entry in present:
        path = str(entry["path"])
        if path not in found:
            diffs.append(f"expected present: {path}")
    for entry in absent:
        path = str(entry["path"])
        if path in found:
            diffs.append(f"expected absent (still present): {path}")
    return diffs


# ── scoring ───────────────────────────────────────────────────────────────


def score(
    tool_calls: list[str],
    final_state_diff: list[str],
    scenario: Scenario,
    *,
    assistant_text: str = "",
) -> ScoreResult:
    """Apply pass/fail rules to the observed tool-use sequence, sandbox state,
    and (optionally) the agent's narration."""
    called = set(tool_calls)
    expected_missing = [t for t in scenario.expected_tools_at_least_once if t not in called]
    prohibited_called = [t for t in scenario.prohibited_tools if t in called]

    text_missing_patterns: list[str] = []
    text_forbidden_patterns: list[str] = []
    for pattern in scenario.assistant_text_must_match:
        if not re.search(pattern, assistant_text, re.MULTILINE | re.DOTALL):
            text_missing_patterns.append(pattern)
    for pattern in scenario.assistant_text_must_not_match:
        if re.search(pattern, assistant_text, re.MULTILINE | re.DOTALL):
            text_forbidden_patterns.append(pattern)

    notes: list[str] = []
    if len(tool_calls) > scenario.max_tool_calls:
        notes.append(
            f"step count {len(tool_calls)} exceeds max_tool_calls={scenario.max_tool_calls} "
            "(likely retry loop or tool-selection confusion)"
        )

    passed = (
        not expected_missing
        and not prohibited_called
        and not final_state_diff
        and not text_missing_patterns
        and not text_forbidden_patterns
    )

    return ScoreResult(
        passed=passed,
        tool_calls=tool_calls,
        expected_missing=expected_missing,
        prohibited_called=prohibited_called,
        final_state_diff=final_state_diff,
        text_missing_patterns=text_missing_patterns,
        text_forbidden_patterns=text_forbidden_patterns,
        notes=notes,
    )


# ── claude invocation ─────────────────────────────────────────────────────


def write_mcp_config(*, repo_root: Path, profile: str, allow_writes: bool) -> Path:
    """Materialise a temp .mcp.json that points the MCP server at the local source tree.

    Critical: we run `uv run --project <repo> mcp-server-sharepoint`, not
    `uvx mcp-server-sharepoint` — the latter would pull the published
    PyPI version, which defeats the whole purpose of testing the
    in-development branch.
    """
    config = {
        "mcpServers": {
            "sharepoint": {
                "command": "uv",
                "args": [
                    "run",
                    "--project",
                    str(repo_root),
                    "mcp-server-sharepoint",
                ],
                "env": {
                    "SP_PROFILE": profile,
                    "SP_ALLOW_WRITES": "true" if allow_writes else "false",
                },
            },
        },
    }
    handle = tempfile.NamedTemporaryFile(
        mode="w",
        suffix=".mcp.json",
        delete=False,
        encoding="utf-8",
    )
    json.dump(config, handle, indent=2)
    handle.close()
    return Path(handle.name)


def run_claude(
    prompt: str,
    *,
    mcp_config: Path,
    timeout_seconds: int = 600,
) -> list[str]:
    """Spawn `claude --print --output-format=stream-json` and return its stdout lines."""
    cmd = [
        "claude",
        "--print",
        "--verbose",
        "--output-format=stream-json",
        "--mcp-config",
        str(mcp_config),
        "--permission-mode=bypassPermissions",
        prompt,
    ]
    proc = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        timeout=timeout_seconds,
        check=False,
    )
    if proc.returncode != 0:
        sys.stderr.write(
            f"claude exited non-zero ({proc.returncode}); stderr:\n{proc.stderr}\n",
        )
    return proc.stdout.splitlines()


# ── entrypoint ────────────────────────────────────────────────────────────


def main(argv: list[str] | None = None) -> int:
    """CLI entrypoint. Returns the process exit code (0 = pass)."""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--scenario",
        required=True,
        help="Scenario name (without .md / .fixture.yaml suffix).",
    )
    parser.add_argument(
        "--scenarios-dir",
        type=Path,
        default=Path(__file__).parent / "scenarios",
        help="Directory holding <scenario>.md and <scenario>.fixture.yaml.",
    )
    parser.add_argument(
        "--profile",
        default="harness",
        help="SP_PROFILE to use for seeding/cleanup + the agent's MCP env.",
    )
    parser.add_argument(
        "--repo-root",
        type=Path,
        default=Path(__file__).resolve().parents[3],
        help="Path to the sharepoint-mcp source tree (for uv run --project).",
    )
    parser.add_argument(
        "--skip-seed",
        action="store_true",
        help="Skip seeding the sandbox (use when re-running against existing state).",
    )
    parser.add_argument(
        "--skip-cleanup",
        action="store_true",
        help="Skip cleanup (leave the scratch tree for post-mortem inspection).",
    )
    args = parser.parse_args(argv)

    scenario = load_scenario(args.scenarios_dir, args.scenario)
    print(f"[harness] scenario: {scenario.name}", file=sys.stderr)
    print(f"[harness] site: {scenario.site_url}", file=sys.stderr)
    print(f"[harness] scratch root: {scenario.scratch_root}", file=sys.stderr)

    if not args.skip_seed and not scenario.skip_sandbox:
        print("[harness] seeding sandbox …", file=sys.stderr)
        seed_sandbox(scenario, profile=args.profile)

    mcp_config = write_mcp_config(
        repo_root=args.repo_root,
        profile=args.profile,
        allow_writes=True,
    )
    try:
        print("[harness] running claude --print …", file=sys.stderr)
        lines = run_claude(scenario.prompt, mcp_config=mcp_config)
        tool_calls = parse_tool_calls(lines)
        assistant_text = parse_assistant_text(lines)
        print(f"[harness] observed {len(tool_calls)} sp_* tool calls", file=sys.stderr)

        diff = (
            [] if scenario.skip_sandbox
            else verify_final_state(scenario, profile=args.profile)
        )
        result = score(tool_calls, diff, scenario, assistant_text=assistant_text)
    finally:
        if not args.skip_cleanup and not scenario.skip_sandbox:
            print("[harness] cleaning up …", file=sys.stderr)
            cleanup_sandbox(scenario, profile=args.profile)
        try:
            mcp_config.unlink()
        except FileNotFoundError:
            pass

    # Report
    print(f"\n=== behavioural harness — {scenario.name} ===")
    print(f"passed:                {result.passed}")
    print(f"tool calls observed:   {len(result.tool_calls)}")
    if result.tool_calls:
        print(f"  sequence: {' → '.join(result.tool_calls)}")
    if result.expected_missing:
        print(f"expected but missing:  {result.expected_missing}")
    if result.prohibited_called:
        print(f"prohibited but called: {result.prohibited_called}")
    if result.final_state_diff:
        print("final-state diff:")
        for line in result.final_state_diff:
            print(f"  - {line}")
    if result.text_missing_patterns:
        print("assistant text — required patterns missing:")
        for pattern in result.text_missing_patterns:
            print(f"  - {pattern}")
    if result.text_forbidden_patterns:
        print("assistant text — forbidden patterns matched:")
        for pattern in result.text_forbidden_patterns:
            print(f"  - {pattern}")
    if result.notes:
        print("notes:")
        for note in result.notes:
            print(f"  - {note}")

    return 0 if result.passed else 1


if __name__ == "__main__":
    sys.exit(main())
