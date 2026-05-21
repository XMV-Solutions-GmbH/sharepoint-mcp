# SPDX-License-Identifier: MIT OR Apache-2.0
# SPDX-FileCopyrightText: 2026 XMV Solutions GmbH
# SPDX-FileContributor: David Koller <david.koller@xmv.de>
"""Unit tests for the behavioural-harness runner's deterministic logic.

The runner itself (live Graph + live Claude) is not exercised here —
that's the whole point of an end-to-end harness. These tests cover the
parsing, scoring, and config-writing helpers that decide pass/fail
once the transcript exists. Run in regular CI; cost zero.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from tests.harness.behavioural.runner import (
    Scenario,
    _extract_prompt_from_markdown,
    _normalise_tool_name,
    load_scenario,
    parse_tool_calls,
    score,
    write_mcp_config,
)

# ── prompt extraction ────────────────────────────────────────────────────


def test_extract_prompt_pulls_first_fenced_block_after_user_prompt_heading() -> None:
    md = """
# Some scenario

Intro text.

## User prompt

```text
Hello agent, do thing one.
Do thing two.
```

## Expected tool sequence

- foo
- bar
"""
    assert _extract_prompt_from_markdown(md) == "Hello agent, do thing one.\nDo thing two."


def test_extract_prompt_accepts_unlabelled_fence() -> None:
    md = "## User prompt\n\n```\nbare block\n```\n"
    assert _extract_prompt_from_markdown(md) == "bare block"


def test_extract_prompt_case_insensitive_heading() -> None:
    md = "## user PROMPT\n\n```\ndo it\n```\n"
    assert _extract_prompt_from_markdown(md) == "do it"


def test_extract_prompt_raises_when_section_missing() -> None:
    with pytest.raises(ValueError, match="User prompt"):
        _extract_prompt_from_markdown("# Scenario\n\nNo prompt here.")


def test_extract_prompt_raises_when_section_has_no_fence() -> None:
    with pytest.raises(ValueError, match="User prompt"):
        _extract_prompt_from_markdown("## User prompt\n\nJust prose, no fence.")


# ── tool-name normalisation ──────────────────────────────────────────────


@pytest.mark.parametrize(
    "raw,expected",
    [
        ("sp_drive_file_upload", "sp_drive_file_upload"),
        ("mcp__sharepoint__sp_drive_file_upload", "sp_drive_file_upload"),
        ("mcp__my-sp-server__sp_search_query", "sp_search_query"),
        ("Bash", "Bash"),  # non-MCP, non-sp → returned as-is
        ("mcp__weird-server__not_sp_prefix", "mcp__weird-server__not_sp_prefix"),
    ],
)
def test_normalise_tool_name(raw: str, expected: str) -> None:
    assert _normalise_tool_name(raw) == expected


# ── stream-json parsing ──────────────────────────────────────────────────


def _assistant_event_with_tool(name: str) -> str:
    return json.dumps(
        {
            "type": "assistant",
            "message": {
                "content": [
                    {"type": "text", "text": "calling tool"},
                    {"type": "tool_use", "name": name, "input": {}, "id": "id"},
                ],
            },
        }
    )


def test_parse_tool_calls_extracts_sp_tools_in_order() -> None:
    lines = [
        json.dumps({"type": "system", "subtype": "init", "tools": []}),
        _assistant_event_with_tool("sp_drive_folder_create"),
        _assistant_event_with_tool("sp_drive_file_upload"),
        _assistant_event_with_tool("sp_drive_file_delete"),
        json.dumps({"type": "result", "subtype": "success"}),
    ]
    assert parse_tool_calls(lines) == [
        "sp_drive_folder_create",
        "sp_drive_file_upload",
        "sp_drive_file_delete",
    ]


def test_parse_tool_calls_strips_mcp_server_prefix() -> None:
    lines = [
        _assistant_event_with_tool("mcp__sharepoint__sp_search_query"),
    ]
    assert parse_tool_calls(lines) == ["sp_search_query"]


def test_parse_tool_calls_ignores_non_sp_tools() -> None:
    """Bash/Read/Write/etc. are out of scope — only sp_* matters for scoring."""
    lines = [
        _assistant_event_with_tool("Bash"),
        _assistant_event_with_tool("Read"),
        _assistant_event_with_tool("sp_drive_file_read"),
    ]
    assert parse_tool_calls(lines) == ["sp_drive_file_read"]


def test_parse_tool_calls_keeps_duplicates() -> None:
    """Call count matters — duplicates feed into max_tool_calls scoring."""
    lines = [
        _assistant_event_with_tool("sp_drive_file_delete"),
        _assistant_event_with_tool("sp_drive_file_delete"),
        _assistant_event_with_tool("sp_drive_file_delete"),
    ]
    assert parse_tool_calls(lines) == ["sp_drive_file_delete"] * 3


def test_parse_tool_calls_ignores_garbage_lines() -> None:
    lines = [
        "",
        "  ",
        "not json",
        _assistant_event_with_tool("sp_drive_file_read"),
    ]
    assert parse_tool_calls(lines) == ["sp_drive_file_read"]


def test_parse_tool_calls_ignores_text_only_assistant_messages() -> None:
    lines = [
        json.dumps(
            {
                "type": "assistant",
                "message": {"content": [{"type": "text", "text": "thinking..."}]},
            }
        ),
    ]
    assert parse_tool_calls(lines) == []


# ── scoring ──────────────────────────────────────────────────────────────


def _scenario(
    *,
    expected: list[str] | None = None,
    prohibited: list[str] | None = None,
    max_calls: int = 25,
    text_must_match: list[str] | None = None,
    text_must_not_match: list[str] | None = None,
    skip_sandbox: bool = False,
) -> Scenario:
    return Scenario(
        name="t",
        prompt="",
        site_url="https://example.sharepoint.com/sites/x",
        scratch_root="scratch",
        local_files={},
        initial_state={},
        expected_final_state={},
        expected_tools_at_least_once=expected or [],
        prohibited_tools=prohibited or [],
        max_tool_calls=max_calls,
        assistant_text_must_match=text_must_match or [],
        assistant_text_must_not_match=text_must_not_match or [],
        skip_sandbox=skip_sandbox,
    )


def test_score_pass_all_expected_called_no_prohibited_no_diff() -> None:
    s = _scenario(expected=["sp_drive_file_move", "sp_drive_file_upload"])
    result = score(["sp_drive_file_move", "sp_drive_file_upload"], [], s)
    assert result.passed is True
    assert result.expected_missing == []
    assert result.prohibited_called == []


def test_score_fails_when_expected_tool_missing() -> None:
    s = _scenario(expected=["sp_drive_file_move", "sp_drive_file_upload"])
    result = score(["sp_drive_file_upload"], [], s)
    assert result.passed is False
    assert result.expected_missing == ["sp_drive_file_move"]


def test_score_fails_when_prohibited_tool_called() -> None:
    s = _scenario(
        expected=["sp_drive_file_move"],
        prohibited=["sp_list_item_delete"],
    )
    result = score(["sp_drive_file_move", "sp_list_item_delete"], [], s)
    assert result.passed is False
    assert result.prohibited_called == ["sp_list_item_delete"]


def test_score_fails_on_final_state_diff() -> None:
    s = _scenario(expected=["sp_drive_file_move"])
    result = score(
        ["sp_drive_file_move"],
        ["expected present: foo.md"],
        s,
    )
    assert result.passed is False
    assert result.final_state_diff == ["expected present: foo.md"]


def test_score_emits_note_when_step_count_exceeds_cap() -> None:
    """Step-count over-cap is a SOFT signal — surfaced in notes, not a hard fail
    on its own. The hard fail comes from missing-expected / prohibited / state-diff."""
    s = _scenario(expected=[], max_calls=3)
    result = score(["sp_drive_file_read"] * 10, [], s)
    assert any("step count" in n for n in result.notes)
    # No expected, no prohibited, no diff → still passes despite the note.
    assert result.passed is True


# ── fixture loading ──────────────────────────────────────────────────────


def test_load_scenario_round_trip(tmp_path: Path) -> None:
    md = tmp_path / "demo.md"
    md.write_text(
        "## User prompt\n\n```text\nReorganise something.\n```\n",
        encoding="utf-8",
    )
    fixture = tmp_path / "demo.fixture.yaml"
    fixture.write_text(
        """
site_url: https://x.sharepoint.com/sites/y
scratch_root: scratch
local_files: {}
initial_state:
  files: []
expected_final_state:
  present: []
  absent: []
expected_tools_at_least_once:
  - sp_drive_file_move
prohibited_tools:
  - sp_list_item_delete
max_tool_calls: 17
""".strip(),
        encoding="utf-8",
    )
    scenario = load_scenario(tmp_path, "demo")
    assert scenario.name == "demo"
    assert scenario.prompt == "Reorganise something."
    assert scenario.scratch_root == "scratch"
    assert scenario.expected_tools_at_least_once == ["sp_drive_file_move"]
    assert scenario.prohibited_tools == ["sp_list_item_delete"]
    assert scenario.max_tool_calls == 17


def test_load_scenario_raises_when_md_missing(tmp_path: Path) -> None:
    (tmp_path / "demo.fixture.yaml").write_text("site_url: x\nscratch_root: y\n")
    with pytest.raises(FileNotFoundError, match=r"\.md"):
        load_scenario(tmp_path, "demo")


def test_load_scenario_raises_when_fixture_missing(tmp_path: Path) -> None:
    (tmp_path / "demo.md").write_text("## User prompt\n\n```\nx\n```\n")
    with pytest.raises(FileNotFoundError, match="fixture"):
        load_scenario(tmp_path, "demo")


def test_anqer_reorg_scenario_loads_from_repo() -> None:
    """The shipped scenario is parseable end-to-end."""
    scenarios_dir = Path(__file__).resolve().parents[2] / "harness" / "behavioural" / "scenarios"
    s = load_scenario(scenarios_dir, "anqer-reorg")
    assert s.site_url.endswith("/sites/sharepoint-mcp-harness")
    assert "scratch" in s.scratch_root
    assert "sp_drive_file_move" in s.expected_tools_at_least_once
    assert "sp_list_item_delete" in s.prohibited_tools
    assert s.prompt.startswith("We need to reorganise")


# ── mcp config writer ────────────────────────────────────────────────────


def test_write_mcp_config_points_at_local_source_tree(tmp_path: Path) -> None:
    """The whole point: don't run uvx mcp-server-sharepoint (PyPI). Run the
    local source tree's binary via `uv run --project <repo>`."""
    repo = tmp_path / "fake-repo"
    repo.mkdir()
    config_path = write_mcp_config(
        repo_root=repo,
        profile="harness",
        allow_writes=True,
    )
    cfg = json.loads(config_path.read_text())
    server = cfg["mcpServers"]["sharepoint"]
    assert server["command"] == "uv"
    assert "run" in server["args"]
    assert "--project" in server["args"]
    assert str(repo) in server["args"]
    assert "mcp-server-sharepoint" in server["args"]
    # NEVER uvx — that fetches the published version.
    assert "uvx" not in server["args"]
    assert "uvx" not in server["command"]
    assert server["env"]["SP_PROFILE"] == "harness"
    assert server["env"]["SP_ALLOW_WRITES"] == "true"


def test_write_mcp_config_writes_false_when_writes_disabled(tmp_path: Path) -> None:
    config_path = write_mcp_config(
        repo_root=tmp_path,
        profile="readonly",
        allow_writes=False,
    )
    cfg = json.loads(config_path.read_text())
    assert cfg["mcpServers"]["sharepoint"]["env"]["SP_ALLOW_WRITES"] == "false"


# ── parse_assistant_text + text-content scoring ──────────────────────────


def _assistant_text_event(text: str) -> str:
    return json.dumps(
        {
            "type": "assistant",
            "message": {"content": [{"type": "text", "text": text}]},
        }
    )


def test_parse_assistant_text_concatenates_text_blocks_only() -> None:
    """Tool_use blocks should be ignored; only `text` blocks contribute."""
    from tests.harness.behavioural.runner import parse_assistant_text

    lines = [
        _assistant_text_event("First sentence."),
        _assistant_event_with_tool("sp_auth_begin"),
        _assistant_text_event("Second sentence."),
    ]
    out = parse_assistant_text(lines)
    assert "First sentence." in out
    assert "Second sentence." in out
    assert "sp_auth_begin" not in out


def test_score_passes_when_required_text_pattern_present() -> None:
    s = _scenario(
        expected=[],
        text_must_match=[r"```\nABC-123\n```"],
        skip_sandbox=True,
    )
    result = score(
        [],
        [],
        s,
        assistant_text="Here you go:\n```\nABC-123\n```\nclick the link",
    )
    assert result.passed is True


def test_score_fails_when_required_text_pattern_missing() -> None:
    s = _scenario(text_must_match=[r"```\n[A-Z0-9-]+\n```"], skip_sandbox=True)
    result = score(
        [],
        [],
        s,
        assistant_text="The code is ABC-123 — please go to login.microsoft.com",
    )
    assert result.passed is False
    assert any("```" in p for p in result.text_missing_patterns)


def test_score_fails_when_forbidden_text_pattern_present() -> None:
    """Bold-wrapped URL kills auto-link in most chat UIs — forbidden."""
    s = _scenario(text_must_not_match=[r"\*\*\s*https?://"], skip_sandbox=True)
    result = score(
        [],
        [],
        s,
        assistant_text="Go to **https://login.microsoftonline.com** and sign in.",
    )
    assert result.passed is False
    assert result.text_forbidden_patterns == [r"\*\*\s*https?://"]


def test_score_text_patterns_use_multiline_and_dotall() -> None:
    """`(?m)^https?://` should match a URL at the start of a later line — verifies
    the runner passes MULTILINE/DOTALL to re.search."""
    s = _scenario(text_must_match=[r"(?m)^https?://"], skip_sandbox=True)
    result = score(
        [],
        [],
        s,
        assistant_text="Here's the link:\nhttps://example.com/auth",
    )
    assert result.passed is True


def test_score_combines_tool_and_text_failures() -> None:
    """A scenario that misses both the expected tool and a required text
    pattern reports both reasons — the user can fix in one round."""
    s = _scenario(
        expected=["sp_auth_begin"],
        text_must_match=[r"```"],
        skip_sandbox=True,
    )
    result = score([], [], s, assistant_text="some prose")
    assert result.passed is False
    assert result.expected_missing == ["sp_auth_begin"]
    assert result.text_missing_patterns == [r"```"]
