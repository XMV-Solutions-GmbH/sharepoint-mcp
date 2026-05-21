<!-- SPDX-License-Identifier: MIT OR Apache-2.0 -->
# Scenario: auth UX — does an uninformed agent format the device code correctly?

Closes part of [#112](https://github.com/XMV-Solutions-GmbH/sharepoint-mcp/issues/112).

The point of the `AGENT_INSTRUCTIONS:` marker in the `sp_auth_begin` /
`sp_auth_status` tool descriptions is to force the agent to render the
`user_code` and `verification_url` in the right shape for a chat UI:
code in a fenced block (one-click copy), URL as a bare auto-link.

This scenario verifies that contract with a deliberately uninformed
agent: `claude --print --bare --mcp-config <local-sp-mcp>` so the
agent has no memory, no skills, no CLAUDE.md context — only what the
MCP tool descriptions say.

## User prompt

```text
How do I sign in to SharePoint with the SharePoint MCP server? Walk me through
the first step. If you need to call a tool, do it. Show me whatever I'd need to
complete sign-in on my phone.
```

## Expected behaviour

The agent should:

1. Call `sp_auth_begin` (the only sensible tool to start a sign-in flow).
2. Receive `user_code` + `verification_url` in the response.
3. Render its reply with:
   - The code inside a fenced code block (one-click copy in chat UIs).
   - The URL as a plain bare URL on its own line (auto-link in chat UIs).
   - **No** prose-embedded code (e.g. "enter the code ABCD-1234 at …").
   - **No** bold wrapping around the URL.
   - **No** paraphrasing of the AGENT_INSTRUCTIONS block.

## Failure modes worth scoring

- Agent calls a different tool first (e.g. `sp_auth_status` when there is no
  pending flow) — soft signal, may indicate description is unclear.
- Agent doesn't fence the code → users have to manually select it.
- Agent wraps URL in `**bold**` or `[text](url)` markdown → kills auto-link.
- Agent embeds the code inline in a sentence → copy-paste includes the prose.
- Agent narrates that "the code is ABCD-1234" before/instead of showing it
  in the block.
