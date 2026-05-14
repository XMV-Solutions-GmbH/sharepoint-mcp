<!-- SPDX-License-Identifier: MIT OR Apache-2.0 -->
# Proposals (RFCs / spike notes / architectural decisions)

This folder is where decisions live that are too big to fit in a single GitHub issue. Each proposal is one Markdown file capturing the decision, the reasoning behind it, the alternatives considered, and the eventual outcome.

Use this folder when:

- The change touches multiple components or repos and needs a written design before anyone starts coding.
- A spike resolves a previously open question and the answer is load-bearing for the rest of the project.
- A choice is reversible-but-expensive and you want a clear paper trail of why the project picked option B over A.
- Reviewers want to align on direction *before* PR-level review.
- **An Engineering Principle names a project-specific trade-off** — the Decision Record here is the documented answer for this project (see § 16 of `ENGINEERING_PRINCIPLES.md`).

Do **not** use this folder for:

- Routine work — that goes in [GitHub Issues + the repo's Project board](../../ENGINEERING_PRINCIPLES.md), per § 2 of the engineering principles.
- Per-PR design notes — those go in the PR description.
- Scratch ideas — those go in chat or your own notes until they're ready to be a proposal.

## Naming convention

```text
docs/proposals/YYYY-MM-DD-short-slug.md
```

The date is the date the proposal was *accepted* (or the date drafted, updated to acceptance date when it is accepted). The slug is short and stable — references to the proposal use the slug, not the date.

Examples:

```text
docs/proposals/2026-05-14-harness-tests-in-ci.md
docs/proposals/2026-05-08-non-blocking-login-tools.md
```

## Format

Use the standard template: status, context, decision, alternatives, consequences, implementation notes. At minimum each proposal has:

- **Status** — `Draft` / `Accepted` / `Implemented` / `Superseded by …` / `Withdrawn`.
- **Context** — what's the situation, what constraints apply, what made this decision necessary.
- **Decision** — the chosen approach, in one or two paragraphs.
- **Alternatives considered** — what else was on the table, and why each was not chosen.
- **Consequences** — what changes about the system / process as a result; what trade-offs are now baked in.
- **Implementation notes** — links to the issues / PRs that carried it out (filled in over time).

## Lifecycle

1. **Draft** — open a PR adding the file with `Status: Draft`. Discuss in the PR.
2. **Accepted** — flip status to `Accepted`, merge the PR. The decision is now binding.
3. **Implemented** — once the work to realise the decision has shipped, flip to `Implemented` and link the issues / PRs.
4. **Superseded** — if a later proposal replaces this one, flip to `Superseded by docs/proposals/YYYY-MM-DD-other-slug.md` and leave the file in place. Do not delete history.
5. **Withdrawn** — if the proposal is abandoned without being implemented, flip to `Withdrawn` and explain why. File stays.

Proposals are append-only history. Once accepted, they are not rewritten — they are superseded.

## Index

| Proposal | Status | Summary |
|---|---|---|
| [2026-05-14 Harness tests in CI](2026-05-14-harness-tests-in-ci.md) | Implemented | Run harness tests as a CI gate on every PR |
