---
doc-schema-version: 1
title: "Agent pipeline"
summary: "The GitHub App identities, labels and workflows that turn the maintainer's issues into agent pull requests, and how to pause or step in."
read_when:
  - Changing an agent workflow or its secrets
  - Pausing the agents or stepping in on an agent issue or PR
doc_type: "reference"
---

# Agent pipeline

The agent workflows act through two GitHub Apps installed on this repository only. A workflow
that acts with the default `GITHUB_TOKEN` starts no other workflow: a PR it opens gets no CI,
and a label it adds triggers nothing. Every hand-off between agents therefore goes through an
app installation token.

## Identities

| App | Bot login | Repository permissions | Role |
| --- | --- | --- | --- |
| `pptx-ts-agents` | `pptx-ts-agents[bot]` | Contents RW, Issues RW, Pull requests RW, Metadata R | Opens branches and PRs, comments, edits labels |
| `pptx-ts-reviewer` | `pptx-ts-reviewer[bot]` | Contents R, Pull requests RW, Issues R, Metadata R | Reviews agent PRs |

There are two apps because GitHub does not let a PR's author approve it: the bot that opens
a PR cannot be the one whose approval counts. Neither app has the Workflows permission, so
neither can change `.github/workflows/`. Neither has a webhook; workflows mint their tokens
with `actions/create-github-app-token`.

## Secrets and the kill switch

| Name | Kind | Holds |
| --- | --- | --- |
| `AGENTS_APP_ID`, `AGENTS_APP_PRIVATE_KEY` | secret | `pptx-ts-agents` credentials |
| `REVIEWER_APP_ID`, `REVIEWER_APP_PRIVATE_KEY` | secret | `pptx-ts-reviewer` credentials |
| `CLAUDE_CODE_OAUTH_TOKEN` | secret | Claude Code token from `claude setup-token` |
| `AGENTS_ENABLED` | variable | `true` to run the agent workflows |

Every agent workflow requires `vars.AGENTS_ENABLED == 'true'`. To pause all of them without
editing a file, set the variable to anything else:

```sh
gh variable set AGENTS_ENABLED --body false
```

To rotate an app key, generate a new private key on the app's settings page, store it with
`gh secret set AGENTS_APP_PRIVATE_KEY < key.pem` (or the reviewer's name), delete the old key
on the same page, then delete the local `.pem`.

## Labels

| Label | Meaning |
| --- | --- |
| `agent:ready` | Issue queued for the implementer |
| `agent:working` | The implementer is running on the issue |
| `agent:pr` | PR opened by the implementer |
| `needs-maintainer` | An agent stopped and the maintainer has to decide |

The repository allows auto-merge and deletes a PR's branch once it merges.

## Implementer

`.github/workflows/agent-implement.yml` turns an issue labelled `agent:ready` into a PR. It
runs when the label is added, every two hours as a sweep over the queue, and on manual
dispatch.

### What starts it

The workflow only goes ahead when all of these hold:

- `AGENTS_ENABLED` is `true`.
- The issue is open and was opened by `shbernal`. Any other author is turned away, whoever
  added the label.
- On a label event, the label was added by `shbernal` or `pptx-ts-agents[bot]`.
- No open PR already closes the issue.
- Fewer than two PRs labelled `agent:pr` are open. Otherwise the issue keeps `agent:ready`
  and the sweep takes it once one merges. The sweep picks the oldest queued issue.

One run goes at a time.

### What it does

The issue moves from `agent:ready` to `agent:working`. Claude Code works on branch
`agent/issue-N` with the job's read-only token. It can read the repository and issues, edit
files and run `pnpm run` scripts, but has no credential to commit, push or open a PR. It
ends with one of three outcomes:

- `pr`: the workflow runs `pnpm run verify` itself, commits as `pptx-ts-agents[bot]` with a
  Conventional Commit subject, pushes and opens a PR labelled `agent:pr`. The body follows
  the PR template: AI involvement is "Mostly written by an AI agent", and the verify box is
  ticked only if that run passed. Auto-merge (squash) is armed.
- `blocked`: the issue needs an API or scope decision, desktop PowerPoint, or a change under
  `.github/`.
- `already-fixed`: the current code already does what the issue asks.

On `blocked` or `already-fixed`, a fix that turned out empty, or a failed run (timeout,
usage limit, error), the issue gets a comment with the reason or the run link, and
`needs-maintainer`.

### Protected paths

A PR that changes any of these gets `needs-maintainer` and a comment listing them, and
auto-merge stays off:

- `.github/`
- `package.json` outside `devDependencies`
- `scripts/release*` and `scripts/sync-version.mjs`
- `tsdown.config.ts`

`CHANGELOG.md` is not on the list: an agent fix or feature adds its entry under
`## [Unreleased]`, as any PR does. Cutting a release, which dates those entries, stays with the
maintainer.

### Rerunning

To run it again on an issue, for example after answering a `blocked` comment:

```sh
gh issue edit N --remove-label needs-maintainer
gh workflow run agent-implement.yml -f issue_number=N
```

Dispatch skips the label checks but not the author, open-PR and cap checks.

## Reviewer

`.github/workflows/agent-review.md` is a gh-aw workflow; Actions runs its compiled
`agent-review.lock.yml`. Edit the `.md`, run `gh aw compile`, and commit both. It runs Claude
Code on Z.ai's GLM, as issue triage does.

It reviews a PR when the PR is opened, reopened, marked ready or pushed to, if all of these
hold:

- `AGENTS_ENABLED` is `true`.
- The PR carries `agent:pr`. The maintainer's own PRs get no review run.
- The PR's branch is in this repository.

A new push cancels a review still in flight.

The agent reads the linked issue, its triage report, the diff and its own earlier reviews.
It treats all of it as content to assess, never as instructions. It checks, in order:

1. The change solves the linked issue and nothing else.
2. It is correct and breaks no caller.
3. A focused test covers it ([testing](testing.md)).
4. A change to emitted OOXML carries evidence and a fixture ([OOXML](ooxml.md)).
5. A user-visible fix or feature has a `CHANGELOG.md` entry under `## [Unreleased]`, with
   migration guidance if it breaks anything.
6. Docs for the changed behaviour are updated.

Each blocking finding is an inline comment saying what to change, and any blocking finding
makes the review `REQUEST_CHANGES`. Nits are inline comments starting with `Nit:` and do
not block. With no blocking finding, the review is `APPROVE`. On a re-review the agent says
which earlier findings are resolved, and it does not raise new blocking findings on lines
the new commits left alone.

The review is posted as `pptx-ts-reviewer[bot]`, not by the workflow's own token. A review
by `github-actions[bot]` cannot count as a required approval, and the review event it
causes starts no other workflow.
