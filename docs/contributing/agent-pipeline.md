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
| `pptx-ts-reviewer` | `pptx-ts-reviewer[bot]` | Contents R, Pull requests RW, Issues RW, Metadata R | Reviews agent PRs |

There are two apps because GitHub does not let a PR's author approve it: the bot that opens
a PR cannot be the one whose approval counts. Neither app has the Workflows permission, so
neither can change `.github/workflows/`. Neither has a webhook; workflows mint their tokens
with `actions/create-github-app-token`. The reviewer needs Issues write because gh-aw's
safe-outputs job requests it for PR comments and failure reports.

### Where the credentials live

Claude can run any code in the job it works in: `pnpm run` executes scripts it may have
edited. Code in a job can read every secret that job uses. So no job where Claude runs holds a
write credential or an app key:

- Claude's job has the read-only job token and `CLAUDE_CODE_OAUTH_TOKEN`. It commits locally,
  through the lefthook pre-commit and pre-push gates, and hands the commit on as a git bundle.
- A separate job mints the app token, checks out `main`, accepts the bundle only if it is
  exactly one commit on top of the expected starting point, and pushes it. It runs no code
  from the agent's branch.

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

The issue moves from `agent:ready` to `agent:working`. Claude Code works with the job's
read-only token: it can read the repository and issues, edit files and run `pnpm run`
scripts, but has no credential to push or open a PR. It ends with one of three outcomes:

- `pr`: the workflow runs `pnpm run verify` itself, commits as `pptx-ts-agents[bot]` with a
  Conventional Commit subject, pushes it to `agent/issue-N` and opens a PR labelled `agent:pr`. The body follows
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

## Fix rounds

`.github/workflows/agent-address-review.yml` fixes an agent PR when either of these happens:

- `pptx-ts-reviewer[bot]` submits a review requesting changes.
- The `CI` workflow fails on the PR.

It skips the PR when it is closed, lacks `agent:pr`, carries `needs-maintainer`, or has moved
past the commit the review or CI run was about; a newer push has its own review and CI run.
One round runs at a time per PR.

Claude gets the review and its inline comments, or the failed jobs' log tails, and the PR's
checkout. It fixes each blocking finding or declines it with a reason, and runs
`pnpm run verify`. The workflow then pushes one commit, `fix: address review` or
`fix: repair CI`, to the PR's branch. The squash merge keeps the PR title, so these subjects
never reach `main`. The push starts CI and a re-review.

Claude's answer to each inline comment is posted as a reply on its thread, starting with
`Fixed:` or `Declined:`, and a summary of the round goes on the PR as a comment.

The PR gets `needs-maintainer`, and auto-merge is turned off, when:

- three fix rounds have run and the reviewer or CI still fails it
- a round changes nothing, so there is nothing to re-review
- a round's change touches a protected path
- the run fails
