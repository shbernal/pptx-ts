---
doc-schema-version: 1
title: "Agent pipeline"
summary: "The GitHub App identities, secrets, kill switch and labels the issue-to-merge agent workflows run on."
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
