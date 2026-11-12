---
name: release-publish
description: Use to cut and publish a new pptx-ts version (any "do a release", "minor/major/patch release", "publish vX.Y.Z", "ship a release" request in this repo). The sequence is a release branch, the CHANGELOG entries, the comparison refresh, `pnpm version` (which checks, gates, bumps and dates the changelog), a version bump pull request, and the tag push on the merged commit that makes CI publish both npm names and create the GitHub Release.
metadata:
  internal: true
---

# Releasing pptx-ts

`main` takes changes only through pull requests, so the version bump lands as one, and the tag
goes on the commit it merged as. Pushing a `vX.Y.Z` tag is the release.
`.github/workflows/publish.yml` runs the full CI gate, publishes `pptx-ts` and its alias `@shbernal/ts-pptx`, then creates the GitHub Release from the
CHANGELOG. `docs/contributing/releasing.md` explains why each step is shaped the way it is.

## Choosing the version

The user's ask wins: "minor release" means minor, even for a lone fix. Otherwise read
`git log --oneline vLAST..HEAD`: a breaking API change is major, any `feat:` is minor, and anything
else is patch.

## Steps

1. **Cut a release branch from `origin/main`.** The preflight refuses any other starting point.

   ```bash
   git fetch origin && git switch -c release/vX.Y.Z origin/main
   ```

2. **Write the CHANGELOG entries.** Every commit since the last tag needs an entry under
   `[Unreleased]` in `CHANGELOG.md`, in the dense prose style of the existing ones. Add the missing
   ones. Leave the heading alone, because `pnpm version` dates it. Stage the file.

3. **Refresh the comparison.** Leave the machine idle while `comparison:measure` runs, because it
   times both libraries. Never pass `--allow-unavailable`. If a fetch fails, re-run later.

   ```bash
   pnpm run comparison:measure && pnpm run comparison:render && pnpm run comparison:diff
   ```

   If `comparison:diff` fails, stop and show the user what it printed. Continue only if they say so.
   Then stage the snapshot and the rendered pages:

   ```bash
   git add scripts/comparison/snapshot.json docs/comparison.md docs/comparison-method.md docs/comparison-syntax.md README.md
   ```

4. **Bump and commit.** No tag yet: the squash merge makes a new commit, and the tag goes on that.

   ```bash
   pnpm version <major|minor|patch> --no-git-tag-version --no-git-checks
   git add package.json && git commit -m 'chore(release): vX.Y.Z'
   ```

   `preversion` runs `release:preflight`, then `check:static` and `verify`, and stops with nothing
   changed if any of them fails. Fix what it names and run the command again. If it gets past the
   preflight and then fails, run `git status` before retrying. The `version` script has already
   staged the constant and the dated changelog; `package.json` is the one file left to add.

5. **Open the version bump pull request and merge it.** Use the PR template, title
   `chore(release): vX.Y.Z`. Merge with a squash once CI is green. Nothing publishes yet, apart
   from the usual `next` prerelease of the merged commit.

   ```bash
   git push -u origin release/vX.Y.Z
   gh pr create --title 'chore(release): vX.Y.Z' --base main
   gh pr checks --watch && gh pr merge --squash --delete-branch
   ```

6. **Tag the merged commit and push the tag.** This publishes. Check that `origin/main` is the
   release commit first: its subject is `chore(release): vX.Y.Z (#N)` and its `package.json` says
   `X.Y.Z`. If something else merged on top, tag the release commit by its hash instead.

   ```bash
   git fetch origin main && git log -1 --format=%s origin/main && git show origin/main:package.json | grep '"version"'
   git tag -a vX.Y.Z -m vX.Y.Z origin/main && git push origin vX.Y.Z
   ```

7. **Watch the run.**

   ```bash
   gh run watch --exit-status "$(gh run list --workflow=publish.yml --event push --limit 1 --json databaseId -q '.[0].databaseId')"
   ```

   The release is done when "Publish to npm", "Publish scoped alias" and "GitHub Release" have
   all succeeded. A publish step marked as skipped means that name already had the version.

8. **Optional, for the maintainer only: comment on the issues the release closes.** Agents skip
   this step. The new CHANGELOG section cites the issues:

   ```bash
   gh issue comment <N> --repo shbernal/pptx-ts --body "Released in X.Y.Z."
   ```

   After a major release, also open an issue on `pptx-html` asking for a matching release.

## If the run fails

- **A CI gate failed:** fix it on `main` through a pull request, then release the next patch.
  Don't move a tag that has already been pushed.
- **The alias publish or a transient step failed:** re-dispatch on the same tag. The steps that
  already succeeded are skipped:
  `gh workflow run publish.yml --repo shbernal/pptx-ts --ref vX.Y.Z`.
- **Only the GitHub Release job failed:** `gh run rerun <run-id> --failed`.
