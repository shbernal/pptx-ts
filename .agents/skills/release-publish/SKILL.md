---
name: release-publish
description: Use to cut and publish a new pptx-ts version (any "do a release", "minor/major/patch release", "publish vX.Y.Z", "ship a release" request in this repo). The sequence is the comparison refresh, the CHANGELOG entries, `pnpm version` (which checks, gates, bumps, dates the changelog and tags), and the tag push that makes CI publish both npm names and create the GitHub Release.
metadata:
  internal: true
---

# Releasing pptx-ts

Pushing a `vX.Y.Z` tag is the release. `.github/workflows/publish.yml` runs the full CI gate,
publishes `pptx-ts` and its alias `@shbernal/ts-pptx`, then creates the GitHub Release from the
CHANGELOG. `docs/contributing/releasing.md` explains why each step is shaped the way it is.

## Choosing the version

The user's ask wins: "minor release" means minor, even for a lone fix. Otherwise read
`git log --oneline vLAST..HEAD`: a breaking API change is major, any `feat:` is minor, and anything
else is patch.

## Steps

1. **Write the CHANGELOG entries.** Every commit since the last tag needs an entry under
   `[Unreleased]` in `CHANGELOG.md`, in the dense prose style of the existing ones. Add the missing
   ones. Leave the heading alone, because `pnpm version` dates it. Stage the file.

2. **Refresh the comparison.** Leave the machine idle while `comparison:measure` runs, because it
   times both libraries. Never pass `--allow-unavailable`. If a fetch fails, re-run later.

   ```bash
   pnpm run comparison:measure && pnpm run comparison:render && pnpm run comparison:diff
   ```

   If `comparison:diff` fails, stop and show the user what it printed. Continue only if they say so.
   Then stage the snapshot and the rendered pages:

   ```bash
   git add scripts/comparison/snapshot.json docs/comparison.md docs/comparison-method.md docs/comparison-syntax.md README.md
   ```

3. **Bump, commit and tag.**

   ```bash
   pnpm version <major|minor|patch> --message 'chore(release): v%s' --no-git-checks
   ```

   `preversion` runs `release:preflight`, then `check:static` and `verify`, and stops with nothing
   changed if any of them fails. Fix what it names and run the command again. If it gets past the
   preflight and then fails, run `git status` before retrying.

4. **Push.** This publishes:

   ```bash
   git push --atomic origin main vX.Y.Z
   ```

5. **Watch the run.**

   ```bash
   gh run watch --exit-status "$(gh run list --workflow=publish.yml --event push --limit 1 --json databaseId -q '.[0].databaseId')"
   ```

   The release is done when "Publish to npm", "Publish scoped alias" and "GitHub Release" have
   all succeeded. A publish step marked as skipped means that name already had the version.

6. **Comment on every issue the release closes.** The new CHANGELOG section cites them:

   ```bash
   gh issue comment <N> --repo shbernal/pptx-ts --body "Released in X.Y.Z."
   ```

   After a major release, also open an issue on `pptx-html` asking for a matching release.

## If the run fails

- **A CI gate failed:** fix it on `main`, then release the next patch. Don't move a tag that has
  already been pushed.
- **The alias publish or a transient step failed:** re-dispatch on the same tag. The steps that
  already succeeded are skipped:
  `gh workflow run publish.yml --repo shbernal/pptx-ts --ref vX.Y.Z`.
- **Only the GitHub Release job failed:** `gh run rerun <run-id> --failed`.
