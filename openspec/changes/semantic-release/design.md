# Design

## Context

`release.yml` runs on a `v*` tag, tests, builds the two binaries and creates the release with
generated notes. Someone has to choose and push the tag. `gjed/okf-operator` already cuts releases
with semantic-release from `main`, behind a GitHub App token; this change brings the same setup
here and keeps the existing tag workflow for the binaries. The `main` ruleset requires signed
commits and linear history, so nothing in the release flow may commit to `main`.

## Decisions

- **Same setup as okf-operator: a private `package.json` with `pnpm-lock.yaml` and
  `.releaserc.js`**, run with `pnpm exec semantic-release`. A wrapper action such as
  `cycjimmy/semantic-release-action` would hide the plugin versions from Renovate and add a third
  party that holds the app token. The Node files sit at the repository root and are ignored by
  setuptools, which only packages `src/`; Renovate's default config detects `package.json` and the
  lockfile without a change to `.renovaterc.json`.
- **Tag only.** Plugins: `commit-analyzer`, `release-notes-generator` (both with the
  `conventionalcommits` preset and okf-operator's release rules and note sections) and `github`.
  No `@semantic-release/git`, no changelog, no npm publish. `branches: ["main"]` and the default
  tag format `v${version}`, which continues from `v0.2.0`.
- **App token, not `GITHUB_TOKEN`.** A tag pushed with `GITHUB_TOKEN` does not start other
  workflows, so `release.yml` would never run. The token is minted from the `gjed-release` app in
  a `release` environment restricted to `main`; the workflow itself has `contents: read`. The
  checkout uses `persist-credentials: false` so git pushes with the app token that semantic-release
  puts in the remote URL.
- **The version comes from `setuptools-scm`.** `dynamic = ["version"]`, `setuptools-scm` in
  `build-system.requires`, default settings: it reads the `v`-prefixed tags and gives
  `0.3.0` at the tag and `0.3.1.devN+g<sha>` after it. Every checkout that installs the package
  needs the history and the tags, so `fetch-depth: 0` there. The asset-upload job installs
  nothing and keeps the shallow checkout.
- **Upload with `gh release upload --clobber`, after waiting for the release.** semantic-release
  pushes the tag first and creates the release afterwards, so the tag workflow can start before
  the release exists. The workflow runs tests and two builds before the upload, which is normally
  longer than the gap, but not guaranteed. The upload step polls `gh release view` for the tag for
  up to ten minutes and fails if it never appears. It never creates the release, so there is no
  duplicate and the notes are never touched, which `softprops/action-gh-release` would risk
  (it creates a missing release and can rewrite the body). `--clobber` lets a re-run replace
  the assets.
- **Cleanup keeps the stable filter.** "Keep assets on the 3 newest stable releases" stays a
  `jq` selection on `isPrerelease == false`; the two old `-rc` releases are not touched.
- **Actions are pinned by major version tag**, as in the rest of this repository, not by commit
  SHA as in okf-operator.
- **A new capability `release-process`** records the behaviour; no existing spec describes the
  release process.

## Risks

- The first run fails until the app is installed and the `release` environment has its variable
  and secret. This is the manual setup listed in the proposal.
- The first release after merge is a minor version (`v0.3.0`), because `feat` commits landed
  since `v0.2.0`.
- A tag pushed by hand still starts `release.yml`; if no release exists for it, the upload step
  fails after the wait. Releases are meant to be cut only by semantic-release.
