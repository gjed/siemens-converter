# Proposal

## Why

Releases are cut by hand: someone picks the next version and pushes a `v*` tag. Nobody remembers
to, so 34 commits, including the production fixes for
[#22](https://github.com/gjed/siemens-converter/issues/22),
[#23](https://github.com/gjed/siemens-converter/issues/23) and
[#25](https://github.com/gjed/siemens-converter/issues/25), sit on `main` since `v0.2.0` while
users keep the over-distributed costs. The commits already follow Conventional Commits, so the
version can be derived instead of chosen. The package version in `pyproject.toml` is still `0.1.0`
because nobody bumps it
([issue #35](https://github.com/gjed/siemens-converter/issues/35)).

## What Changes

- Every push to `main` computes the next version from the Conventional Commits since the last tag,
  pushes the `v*` tag and creates the GitHub release with its notes, using semantic-release. It
  commits nothing to `main`: no version bump, no changelog file.
- Breaking changes release a major version, `feat` a minor one, `fix`, `perf` and `revert` a patch.
  Every other type releases nothing.
- The tag is pushed with a token of the `gjed-release` GitHub App, so the existing tag-triggered
  workflow runs and builds the binaries.
- The package version comes from the tag (`setuptools-scm`) instead of `pyproject.toml`.
- The tag-triggered release workflow no longer creates the release or its notes. It attaches the
  two binaries and `static_data_template.xlsx` to the release semantic-release created, waiting for
  that release if the tag workflow starts first.
- No pre-releases: the pre-release detection and the pre-release branch of the asset cleanup go.
  Only the 3 newest stable releases keep their assets.
- The Git Workflow section of `openspec/config.yaml` and `docs/development.md` describe how a
  release is cut now.

## Capabilities

### New Capabilities

- `release-process`: how a release is versioned, tagged and published, and what the built package
  reports as its version.

### Modified Capabilities

None. `static-data-input` already requires `static_data_template.xlsx` to be attached to every
release; that does not change.

## Impact

- `.github/workflows/semantic-release.yml` (new), `.releaserc.js`, `package.json`,
  `pnpm-lock.yaml` (new): the release automation. The Node files are tooling only; they are not
  part of the Python package.
- `.github/workflows/release.yml`: release job attaches assets instead of creating the release;
  pre-release handling removed; `fetch-depth: 0` where the package is installed.
- `.github/workflows/ci.yml`: `fetch-depth: 0`.
- `pyproject.toml`: dynamic version from `setuptools-scm`.
- `.gitignore`: `node_modules/`.
- `openspec/config.yaml`, `docs/development.md`: release process.
- Manual setup after merge, without which the first run fails: the `gjed-release` app installed on
  the repository, and a `release` environment restricted to `main` with the
  `RELEASE_APP_CLIENT_ID` variable and the `RELEASE_APP_PRIVATE_KEY` secret.
- No change to the binaries, the PyInstaller options or the names of the release assets.
