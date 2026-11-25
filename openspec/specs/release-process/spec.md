# release-process Specification

## Purpose

Define how a release is versioned, tagged and published: the version is derived from the Conventional Commits that reach `main`, the tag starts the build of the binaries that are attached to the release, and the package reports the version of its tag.

## Requirements

### Requirement: Releases are cut from main

The system SHALL compute the next version from the Conventional Commits pushed to `main` since the last `v*` tag, push the tag `v<version>` and create the GitHub release with its notes, without any manual step. Releases SHALL come from `main` only; there SHALL be no pre-release channel.

#### Scenario: A feature is merged

- **WHEN** commits containing a `feat` and no breaking change reach `main`
- **THEN** a minor version is tagged and released, continuing from the last tag

#### Scenario: A fix is merged

- **WHEN** the commits that reach `main` contain a `fix`, `perf` or `revert` and no `feat` or breaking change
- **THEN** a patch version is tagged and released

#### Scenario: A breaking change is merged

- **WHEN** a commit that reaches `main` is marked as a breaking change
- **THEN** a major version is tagged and released

#### Scenario: Nothing releasable is merged

- **WHEN** the commits that reach `main` are all of another type, such as `docs`, `chore`, `ci`, `test`, `refactor`, `style` or `build`
- **THEN** no tag and no release is created

#### Scenario: Releasing never writes to main

- **WHEN** a release is cut
- **THEN** no commit is pushed to `main`: the version is not committed and no changelog file is written

#### Scenario: Releases do not overlap

- **WHEN** two pushes to `main` follow each other closely
- **THEN** their releases run one after the other, never at the same time

### Requirement: The release triggers the build

The tag SHALL be pushed with credentials that start the tag-triggered release workflow, which builds the Windows executable and the Linux binary and attaches them, with `static_data_template.xlsx`, to the release created for the tag.

#### Scenario: Binaries are attached to the release

- **WHEN** the tag `v<version>` is pushed and the build and tests pass
- **THEN** the release `v<version>` has the assets `siemens-converter-windows-v<version>.exe`, `siemens-converter-linux-v<version>` and `static_data_template.xlsx`, and its notes are unchanged

#### Scenario: The release is published with its binaries

- **WHEN** a release is created for a tag
- **THEN** it stays a draft, and the latest release unchanged, until its binaries and `static_data_template.xlsx` are attached, and only then becomes the latest release

#### Scenario: The build fails

- **WHEN** the tests or a build of the tag workflow fail
- **THEN** the release stays a draft and the previous release stays the latest, until the failed workflow run is re-run and passes

#### Scenario: The build starts before the release exists

- **WHEN** the tag workflow reaches the upload before the release for the tag has been created
- **THEN** it waits for the release instead of creating a second release or failing at once

#### Scenario: Old releases lose their assets

- **WHEN** a release has been published
- **THEN** only the 3 newest stable releases keep their assets

### Requirement: The package version comes from the tag

The version of the built and installed package SHALL be derived from the git tags, not from a value committed in `pyproject.toml`.

#### Scenario: Build at a tag

- **WHEN** the package is built from a checkout of the tag `v<version>`
- **THEN** the package version is `<version>`

#### Scenario: Build after a tag

- **WHEN** the package is built from a commit after the last tag
- **THEN** the package version is a development version of the next patch version, carrying the number of commits since the tag and the commit hash
