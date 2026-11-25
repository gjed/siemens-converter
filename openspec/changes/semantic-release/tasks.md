# Tasks

## 1. Version from the tag

- [ ] 1.1 Make the version dynamic with `setuptools-scm` in `pyproject.toml`
- [ ] 1.2 Give the checkouts that install the package `fetch-depth: 0` in `ci.yml` and `release.yml`

## 2. semantic-release

- [ ] 2.1 Add `package.json`, `pnpm-lock.yaml` and `.releaserc.js`, and ignore `node_modules/`
- [ ] 2.2 Add `.github/workflows/semantic-release.yml`: `release` environment, app token, no overlapping runs

## 3. Tag workflow

- [ ] 3.1 Attach the binaries and `static_data_template.xlsx` to the existing release, waiting for it
- [ ] 3.2 Remove the pre-release detection and the pre-release branch of the asset cleanup

## 4. Docs

- [ ] 4.1 Describe the release process in `openspec/config.yaml` and `docs/development.md`

## 5. Verification

- [ ] 5.1 Run the tests, build the package and check the version it reports
- [ ] 5.2 Dry-run semantic-release and lint the workflows
