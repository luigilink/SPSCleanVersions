# Releasing

This document describes how a new version of **SPSCleanVersions** is released. It follows
[Keep a Changelog](https://keepachangelog.com/) and [Semantic Versioning](https://semver.org/).

## Principle

Nothing is "frozen" during day-to-day development:

- The version numbers inside `scripts/SPSCleanVersions.ps1` stay at their **last published
  value** and are **only bumped on a release branch**.
- `CHANGELOG.md` accumulates entries under a single `## [Unreleased]` heading.
- `RELEASE-NOTES.md` holds the notes for the next release under `## [Unreleased]`.

This way `main` never advertises a version that has not actually been published to the
PowerShell Gallery.

## During development

For every pull request:

1. Work on a `feature/...`, `fix/...`, `docs/...` or `chore/...` branch off `main`.
2. Add a bullet under the relevant `### Added` / `### Changed` / `### Fixed` /
   `### Removed` / `### Documentation` group of the `## [Unreleased]` section in
   `CHANGELOG.md` (and, when it is worth surfacing in the GitHub Release, in
   `RELEASE-NOTES.md`).
3. Do **not** bump the script version and do **not** add a dated changelog section.

## The three version fields

`SPSCleanVersions` ships as a single script, so its version lives in **three** places that
must always match. They are bumped **together, only when cutting a release**:

1. `.VERSION` in the `<#PSScriptInfo ... #>` block (used by `Publish-Script` for the
   PowerShell Gallery).
2. `Version:` in the comment-based help `.NOTES` block.
3. `$script:ScriptVersion` (printed at runtime and shown in the HTML report).

## Cutting a release

When `main` is in a state worth publishing, decide the new version `x.y.z`
(MAJOR.MINOR.PATCH per SemVer) and:

1. Create a release branch from `main`:

   ```bash
   git checkout main && git pull
   git checkout -b release/x.y.z
   ```

2. **Freeze the changelog.** In `CHANGELOG.md`, rename `## [Unreleased]` to
   `## [x.y.z] - YYYY-MM-DD` (today's date) and add a fresh empty `## [Unreleased]`
   above it.

3. **Fill the release notes.** In `RELEASE-NOTES.md`, rename the `## [Unreleased]`
   section to `## [x.y.z] - YYYY-MM-DD` (this becomes the GitHub Release body).

4. **Bump the three version fields** listed above to `x.y.z`.

5. Open a PR from `release/x.y.z` to `main`, get it reviewed and merged.

6. **Tag the merge commit** on `main`:

   ```bash
   git checkout main && git pull
   git tag vx.y.z
   git push origin vx.y.z
   ```

   The `Release` workflow then builds the `scripts/` ZIP, creates a GitHub Release using
   `RELEASE-NOTES.md` as the body, and publishes the script to the PowerShell Gallery
   (`Publish-Script`).

## Notes

- Keep the three version fields on the **same** value as the release tag.
- Re-pushing a tag re-runs the release workflow, refreshing the release assets and
  re-publishing to the Gallery (subject to Gallery versioning rules — a version already
  published cannot be overwritten, so bump before re-publishing).
