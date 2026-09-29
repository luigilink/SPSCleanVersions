# Release Process

This page documents how to ship a new version of SPSCleanVersions. The tool is a **single
script** (`scripts/SPSCleanVersions.ps1`) published to the PowerShell Gallery, so a release is
driven by two things: the **three version fields** inside that script (which must always match)
and a **`v*` git tag** that triggers the GitHub release workflow.

## Nothing is frozen during development

Day to day, `main` never advertises a version that has not actually been published:

- The three version fields in `scripts/SPSCleanVersions.ps1` stay at their **last published
  value** and are bumped **only on a release branch**.
- `CHANGELOG.md` accumulates entries under a single `## [Unreleased]` heading.
- `RELEASE-NOTES.md` holds the notes for the **next release only**, under `## [Unreleased]`.
  The whole file is used **verbatim** as the GitHub Release body, so it must never accumulate
  past releases — their history lives in `CHANGELOG.md`.

For every pull request, add a bullet under the relevant `### Added` / `### Changed` / `### Fixed`
/ `### Removed` / `### Documentation` group of `## [Unreleased]` in `CHANGELOG.md` (and, when it
is worth surfacing in the GitHub Release, in `RELEASE-NOTES.md`). Do **not** bump the version and
do **not** add a dated changelog section during development.

## Versioning policy

SPSCleanVersions follows [Semantic Versioning 2.0](https://semver.org/spec/v2.0.0.html).

| Bump | When |
|---|---|
| MAJOR (X.0.0) | Breaking change in the config JSON format, the parameters, the report/output contract, or the package layout. |
| MINOR (X.Y.0) | New backward-compatible feature (new mode, new optional setting, new parameter). |
| PATCH (X.Y.Z) | Bug fix or documentation-only change. |

## The three version fields

Because SPSCleanVersions ships as a single script, its version lives in **three** places that
must always match. They are bumped **together, only when cutting a release**:

1. `.VERSION` in the `<#PSScriptInfo ... #>` block — used by `Publish-Script` for the PowerShell
   Gallery.
2. `Version:` in the comment-based help `.NOTES` block.
3. `$script:ScriptVersion` — printed at runtime and shown in the HTML report.

## Release checklist

### 1. Create a release branch

```bash
git checkout main && git pull
git checkout -b release/x.y.z
```

### 2. Bump the three version fields

Set the `.VERSION`, `.NOTES` `Version:` and `$script:ScriptVersion` values in
`scripts/SPSCleanVersions.ps1` to `x.y.z` (all three identical).

### 3. Freeze `CHANGELOG.md`

Rename `## [Unreleased]` to `## [x.y.z] - YYYY-MM-DD` (today's date) and add a fresh empty
`## [Unreleased]` above it.

### 4. Replace `RELEASE-NOTES.md`

`RELEASE-NOTES.md` is used **verbatim** as the body of the GitHub Release. It must contain
**only** the section of the version being released:

- Rename the `## [Unreleased]` section to `## [x.y.z] - YYYY-MM-DD`.
- **Delete the previous release's section** so the file contains only the new version — no
  stacked history (that lives in `CHANGELOG.md`).

### 5. Validate locally

```powershell
$config = New-PesterConfiguration
$config.Run.Path = './tests'
Invoke-Pester -Configuration $config
```

All tests must pass. The `pester.yml` workflow re-runs them on the pull request.

### 6. Commit, open a PR, review and merge

```bash
git add -A
git commit -m "Release x.y.z"
git push -u origin release/x.y.z
```

Open a Pull Request from `release/x.y.z` to `main`, get it reviewed, and merge.

### 7. Tag the merge commit on `main`

```bash
git checkout main && git pull
git tag vx.y.z
git push origin vx.y.z
```

The `.github/workflows/release.yml` workflow then runs automatically. It:

1. Packages the **contents** of `scripts/` into `SPSCleanVersions-vx.y.z.zip` (the archive
   extracts straight to `SPSCleanVersions.ps1` at its root, with no `scripts/` wrapper).
2. Publishes a GitHub Release using `RELEASE-NOTES.md` as the body, attaching the ZIP and
   `LICENSE`.
3. Publishes the script to the PowerShell Gallery with `Publish-Script`.

### 8. Verify

- **Releases**: <https://github.com/luigilink/SPSCleanVersions/releases> — the new release is
  listed, its body shows **only** the released version, and the flat ZIP is attached.
- **Actions**: <https://github.com/luigilink/SPSCleanVersions/actions> — `release.yml` and
  `pester.yml` ran green.
- **PowerShell Gallery**: <https://www.powershellgallery.com/packages/SPSCleanVersions> — the
  new version is live (`Find-Script SPSCleanVersions`).
- **Wiki**: `wiki.yml` syncs any `wiki/` changes pushed to `main`.

## Undoing a release

If you tagged too early:

```bash
git tag -d vx.y.z
git push origin --delete vx.y.z
```

Then delete the auto-created GitHub Release, fix what needs fixing, and re-tag from the new HEAD.

> ⚠️ **The PowerShell Gallery cannot overwrite a published version.** Once `vx.y.z` is live on
> the Gallery, that version number is burned — ship a `vx.y.(z+1)` patch instead of trying to
> republish it. For the same reason, don't move a `v*` tag that has already triggered a
> successful publish.

## See also

- [Keep a Changelog](https://keepachangelog.com/en/1.0.0/)
- [Semantic Versioning 2.0](https://semver.org/spec/v2.0.0.html)
- [Configuration reference](Configuration)
- [Usage](Usage)
