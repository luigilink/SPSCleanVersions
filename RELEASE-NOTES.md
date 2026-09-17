# SPSCleanVersions - Release Notes

## [Unreleased]

### Changed

- SPSCleanVersions.ps1
  - **Azure Automation (app-only) — existing document libraries.** Applying a site
    version policy to *existing* document libraries is not supported app-only
    (*"Cannot call this API with an app-only principal"*). The runbook now gracefully
    drops the existing-libraries target instead of hard-failing: `ApplyTo=Both` is
    downgraded to `New`, `ApplyTo=Existing` is skipped, each with a clear warning to run
    the existing pass locally / interactively with a SharePoint Administrator.

### Fixed

- SPSCleanVersions.ps1
  - **Site version policy (`ExpireAfter` / `NoExpiration`)** — applying the policy to
    existing document libraries no longer fails when no minor-version count is configured
    (`KeepMinorVersions` absent / `0`). `MajorWithMinorVersions` is now sent (including
    `0`) whenever existing libraries are targeted, and still omitted for
    new-libraries-only requests.
    ([#33](https://github.com/luigilink/SPSCleanVersions/issues/33))

### Documentation

- Wiki
  - New **"Azure Automation (app-only) limitations"** section documenting what works and
    what requires a delegated context under a Managed Identity runbook.

A full list of changes in each version can be found in the [change log](CHANGELOG.md)
