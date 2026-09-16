# SPSCleanVersions - Release Notes

## [3.1.5] - 2026-09-16

### Fixed

- SPSCleanVersions.ps1
  - **Site version policy (`ExpireAfter` / `NoExpiration`)** — applying the policy to
    existing document libraries no longer fails when no minor-version count is configured
    (`KeepMinorVersions` absent / `0`). `MajorWithMinorVersions` is now sent (including
    `0`) whenever existing libraries are targeted, and still omitted for
    new-libraries-only requests. Fixes *"You must specify ExpireVersionsAfterDays,
    MajorVersions and MajorWithMinorVersions when EnableAutoExpirationVersionTrim is false
    for document libraries that including existing ones."*
    ([#33](https://github.com/luigilink/SPSCleanVersions/issues/33))

A full list of changes in each version can be found in the [change log](CHANGELOG.md)
