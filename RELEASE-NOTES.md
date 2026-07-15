# SPSCleanVersions - Release Notes

## [3.1.4] - 2026-07-15

### Added

- CI/CD
  - Publish the script to the **PowerShell Gallery** on tagged releases
    (`Publish-Script` step in `release.yml`). Install with
    `Install-Script -Name SPSCleanVersions`. Requires the `PSGALLERYKEY`
    repository secret
- Documentation
  - Add PowerShell Gallery install instructions (Getting Started, README) and a
    PSGallery version badge

A full list of changes in each version can be found in the [change log](CHANGELOG.md)
