# Change log for SPSCleanVersions

The format is based on and uses the types of changes according to [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.1] - 2026-02-18

### Added

- Add Pester file: .github/workflows/pester.yml
- Add test file: tests/SPSCleanVersions.Tests.ps1

## [1.1.0] - 2026-02-16

### Added

- SPSCleanVersions.ps1
  - Add `ForceDeleteOldVersions` switch parameter to force deletion of old
    file version history using `New-PnPSiteFileVersionBatchDeleteJob` cmdlet

## [1.0.0] - 2026-02-12

### Added

- README.md
  - Add code_of_conduct.md badge
  - Add Requirement and Changelog sections
- Add CODE_OF_CONDUCT.md file
- Add Issue Templates files:
  - 1_bug_report.yml
  - 2_feature_request.yml
  - 3_documentation_request.yml
  - 4_improvement_request.yml
  - config.yml
- Add RELEASE-NOTES.md file
- Add CHANGELOG.md file
- Add CONTRIBUTING.md file
- Add release.yml file
- Add scripts folder with first version of SPSCleanVersions
- Wiki Documentation in repository - Add :
  - wiki/Configuration.md
  - wiki/Getting-Started.md
  - wiki/Home.md
  - wiki/Usage.md
  - .github/workflows/wiki.yml
