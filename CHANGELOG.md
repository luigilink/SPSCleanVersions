# Change log for SPSCleanVersions

The format is based on and uses the types of changes according to [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.0.1] - 2026-03-03

### Changed

- SPSCleanVersions.ps1
  - **[BREAKING CHANGE]:** Replace all individual parameters with a single `-InputJson`
    string parameter accepting a JSON object. This ensures full compatibility
    with Azure Automation Runbooks (no array, switch, or boolean parameter
    limitations)
  - All configuration (SiteUrls, KeepMajorVersions, KeepMinorVersions,
    ClientId, ForceDeleteOldVersions, DryRun) is parsed from JSON with
    validation and defaults
  - Replace all `Write-Host` calls with `Write-Output` to avoid unhandled
    exceptions in Azure Automation (no interactive host)
  - Replace `$env:AUTOMATION_ASSET_NAME` check with new
    `Test-IsAzureAutomation` helper function that detects multiple environment
    signals (`AZUREPS_HOST_ENVIRONMENT`, `AUTOMATION_ASSET_SANDBOX_ID`,
    `IDENTITY_ENDPOINT`, etc.)
  - Use `[string]::IsNullOrEmpty()` for `ClientId` check instead of
    `$null -ne`
  - Disable PnP PowerShell update check via
    `$env:PNPPOWERSHELL_UPDATECHECK = "false"`
  - Skip `New-PnPSiteFileVersionBatchDeleteJob` when running in Azure
    Automation (app-only / Managed Identity context) since the API requires
    delegated user context; uses `Test-IsAzureAutomation` for detection

### Added

- SPSCleanVersions.ps1
  - Add `-InputJson` parameter with `ValidateNotNullOrEmpty` attribute
  - Add JSON parsing with `ConvertFrom-Json` and structured validation
    (required `SiteUrls`, optional properties with defaults)
  - Add `DryRun` JSON property to simulate changes without applying them
    (replacement for `-WhatIf` which is not supported in Azure Automation
    Runbooks)
  - Add `#Requires -Version 7.2` statement
  - Add `#Requires -PSEdition Core` statement
  - Add `#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '2.12.0' }` statement
- SPSCleanVersions.Tests.ps1
  - Add tests for `InputJson` parameter (mandatory, typed, validated)
  - Add tests for JSON parsing and validation (ConvertFrom-Json, required
    SiteUrls, default values, error on invalid JSON)
  - Add tests for multi-site `foreach` loop over `$SiteUrls`
  - Add tests for `#Requires` statements (PS 7.2, Core, PnP.PowerShell)
  - Add tests for `Test-IsAzureAutomation` function and Azure environment
    signal detection
  - Add test for PnP PowerShell update check disable
  - Add tests for app-only batch delete guard in Azure Automation context
- Wiki Documentation
  - Rewrite wiki/Configuration.md with JSON schema, properties table, and
    6 detailed examples (minimal, multi-site, dry run, force delete, full
    config, Azure Automation Runbook)
  - Update wiki/Home.md with new features and Configuration link
  - Update wiki/Usage.md with `-InputJson` parameter table and examples
  - Update wiki/Getting-Started.md with PS 7.2+ Core and PnP 2.12.0 requirements
- README.md
  - Update Key Features, Requirements, and documentation links to match
    current `-InputJson` parameter design

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
