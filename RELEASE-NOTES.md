# SPSCleanVersions - Release Notes

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

A full list of changes in each version can be found in the [change log](CHANGELOG.md)
