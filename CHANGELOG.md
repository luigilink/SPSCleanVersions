# Change log for SPSCleanVersions

The format is based on and uses the types of changes according to [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [3.1.1] - 2026-07-15

### Fixed

- SPSCleanVersions.ps1
  - **Azure Automation runbook compatibility:** remove the parameter sets introduced
    with `-ConfigFile` in 3.0.0. Azure Automation rejects runbooks that use parameter
    sets (*"Parameter sets in runbooks are not supported in this release"*), which
    prevented the script — whose primary target is Azure Automation — from starting as
    a runbook. `-InputJson` and `-ConfigFile` are now plain optional parameters whose
    mutual exclusivity is validated in the body (exactly one is required)
  - **Azure Automation runbook compatibility:** explicitly `Import-Module PnP.PowerShell`.
    In a runbook the `#Requires -Modules` directive does not import the module and command
    auto-loading is unreliable in the sandbox, so `Connect-PnPOnline` was *"not recognized"*;
    the explicit import fixes this (harmless locally)

### Documentation

- Wiki (Getting Started)
  - Add a PowerShell / PnP.PowerShell version compatibility matrix. PnP.PowerShell 3.x
    requires PowerShell 7.4, so in Azure Automation use a **PowerShell 7.4 Runtime
    Environment** with PnP 3.x, or PnP.PowerShell 2.12.x on a classic 7.2 runbook. A
    mismatch fails at startup with *"requires a minimum PowerShell version of '7.4.0'"*

## [3.1.0] - 2026-07-15

### Added

- SPSCleanVersions.ps1
  - Add `VersionPolicyMode` property selecting the version policy mechanism:
    `Legacy` (default) keeps the per-library count-based `Set-PnPList` behaviour;
    `AutoExpiration`, `ExpireAfter`, `NoExpiration` and `InheritFromTenant` drive
    `Set-PnPSiteVersionPolicy` at the site level (the modern version-history model)
  - Add `ExpireVersionsAfterDays` (for `ExpireAfter`, must be >= 30; `NoExpiration`
    forces 0) with validation, addressing the customer request for version expiration
  - Add `ApplyTo` (`New` / `Existing` / `Both`, default `Both`) mapping to
    `-ApplyToNewDocumentLibraries` / `-ApplyToExistingDocumentLibraries`
  - Add `Set-SiteVersionPolicy` helper that builds the `Set-PnPSiteVersionPolicy`
    call per mode and honours `ShouldProcess` / `-WhatIf` / `DryRun`
  - Add `Test-SiteVersionPolicyDrift`: the site version policy is applied only when
    the current policy (read with `Get-PnPSiteVersionPolicy`) differs from the desired
    settings; compliant sites are skipped. Reading fails safe — an unreadable current
    policy is treated as a drift and applied
  - Only pass `MajorWithMinorVersions` when the request targets existing document
    libraries, per the `Set-PnPSiteVersionPolicy` constraint (it is rejected for a
    new-libraries-only request when auto expiration is off)
  - Add `SiteScope` (`Selected` default, or `All`), `TenantAdminUrl` and `SiteFilter`
    properties: `SiteScope: All` enumerates every site collection via
    `Get-PnPTenantSite` (OneDrive excluded, optional server-side filter) and applies
    the version policy across the tenant, still gated by the per-site drift check.
    `SiteUrls` is optional for `All` (requires `TenantAdminUrl`); `All` is not supported
    with the `Legacy` mode
  - Add a warning about the app-only (Azure Automation / Managed Identity) limitation
    of `Get-`/`Set-PnPSiteVersionPolicy`, which require a delegated site-collection-admin
    context
  - Add run logging and a self-contained HTML report: a per-site result
    (Applied / Skipped / Compliant / Failed) is collected and rendered as a
    dependency-free HTML report (summary cards + filterable table, HTML-encoded).
    Local runs write a transcript to `Logs/` and the report to `Results/`;
    Azure Automation emits the HTML into the output stream. Add `EnableReport`
    (default `true`) and `LogRetentionDays` (default `180`) properties
- SPSCleanVersions.Tests.ps1
  - Add tests for the site version policy modes, `ApplyTo` mapping, and functional
    contexts validating the `ExpireVersionsAfterDays` rules, the drift comparison
    against the real `Get-PnPSiteVersionPolicy` field shapes, and the HTML report
    generation
- Wiki Documentation
  - Document `VersionPolicyMode`, `ExpireVersionsAfterDays`, `ApplyTo`, `SiteScope`,
    the drift-based apply behaviour and the logging/report options

## [3.0.0] - 2026-07-15

### Added

- SPSCleanVersions.ps1
  - Add `-ConfigFile` parameter (separate parameter set) to load configuration
    from a local JSON file for local execution and testing. `-InputJson` remains
    the inline-string input for Azure Automation Runbooks; both sources share the
    same `ConvertFrom-Json` parsing and validation
  - Add `DefaultParameterSetName = 'InlineJson'` to `CmdletBinding` and split
    `-InputJson` / `-ConfigFile` into mutually exclusive parameter sets
  - Add missing-file and empty-file guards for `-ConfigFile`
  - Harden JSON parsing: trim the input, strip an accidental wrapping pair of
    single quotes (`'...'`) copied from a PowerShell command line, and reject
    non-object JSON (string/scalar) with a clear, actionable message instead of
    the misleading `SiteUrls is required` error (fixes #9)
- Config/SPSCleanVersions.example.json
  - Add example JSON configuration template for the `-ConfigFile` parameter
- SPSCleanVersions.Tests.ps1
  - Add tests for the two mutually exclusive parameter sets and the config-file
    loading branch (Get-Content, missing-file guard, parameter-set switch)
  - Add tests for the hardened JSON parsing (trim, single-quote stripping,
    non-object rejection) including a functional context reproducing the
    wrapped-quotes and curly-quotes cases
- Wiki Documentation
  - Document the `-ConfigFile` parameter and the file-based configuration workflow

### Changed

- .gitignore
  - Ignore real `Config/*.json` files (only `*.example.json` is tracked) and local
    `Logs/` and `Results/` run artifacts

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
