# SPSCleanVersions - Release Notes

## [3.1.3] - 2026-07-15

### Fixed

- SPSCleanVersions.ps1
  - **SiteScope: All** — `Get-TenantSiteUrls` no longer pollutes the enumerated site
    list. Its informational messages used `Write-Output`, which PowerShell captured into
    the returned value, so the status strings were processed as bogus site URLs and
    failed with *"Invalid URI"*. Status now uses `Write-Verbose` and the function returns
    only the URL array; the caller logs the discovered/processed count
  - **DryRun** — the site version policy path recorded `Applied` even though
    `ShouldProcess`/`-WhatIf` prevented the change. It now records a **WouldApply**
    outcome and logs *"would apply"*; the report card and run summary reflect the
    simulation instead of claiming sites were changed

### Changed

- Wiki Documentation
  - Soften the app-only note: live testing confirmed `Get-PnPTenantSite` and
    `Get-PnPSiteVersionPolicy` (reads) work with a Managed Identity; writes to existing
    document libraries may still require a delegated context

A full list of changes in each version can be found in the [change log](CHANGELOG.md)
