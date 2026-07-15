# SPSCleanVersions - Release Notes

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
    libraries, per the `Set-PnPSiteVersionPolicy` constraint
- Wiki Documentation
  - Document `VersionPolicyMode`, `ExpireVersionsAfterDays`, `ApplyTo` and the
    drift-based apply behaviour

A full list of changes in each version can be found in the [change log](CHANGELOG.md)
