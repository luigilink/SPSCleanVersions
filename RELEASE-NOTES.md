# SPSCleanVersions - Release Notes

## [3.5.1] - 2026-11-18

### Added

- **Example configuration files, one per scenario** in `Config/`: `legacy-count`, `auto-expiration`,
  `expire-after`, `tenant-wide`, `jit-admin`, `trim-count` and `azure-automation`, next to the existing
  `trim-only` example and the full template. Each one holds placeholder values only and `"DryRun": true`.
  The new *Example files* section of the wiki (Configuration) lists them and says when to use each one, and
  the `-ConfigFile` help links to it: the release archive and the PowerShell Gallery package contain the
  script only.

### Fixed

- SPSCleanVersions.ps1
  - **`DryRun` no longer grants or revokes site collection admin.** With `AddSiteCollectionAdmin`, a DryRun
    used to add the operator as site collection administrator on every site it did not administer and remove
    it afterwards, because the cmdlets that do it do not honour `-WhatIf`: the administrators of each site
    were changed twice during a simulation. Now nothing is granted, journalled or revoked, and such a site is
    reported as *would grant* (`JIT  N would grant` in the summary).
  - **Behaviour change in a DryRun with `AddSiteCollectionAdmin`.** A site cannot be read without the grant,
    so `TrimOnly` still reports `WouldSubmit` for it (the job needs no site access) and any other mode
    reports it as `Skipped` (not evaluated). The sites whose grant would fail (deleted or locked sites) are
    no longer surfaced by a DryRun: `-ExportJobStatus` still does, because it is read-only, ignores `DryRun`
    and still grants and revokes. `-CleanupAdminsOnly` also ignores `DryRun`.

A full list of changes in each version can be found in the [change log](CHANGELOG.md)
