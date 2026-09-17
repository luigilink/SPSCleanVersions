# SPSCleanVersions - Release Notes

## [Unreleased]

### Added

- SPSCleanVersions.ps1
  - **Multi-threading (local only)** via a new `Threads` config property (default `1`).
    `Threads > 1` splits the site list across that many child `pwsh` processes sharing the
    single interactive sign-in through a secured, auto-refreshed token file; workers' results
    are merged into one consolidated HTML report. Azure Automation stays sequential.
  - **Throttling-aware retry** (`Invoke-RetryCommand` + `Get-RetryAfterDelay` +
    `Test-IsAuthError`) around the SharePoint calls: HTTP 429/503 responses honour the
    server `Retry-After` hint (capped at 300s), otherwise exponential backoff. Adapted (MIT)
    from the SPO Storage Assessment toolkit.

### Changed

- SPSCleanVersions.ps1
  - **Local execution — single sign-in with delegated token reuse.** Instead of
    `Connect-PnPOnline -Interactive` per site (which re-prompts on every site and breaks
    large batches), the script now signs in interactively **once** and reuses the
    delegated, tenant-wide SharePoint token for all sites, refreshing it silently via MSAL
    before expiry. Falls back to per-site interactive if the single sign-in fails.
    `ClientId` is now required for local execution.
    ([#37](https://github.com/luigilink/SPSCleanVersions/issues/37))
  - **Azure Automation (app-only) — existing document libraries.** Applying a site
    version policy to *existing* document libraries is not supported app-only
    (*"Cannot call this API with an app-only principal"*). The runbook now gracefully
    drops the existing-libraries target instead of hard-failing: `ApplyTo=Both` is
    downgraded to `New`, `ApplyTo=Existing` is skipped, each with a clear warning to run
    the existing pass locally / interactively with a SharePoint Administrator.

### Fixed

- SPSCleanVersions.ps1
  - **Local DryRun now writes the HTML report and transcript.** `DryRun` sets
    `$WhatIfPreference` globally, which previously also suppressed the tool's own local
    artifact writes (folder creation, transcript, HTML report, retention pruning), so a
    local DryRun produced no report. These local operations now use `-WhatIf:$false`, while
    SharePoint changes stay simulated. No effect in Azure Automation.
    ([#39](https://github.com/luigilink/SPSCleanVersions/issues/39))
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
