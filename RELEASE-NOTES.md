# SPSCleanVersions - Release Notes

## [3.2.0] - 2026-09-17

### Changed

- SPSCleanVersions.ps1
  - **Single interactive prompt on every platform** — sign in once, then reuse the
    single sign-in's **SharePoint-audience** delegated token (`Get-PnPAccessToken
    -ResourceTypeName SharePoint`) for every site via `Connect-PnPOnline -AccessToken`
    (was one prompt per site on macOS). The default token is Microsoft Graph, which CSOM
    SharePoint cmdlets reject, so the SharePoint resource is requested explicitly.
  - **Retry fails fast on auth/authorization errors** instead of 5x exponential backoff on
    structural 401s.

### Added

- SPSCleanVersions.ps1
  - **Actionable "access denied" detection** — a site the signed-in account cannot manage
    (not a **site collection administrator**) is recorded as a distinct `AccessDenied`
    outcome with a clear message, the run continues, and an end-of-run advisory plus a
    dedicated report count/badge are shown. (A future opt-in will add the admin
    automatically.)

- SPSCleanVersions.ps1
  - **Full reporting**: the HTML report gains **Library / Major / Minor / ExpireAfterDays**
    columns; Legacy mode reports one row per document library with its real outcome; site
    version policy modes can list in-scope libraries via an optional `EnumerateLibraries`
    flag (informative `InScope` rows); and a machine-readable **`SPSCleanVersions-<timestamp>.json`**
    is written next to the HTML report.
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
  - New **"Site collection administrator requirement (delegated runs)"** section: delegated
    rights are the intersection of the app scope **and** the signed-in user's rights on each
    site, so the account must be a site collection administrator on every target site (a
    tenant SharePoint Administrator role is not sufficient by itself). Documents the
    `AccessDenied` outcome and the fix (`Set-PnPTenantSite -Owners`), plus the run resilience
    (single sign-in, throttling retry, fail-fast) and the enriched report.
- Config
  - `Config/SPSCleanVersions.example.json` expanded into a full template covering the site
    version policy modes and all supported properties.

A full list of changes in each version can be found in the [change log](CHANGELOG.md)
