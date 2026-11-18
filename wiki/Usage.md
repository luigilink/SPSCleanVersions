# Usage Guide for `SPSCleanVersions.ps1`

## Overview

`SPSCleanVersions.ps1` is a PowerShell script tool to clean Version History in your SharePoint Tenant. Optimize your storage costs by managing major and minor versions across libraries and lists, apply site-level version policies (including expiration), and process selected sites or the whole tenant — only where a change is actually needed.

Install from the PowerShell Gallery:

```powershell
Install-Script -Name SPSCleanVersions
```

## Requirements

### PowerShell 7.2+ / 7.4 for PnP 3.x

Requires PowerShell 7 (Core). **PnP.PowerShell 3.x requires PowerShell 7.4** (the 2.12.x line supports 7.2). In Azure Automation, use a **PowerShell 7.4 Runtime Environment** with PnP 3.x. See the version matrix in [Getting Started](./Getting-Started).

### Module PnP.PowerShell

This tool relies on the PnP.PowerShell module. Use **3.x** for the site version policy modes (requires PS 7.4); the 2.12.x line covers the Legacy mode on PS 7.2. [Installation guide](https://pnp.github.io/powershell/articles/installation.html).

### Permissions

* **Role:** SharePoint Administrator or Global Administrator.
* **API Permissions:** `Sites.FullControl.All` (when using App Registration / Managed Identity).

## Parameters

Configuration is provided as JSON, from one of two **mutually exclusive** parameters (supply exactly one):

| Parameter | Type | Description |
|---|---|---|
| `-InputJson` | string | Inline JSON string with all configuration. Ideal for Azure Automation Runbooks (paste the raw JSON, no surrounding quotes). |
| `-ConfigFile` | string | Path to a local JSON file with the same schema. Ideal for local execution and testing. |

> Parameter sets are **not** used (they are unsupported in Azure Automation runbooks); the script validates at runtime that exactly one source is provided.

A switch selects a read-only mode: **`-ExportJobStatus`** reads the trim job state of each site and writes it to `Results/SPSCleanVersions-JobStatus-<timestamp>.csv` (local runs only; see [Example 7](#example-7-follow-the-trim-jobs-export-their-status)).

### JSON schema (shared by both sources)

| Property | Type | Required | Default | Description |
|---|---|---|---|---|
| `SiteUrls` | string[] | Conditional | — | Site Collection URLs to process. Required unless `SiteScope` is `All`. |
| `KeepMajorVersions` | integer | No | `50` | Major versions to keep (maps to `-MajorVersions` in the site policy modes). |
| `KeepMinorVersions` | integer | No | `0` | Minor versions to keep (maps to `-MajorWithMinorVersions`). |
| `ClientId` | string | No | — | Azure AD App Registration / Managed Identity Client ID. |
| `ForceDeleteOldVersions` | boolean | No | `false` | Submit a file version trim job on each site so existing versions are cleaned; the job follows `VersionPolicyMode` (age for `ExpireAfter`, automatic for `AutoExpiration`, count otherwise). Delegated context; skipped app-only. |
| `TrimJobMode` | string | No | `Auto` | Trim job submitted by `ForceDeleteOldVersions`: `Auto` (follows `VersionPolicyMode`), `Age`, `Count` or `Automatic`. One criterion per job: use one pass per criterion. |
| `TrimOnly` | boolean | No | `false` | With `ForceDeleteOldVersions`, submit **only** the trim job: the site version policy is not read or changed. Delegated context only. |
| `DryRun` | boolean | No | `false` | Simulate changes without applying them (runbook-friendly `-WhatIf`). |
| `VersionPolicyMode` | string | No | `Legacy` | `Legacy` (per-library `Set-PnPList`), or `AutoExpiration` / `ExpireAfter` / `NoExpiration` / `InheritFromTenant` (site-level `Set-PnPSiteVersionPolicy`). |
| `ExpireVersionsAfterDays` | integer | No | `0` | Expiration window for `ExpireAfter` (must be `0` or `>= 30`). |
| `ApplyTo` | string | No | `Both` | `New` / `Existing` / `Both` document libraries (site policy modes only). |
| `SiteScope` | string | No | `Selected` | `Selected` (process `SiteUrls`) or `All` (enumerate the tenant via `Get-PnPTenantSite`). |
| `TenantAdminUrl` | string | Conditional | — | SharePoint admin center URL, required when `SiteScope` is `All`. |
| `SiteFilter` | string | No | — | Server-side `-Filter` for `Get-PnPTenantSite` when `SiteScope` is `All`. |
| `EnableReport` | boolean | No | `true` | Write a local HTML report to `Results/` (plus a machine-readable `SPSCleanVersions-<timestamp>.json`) — local execution only. |
| `EnumerateLibraries` | boolean | No | `false` | Site policy modes only. Also list the in-scope document libraries as informative `InScope` rows (adds a `Get-PnPList` per site). |
| `AddSiteCollectionAdmin` | boolean | No | `false` | Delegated only. Temporarily add the operator as site collection admin per site (JIT), then revoke. Needs `TenantAdminUrl` + SharePoint Administrator role. Nothing is granted in a `DryRun` (reported as *would grant*). See [Configuration](./Configuration#jit-site-collection-admin-addsitecollectionadmin). |
| `LogRetentionDays` | integer | No | `180` | Prune `Logs/`/`Results/` files older than N days (local only; `0` disables). |

See the [Configuration](./Configuration) page for the full reference and more examples.

## Examples

### Example 1: Single Site with WhatIf (local)

```powershell
.\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/News"],"KeepMajorVersions":50}' -WhatIf
```

### Example 2: Local run from a JSON config file

```powershell
.\SPSCleanVersions.ps1 -ConfigFile '.\Config\contoso-PROD.json'
```

### Example 3: Multiple Sites with DryRun (Azure Automation Runbook)

```powershell
.\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/News","https://contoso.sharepoint.com/sites/HR"],"KeepMajorVersions":50,"DryRun":true}'
```

### Example 4: Force Deletion of Old File Version History

Set `"ForceDeleteOldVersions": true` to submit a file version trim job on each site, via `New-PnPSiteFileVersionBatchDeleteJob`, so that **existing** versions are cleaned (a version policy only governs versions created afterwards). The job matches `VersionPolicyMode`: `ExpireAfter` deletes versions older than `ExpireVersionsAfterDays`, `AutoExpiration` runs the automatic algorithm, and `Legacy` / `NoExpiration` / `InheritFromTenant` delete the versions beyond `KeepMajorVersions` / `KeepMinorVersions`. See [Configuration](./Configuration#trimming-existing-versions-forcedeleteoldversions) for details.

> **Note:** Trim jobs require **delegated user context**. They are automatically skipped when running in Azure Automation (Managed Identity / app-only). A warning is displayed in that case.

> **Warning:** versions deleted by a trim job are permanently deleted (no recycle bin). Run with `"DryRun": true` first.

```powershell
.\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/News"],"KeepMajorVersions":50,"ForceDeleteOldVersions":true,"DryRun":true}'
```

To trim by age with an expiration policy (versions older than 365 days are deleted, up to 100 major versions are kept):

```powershell
.\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/News"],"VersionPolicyMode":"ExpireAfter","ExpireVersionsAfterDays":365,"KeepMajorVersions":100,"ForceDeleteOldVersions":true,"DryRun":true}'
```

To submit **only** the trim job, without reading or changing the version policy (the policy is already in place, or this is the second pass of a trim that uses two criteria), add `TrimOnly` and choose the job with `TrimJobMode`:

```powershell
.\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/News"],"ClientId":"<app id>","ForceDeleteOldVersions":true,"TrimOnly":true,"TrimJobMode":"Age","ExpireVersionsAfterDays":365,"DryRun":true}'
```

> **One criterion per job:** an age limit and a count limit cannot be combined in one job, and a site accepts one job at a time. Run the script once per criterion (`"TrimJobMode":"Age"`, then `"TrimJobMode":"Count"`) and wait for the first jobs to finish. The age-based job never deletes versions created in the last 30 days and the script refuses a cut-off before January 1, 2023, which that job always deletes (see [Configuration](./Configuration#choosing-the-trim-job-trimjobmode)).

### Example 5: Apply a site version policy with expiration

Set `"VersionPolicyMode"` to a site-level mode to apply a policy via `Set-PnPSiteVersionPolicy`. The example below expires versions after 180 days and keeps up to 100 major versions, on new and existing libraries.

```powershell
.\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/News"],"VersionPolicyMode":"ExpireAfter","ExpireVersionsAfterDays":180,"KeepMajorVersions":100,"ApplyTo":"Both"}'
```

> **Note:** `ExpireVersionsAfterDays` must be `0` (no expiration) or `>= 30`. See [Configuration](./Configuration#version-policy-modes) for all modes.

### Example 6: Apply a policy tenant-wide (SiteScope: All)

Set `"SiteScope": "All"` with a `TenantAdminUrl` to enumerate every site collection via `Get-PnPTenantSite` and apply the policy across the tenant (only where a drift is detected). Always dry-run first.

```powershell
.\SPSCleanVersions.ps1 -InputJson '{"SiteScope":"All","TenantAdminUrl":"https://contoso-admin.sharepoint.com","VersionPolicyMode":"ExpireAfter","ExpireVersionsAfterDays":180,"KeepMajorVersions":100,"DryRun":true}'
```

> **⚠️ Warning:** `SiteScope: All` can touch thousands of sites. Run with `"DryRun": true` first and consider narrowing with `SiteFilter`. See [Tenant-wide scope](./Configuration#tenant-wide-scope-sitescope-all).

### Example 7: Follow the trim jobs (export their status)

`-ExportJobStatus` is read-only: it reads the state of the trim job of each site (`Get-PnPSiteFileVersionBatchDeleteJobStatus`) and writes it to `Results/SPSCleanVersions-JobStatus-<timestamp>.csv`, one row per site with the status, the times, the versions deleted and the storage released. The same configuration file as the run that submitted the jobs can be used. Local (delegated) runs only.

```powershell
.\SPSCleanVersions.ps1 -ConfigFile '.\Config\trim-age.json' -ExportJobStatus
```

SharePoint does not keep the state of a completed job indefinitely, so export it soon after the jobs finish. See [Configuration](./Configuration#tracking-trim-jobs-exportjobstatus) for the columns and the summary line.

## Error Handling

The script is designed to process a batch resiliently — one bad site never aborts the run:

- **Permission failures (`AccessDenied`).** For delegated (local/interactive) runs the signed-in account must be a **site collection administrator** on each target site (delegated rights = app scope ∩ user rights). A site the account cannot manage is reported with a dedicated `AccessDenied` outcome and an actionable warning, then **skipped** — the other sites keep processing. The end-of-run summary reports the access-denied count and the HTML report flags those sites. See [Site collection administrator requirement](./Configuration#site-collection-administrator-requirement-delegated-runs).
- **Throttling (HTTP 429/503).** SharePoint calls are retried automatically, honouring the server `Retry-After` hint (capped at 300s) and otherwise using exponential backoff.
- **Missing sites (`NotFound`, HTTP 404).** A site that does not exist, was deleted, or whose URL is malformed (e.g. a browser/OneDrive sharing URL) is reported with a dedicated `NotFound` outcome and **skipped**. A 404 is structural, so it **fails fast** — it is not retried — which keeps a stale list from slowing the run. `SiteUrls` are also normalized (query strings/fragments stripped, de-duplicated) before processing to avoid this whole class of failure.
- **Authentication / token errors.** These are surfaced immediately (not retried, since retrying with the same token cannot recover them) with guidance to check the `ClientId` / app registration.
- **Trim job already running (`InProgress`).** A site accepts one trim job at a time. When a previous job is still running, the new request is skipped (not retried) and the site is reported as `InProgress`: run it again later.
- **Other per-site errors** are recorded as `Failed` with the error message in the report, and the run continues.

Ensure the provided identity has access to the SharePoint sites being processed, and review the run summary line (`--- SPSCleanVersions finished: ... ---`) and the HTML report for per-site outcomes.

## Notes

Test the script in a non-production environment before deploying it widely. The HTML report is written to `Results/` **during local execution only**; in Azure Automation the per-site actions and a run summary are printed to the job output instead.

## Support

For issues or questions, please contact the script maintainer or refer to the project documentation.
