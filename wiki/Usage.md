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

### JSON schema (shared by both sources)

| Property | Type | Required | Default | Description |
|---|---|---|---|---|
| `SiteUrls` | string[] | Conditional | — | Site Collection URLs to process. Required unless `SiteScope` is `All`. |
| `KeepMajorVersions` | integer | No | `50` | Major versions to keep (maps to `-MajorVersions` in the site policy modes). |
| `KeepMinorVersions` | integer | No | `0` | Minor versions to keep (maps to `-MajorWithMinorVersions`). |
| `ClientId` | string | No | — | Azure AD App Registration / Managed Identity Client ID. |
| `ForceDeleteOldVersions` | boolean | No | `false` | Trigger a batch delete of old file versions (delegated context; skipped app-only). |
| `DryRun` | boolean | No | `false` | Simulate changes without applying them (runbook-friendly `-WhatIf`). |
| `VersionPolicyMode` | string | No | `Legacy` | `Legacy` (per-library `Set-PnPList`), or `AutoExpiration` / `ExpireAfter` / `NoExpiration` / `InheritFromTenant` (site-level `Set-PnPSiteVersionPolicy`). |
| `ExpireVersionsAfterDays` | integer | No | `0` | Expiration window for `ExpireAfter` (must be `0` or `>= 30`). |
| `ApplyTo` | string | No | `Both` | `New` / `Existing` / `Both` document libraries (site policy modes only). |
| `SiteScope` | string | No | `Selected` | `Selected` (process `SiteUrls`) or `All` (enumerate the tenant via `Get-PnPTenantSite`). |
| `TenantAdminUrl` | string | Conditional | — | SharePoint admin center URL, required when `SiteScope` is `All`. |
| `SiteFilter` | string | No | — | Server-side `-Filter` for `Get-PnPTenantSite` when `SiteScope` is `All`. |
| `EnableReport` | boolean | No | `true` | Write a local HTML report to `Results/` (local execution only). |
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

Set `"ForceDeleteOldVersions": true` to submit a batch delete job that removes old file versions exceeding the configured limits via `New-PnPSiteFileVersionBatchDeleteJob`.

> **Note:** Batch delete requires **delegated user context**. It is automatically skipped when running in Azure Automation (Managed Identity / app-only). A warning is displayed in that case.

```powershell
.\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/News"],"KeepMajorVersions":50,"ForceDeleteOldVersions":true,"DryRun":true}'
```

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

## Error Handling

Ensure the provided credentials have access to the SharePoint Sites.

## Notes

Test the script in a non-production environment before deploying it widely. The HTML report is written to `Results/` **during local execution only**; in Azure Automation the per-site actions and a run summary are printed to the job output instead.

## Support

For issues or questions, please contact the script maintainer or refer to the project documentation.
