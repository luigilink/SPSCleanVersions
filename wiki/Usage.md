# Usage Guide for `SPSCleanVersions.ps1`

## Overview

`SPSCleanVersions.ps1` is a PowerShell script tool to clean Version History in your SharePoint Tenant. Optimize your storage costs by managing major and minor versions across libraries and lists.

## Requirements

### PowerShell 7.2+ (Core)

Requires PowerShell 7.2 or later with PSEdition Core. [Installation guide](https://learn.microsoft.com/en-us/powershell/scripting/install/install-powershell?view=powershell-7.5).

### Module PnP.PowerShell (>= 2.12.0)

This tool relies on the PnP.PowerShell module version 2.12.0 or later. [Installation guide](https://pnp.github.io/powershell/articles/installation.html).

### Permissions

* **Role:** SharePoint Administrator or Global Administrator.
* **API Permissions:** `Sites.FullControl.All` (when using App Registration).

## Parameter: `-InputJson`

All configuration is passed as a single JSON string. Supported properties:

| Property | Type | Required | Default | Description |
|---|---|---|---|---|
| `SiteUrls` | string[] | **Yes** | — | Site Collection URLs to process |
| `KeepMajorVersions` | integer | No | `50` | Number of major versions to keep |
| `KeepMinorVersions` | integer | No | `0` | Number of minor versions to keep |
| `ClientId` | string | No | — | Azure AD App Registration Client ID |
| `ForceDeleteOldVersions` | boolean | No | `false` | Trigger batch delete of old file versions |
| `DryRun` | boolean | No | `false` | Simulate changes without applying them |

## Examples

### Example 1: Single Site with WhatIf (local)

```powershell
.\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/News"],"KeepMajorVersions":50}' -WhatIf
```

### Example 2: Multiple Sites with DryRun (Azure Automation Runbook)

```powershell
.\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/News","https://contoso.sharepoint.com/sites/HR"],"KeepMajorVersions":50,"DryRun":true}'
```

### Example 3: Force Deletion of Old File Version History

Set `"ForceDeleteOldVersions": true` to submit a batch delete job that removes old file versions exceeding the configured limits via `New-PnPSiteFileVersionBatchDeleteJob`.

> **Note:** Batch delete requires **delegated user context**. It is automatically skipped when running in Azure Automation (Managed Identity / app-only). A warning is displayed in that case.

```powershell
.\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/News"],"KeepMajorVersions":50,"ForceDeleteOldVersions":true,"DryRun":true}'
```

## Error Handling

Ensure the provided credentials have access to the SharePoint Sites.

## Notes

Test the script in a non-production environment before deploying it widely.

## Support

For issues or questions, please contact the script maintainer or refer to the project documentation.
