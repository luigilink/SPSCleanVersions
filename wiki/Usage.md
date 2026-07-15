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

## Parameters

The script accepts its configuration from one of two mutually exclusive sources:

| Parameter | Type | Parameter Set | Description |
|---|---|---|---|
| `-InputJson` | string | InlineJson (default) | Inline JSON string with all configuration. Ideal for Azure Automation Runbooks. |
| `-ConfigFile` | string | ConfigFile | Path to a local JSON file with the same schema. Ideal for local execution and testing. |

### JSON schema (shared by both sources)

All configuration is expressed as JSON. Supported properties:

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

## Error Handling

Ensure the provided credentials have access to the SharePoint Sites.

## Notes

Test the script in a non-production environment before deploying it widely.

## Support

For issues or questions, please contact the script maintainer or refer to the project documentation.
