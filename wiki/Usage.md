# Usage Guide for `SPSCleanVersions.ps1`

## Overview

`SPSCleanVersions.ps1` is a PowerShell script tool to clean Version History in your SharePoint Tenant. Optimize your storage costs by managing major and minor versions across libraries and lists.

## Requirements

### PowerShell 7.x

Required for performance and class-based resource management. [Installation guide](https://learn.microsoft.com/en-us/powershell/scripting/install/install-powershell?view=powershell-7.5).

### Module PnP.PowerShell

This tool relies on the PnP.PowerShell module. [Installation guide](https://pnp.github.io/powershell/articles/installation.html).

### Permissions

* **Role:** SharePoint Administrator or Global Administrator.
* **API Permissions:** `Sites.FullControl.All` (when using App Registration).

## Examples

### Example 1: Default Usage Example

```powershell
.\SPSCleanVersions.ps1 -SiteUrls @('https://contoso.sharepoint.com/sites/News') -KeepMajorVersions 50 -WhatIf
```

### Example 2: Multiple Sites Usage Example

```powershell
.\SPSCleanVersions.ps1 -SiteUrls @('https://contoso.sharepoint.com/sites/News','https://contoso.sharepoint.com/sites/HR') -KeepMajorVersions 50 -WhatIf
```

### Example 3: Force Deletion of Old File Version History

Use the `-ForceDeleteOldVersions` switch to submit a batch delete job that removes old file versions exceeding the configured limits via `New-PnPSiteFileVersionBatchDeleteJob`.

```powershell
.\SPSCleanVersions.ps1 -SiteUrls @('https://contoso.sharepoint.com/sites/News') -KeepMajorVersions 50 -ForceDeleteOldVersions -WhatIf
```

## Error Handling

Ensure the provided credentials have access to the SharePoint Sites.

## Notes

Test the script in a non-production environment before deploying it widely.

## Support

For issues or questions, please contact the script maintainer or refer to the project documentation.
