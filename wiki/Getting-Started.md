# Getting Started

## Requirements

### PowerShell 7.2+ (Core)

Requires PowerShell 7.2 or later with PSEdition Core. [Installation guide](https://learn.microsoft.com/en-us/powershell/scripting/install/install-powershell?view=powershell-7.5).

### Module PnP.PowerShell (>= 2.12.0)

This tool relies on the PnP.PowerShell module version 2.12.0 or later. [Installation guide](https://pnp.github.io/powershell/articles/installation.html).

> **⚠️ PowerShell / PnP.PowerShell version matrix.** PnP.PowerShell **3.x requires PowerShell 7.4+**, while the 2.x line supports PowerShell 7.2. Make sure the two match, especially in **Azure Automation**:
>
> | Runbook / host runtime | Compatible PnP.PowerShell | Notes |
> |---|---|---|
> | PowerShell **7.4** (Runtime Environment) | **3.x** (e.g. 3.3.0) | Recommended. All modes work (Legacy + site version policy + tenant scope). |
> | PowerShell **7.2** (classic runbook) | **2.12.x** | Legacy mode works; the newer site version policy cmdlets may be unavailable. |
>
> A mismatch (e.g. PnP 3.x on a 7.2 runbook) fails at startup with *"The module requires a minimum PowerShell version of '7.4.0'"*. In Azure Automation, create a **PowerShell 7.4 Runtime Environment** and import PnP.PowerShell 3.x into it.

### Permissions

* **Role:** SharePoint Administrator or Global Administrator.
* **API Permissions:** `Sites.FullControl.All` (when using App Registration).

## Installation

1. [Download the latest release](https://github.com/luigilink/SPSCleanVersions/releases/latest) and unzip to a directory on your machine tool.

## Azure Automation runbook setup

> **⚠️ Breaking change vs 2.0.1.** SPSCleanVersions 3.x relies on **PnP.PowerShell 3.x**, which requires **PowerShell 7.4**. The classic PowerShell **7.2** runbook runtime (used by 2.0.1) is **not** compatible with PnP 3.x. To run 3.x in Azure Automation you must use a **PowerShell 7.4 Runtime Environment**. If you must stay on the 7.2 runtime, keep using 2.0.1 (or use PnP.PowerShell 2.12.x, Legacy mode only).

Validated setup for running the script as an Azure Automation runbook:

1. **Automation Account** with a **system-assigned Managed Identity** enabled.
2. **PowerShell 7.4 Runtime Environment**: Automation Account → *Runtime Environments* → **Create** (Language: PowerShell, Version: **7.4**) and, on the **Packages** tab, add **PnP.PowerShell 3.x**.
3. **Grant the Managed Identity SharePoint permission** (app-only). Using Microsoft Graph PowerShell as an administrator:
   ```powershell
   Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All'
   $spo  = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0ff1-ce00-000000000000'"
   $role = $spo.AppRoles | Where-Object { $_.Value -eq 'Sites.FullControl.All' -and $_.AllowedMemberTypes -contains 'Application' }
   New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId '<MI-object-id>' -BodyParameter @{
       principalId = '<MI-object-id>'; resourceId = $spo.Id; appRoleId = $role.Id
   }
   ```
   (The MI object id is on the Automation Account → *Identity* → *System assigned* page. Allow a few minutes for propagation.)
4. **Create the runbook** (PowerShell), **link it to the 7.4 Runtime Environment**, paste `scripts/SPSCleanVersions.ps1`, then **Save** and **Publish**.
5. **Test** from the Test pane with an `InputJson` value (raw JSON, no surrounding quotes), e.g.:
   ```json
   {"SiteUrls":["https://contoso.sharepoint.com/sites/Team"],"VersionPolicyMode":"Legacy","KeepMajorVersions":50,"DryRun":true}
   ```

> **Note on the site version policy modes in a runbook.** `Legacy` works with the Managed Identity (app-only). The site version policy modes (`AutoExpiration`, `ExpireAfter`, `NoExpiration`, `InheritFromTenant`) call `Get-`/`Set-PnPSiteVersionPolicy`, which require a **delegated** site-collection-admin context and may fail app-only — the script warns about this. For those modes, prefer running interactively/locally with a SharePoint Administrator account.

## Next Step

For JSON parameter configuration and examples, go to the [Configuration](./Configuration) page.
For usage details, go to the [Usage](./Usage) page.

## Change log

A full list of changes in each version can be found in the [change log](https://github.com/luigilink/SPSCleanVersions/blob/main/CHANGELOG.md).
