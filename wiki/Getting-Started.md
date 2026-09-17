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

### Local authentication (interactive) — `ClientId` required

For **local execution**, the script signs in interactively and therefore needs an Azure AD
App Registration. Register one once (public client, `http://localhost` redirect,
delegated `Sites.FullControl.All`):

```powershell
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "SPSCleanVersions" -Tenant <tenant>.onmicrosoft.com -Interactive
```

Pass the resulting **Client ID** as the `ClientId` config property. `ClientId` is
**mandatory** for local runs (the script raises a clear error if it is missing).

> **Batch runs sign in once.** For a list of many sites the script signs in interactively
> **once** (to the first site) and then reuses the delegated SharePoint token — which is
> valid tenant-wide — for every site, refreshing it silently via MSAL before it expires.
> You are prompted a single time, not once per site, so batches of hundreds/thousands of
> sites run unattended after the initial sign-in. Keep the machine awake and the sign-in
> session fresh; a tenant Conditional Access policy that forces re-authentication mid-run
> can still interrupt a very long batch.

## Installation

Install from the [PowerShell Gallery](https://www.powershellgallery.com/packages/SPSCleanVersions):

```powershell
Install-Script -Name SPSCleanVersions
```

Or [download the latest release](https://github.com/luigilink/SPSCleanVersions/releases/latest) and unzip to a directory on your machine.

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

> **Note on the site version policy modes in a runbook.** `Legacy` works with the Managed Identity (app-only). For the site version policy modes (`AutoExpiration`, `ExpireAfter`, `NoExpiration`, `InheritFromTenant`), reads (`Get-PnPSiteVersionPolicy`) and the **site default that governs new libraries** work app-only, but applying the policy to **existing** document libraries does **not** (`Set-PnPSiteVersionPolicy -ApplyToExistingDocumentLibraries` fails with *"Cannot call this API with an app-only principal"*). See the limitations table below.

### Azure Automation (app-only) limitations

When the script runs as a runbook it authenticates with the Automation Account's **Managed Identity** (app-only). Some SharePoint APIs require a **delegated** user context and therefore cannot run app-only. The script detects the runbook context and degrades gracefully (warns and skips) instead of failing hard.

| Capability | App-only (runbook) | Notes |
|---|---|---|
| Enumerate tenant sites (`Get-PnPTenantSite`, `SiteScope: All`) | ✅ Works | Confirmed with Managed Identity. |
| Read the site version policy (`Get-PnPSiteVersionPolicy`, drift detection) | ✅ Works | Reads succeed app-only. |
| `Legacy` mode — per-library major/minor limits (`Set-PnPList`) | ✅ Works | Fully supported app-only. |
| Site version policy → **site default / new libraries** (`-ApplyToNewDocumentLibraries`) | ✅ Works | *"The setting for new libraries takes effect immediately."* |
| Site version policy → **existing libraries** (`-ApplyToExistingDocumentLibraries`) | ❌ Not supported | *"Cannot call this API with an app-only principal."* In a runbook, `ApplyTo=Both` is downgraded to `New` and `ApplyTo=Existing` is skipped, with a warning. |
| Force delete old versions (`New-PnPSiteFileVersionBatchDeleteJob`, `ForceDeleteOldVersions`) | ❌ Not supported | Requires a delegated context; auto-skipped in Azure Automation. |
| HTML report / transcript | ❌ Not produced | The Automation sandbox has no persistent filesystem; per-site actions and the summary line appear in the job output instead. |
| Interactive login | ❌ Not available | Managed Identity only — there is no user context. |

> **To cover existing document libraries**, run the same mode **locally / interactively** with a **SharePoint Administrator** account, e.g. `ApplyTo=Existing` with a `ClientId` (App Registration) for interactive PnP sign-in. The runbook handles the site default and new libraries; the local delegated run handles existing libraries.

## Next Step

For JSON parameter configuration and examples, go to the [Configuration](./Configuration) page.
For usage details, go to the [Usage](./Usage) page.

## Change log

A full list of changes in each version can be found in the [change log](https://github.com/luigilink/SPSCleanVersions/blob/main/CHANGELOG.md).
