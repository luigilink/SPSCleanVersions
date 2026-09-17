# Configuration

`SPSCleanVersions.ps1` accepts its configuration as JSON, from one of two mutually exclusive sources:

- **`-InputJson '<json string>'`** — an inline JSON string. Ideal for **Azure Automation Runbooks**, where the string is pasted directly into the runbook parameter field. This avoids all runbook limitations with array, switch, and boolean parameter types.
- **`-ConfigFile '<path>.json'`** — a path to a local JSON file with the same schema. Ideal for **local execution and testing**, where a versionable config file on disk is more convenient.

Both sources are parsed with `ConvertFrom-Json` and share the exact same schema, defaults, and validation.

> **Why JSON and not `.psd1`?** `ConvertFrom-Json` is fully supported in the Azure Automation sandbox, whereas reading a `.psd1` from disk with `Import-PowerShellDataFile` is not applicable there. Keeping a single JSON schema means the runbook input and the local file stay identical.

## JSON Schema

```json
{
  "SiteUrls": ["<string>"],
  "KeepMajorVersions": <integer>,
  "KeepMinorVersions": <integer>,
  "ClientId": "<string>",
  "ForceDeleteOldVersions": <boolean>,
  "DryRun": <boolean>,
  "VersionPolicyMode": "<string>",
  "ExpireVersionsAfterDays": <integer>,
  "ApplyTo": "<string>",
  "SiteScope": "<string>",
  "TenantAdminUrl": "<string>",
  "SiteFilter": "<string>",
  "EnableReport": <boolean>,
  "EnumerateLibraries": <boolean>,
  "LogRetentionDays": <integer>
}
```

## Properties

| Property | Type | Required | Default | Description |
|---|---|---|---|---|
| `SiteUrls` | string[] | Conditional | — | One or more SharePoint Site Collection URLs to process. **Required** when `SiteScope` is `Selected` (default); optional when `SiteScope` is `All`. |
| `KeepMajorVersions` | integer | No | `50` | Maximum number of major versions to retain. Maps to `-MajorVersions` in the site version policy modes. |
| `KeepMinorVersions` | integer | No | `0` | Maximum number of minor versions to retain. Set to `0` to disable minor versioning. Maps to `-MajorWithMinorVersions` in the site version policy modes. |
| `ClientId` | string | No | — | Azure AD App Registration Client ID used for authentication. Required for Interactive login (local) and optional for Managed Identity (Azure Automation). |
| `ForceDeleteOldVersions` | boolean | No | `false` | When `true`, submits a batch delete job via `New-PnPSiteFileVersionBatchDeleteJob` to remove file versions exceeding the configured limits. **Requires delegated user context** — automatically skipped in Azure Automation. |
| `DryRun` | boolean | No | `false` | When `true`, simulates all changes without applying them. Use this instead of `-WhatIf` when running as an Azure Automation Runbook. |
| `VersionPolicyMode` | string | No | `Legacy` | Version policy mechanism. `Legacy` keeps the per-library count-based `Set-PnPList` behaviour. `AutoExpiration`, `ExpireAfter`, `NoExpiration` and `InheritFromTenant` apply a site-level policy via `Set-PnPSiteVersionPolicy`. See [Version policy modes](#version-policy-modes). |
| `ExpireVersionsAfterDays` | integer | No | `0` | Number of days after which versions expire. Used by `ExpireAfter` (must be **>= 30**). `NoExpiration` forces `0`. |
| `ApplyTo` | string | No | `Both` | `New`, `Existing` or `Both` document libraries. Maps to `-ApplyToNewDocumentLibraries` / `-ApplyToExistingDocumentLibraries`. Site version policy modes only. |
| `SiteScope` | string | No | `Selected` | `Selected` processes `SiteUrls`. `All` enumerates **every** site collection in the tenant via `Get-PnPTenantSite`. Site version policy modes only (not `Legacy`). See [Tenant-wide scope](#tenant-wide-scope-sitescope-all). |
| `TenantAdminUrl` | string | Conditional | — | SharePoint admin center URL (e.g. `https://contoso-admin.sharepoint.com`). **Required** when `SiteScope` is `All`. |
| `SiteFilter` | string | No | — | Optional server-side `-Filter` passed to `Get-PnPTenantSite` to narrow the enumeration when `SiteScope` is `All` (e.g. `"Url -like 'sales'"`). |
| `EnableReport` | boolean | No | `true` | Write a local HTML report of the run to `Results/` (plus a machine-readable `SPSCleanVersions-<timestamp>.json` next to it). The report has **Site / Scope / Library / Outcome / Major / Minor / ExpireAfterDays / Detail** columns; Legacy mode reports one row per document library. **Local execution only** — no report is produced when running in Azure Automation. |
| `EnumerateLibraries` | boolean | No | `false` | Site version policy modes only. When `true`, also list the document libraries *in scope* for each site as informative `InScope` rows in the report. The policy applies to existing libraries via an **asynchronous server job**, so these rows carry no per-library Applied/Failed status. Adds a `Get-PnPList` call per site — leave off for large tenant-scale runs. |
| `LogRetentionDays` | integer | No | `180` | Prune `Logs/` and `Results/` files older than this many days (local only). `0` disables pruning. |

## Version policy modes

`VersionPolicyMode` selects how versioning is configured:

| Mode | Mechanism | Effect |
|---|---|---|
| `Legacy` (default) | `Set-PnPList` per document library | Count-based major/minor limits, applied library by library. Backward compatible with earlier versions. |
| `AutoExpiration` | `Set-PnPSiteVersionPolicy -EnableAutoExpirationVersionTrim $true` | SharePoint automatically trims versions (Microsoft-recommended). |
| `ExpireAfter` | `Set-PnPSiteVersionPolicy -EnableAutoExpirationVersionTrim $false -ExpireVersionsAfterDays <n> -MajorVersions <n> [-MajorWithMinorVersions <n>]` | Versions expire after `ExpireVersionsAfterDays` (>= 30) and are capped by the major/minor counts. |
| `NoExpiration` | `Set-PnPSiteVersionPolicy -EnableAutoExpirationVersionTrim $false -ExpireVersionsAfterDays 0 -MajorVersions <n>` | No expiration; only the major version count caps history. |
| `InheritFromTenant` | `Set-PnPSiteVersionPolicy -InheritFromTenant` | Clears the site-level setting so libraries follow the tenant default. |

> **Important:** the site version policy modes require **SharePoint Administrator** privileges and a PnP connection that can call `Set-PnPSiteVersionPolicy`. Applying to **existing** libraries submits a background request that may take time to complete across a large site.

> **Drift-based apply:** in the site version policy modes, the script first reads the current policy with `Get-PnPSiteVersionPolicy` and only calls `Set-PnPSiteVersionPolicy` when it differs from the desired settings (a *drift*). Sites that already match are logged as compliant and skipped. If the current policy cannot be read for a *transient* reason, the script fails safe and applies the policy anyway — **except** for a permission failure (*access denied*), which is never masked as drift: the site is recorded as `AccessDenied` and skipped (see [Site collection administrator requirement](#site-collection-administrator-requirement-delegated-runs)).

> **ℹ️ Azure Automation / app-only behaviour:** `Get-PnPSiteVersionPolicy` and `Set-PnPSiteVersionPolicy` are documented as requiring a **delegated** context that is **site collection administrator**. Live testing with a **Managed Identity** (app-only) confirmed that **reads work** — `Get-PnPTenantSite` (for `SiteScope: All`) and `Get-PnPSiteVersionPolicy` (drift detection) both succeed app-only. **Writes** to existing document libraries may still require a delegated context and can fail with *"Attempted to perform an unauthorized operation"*; the script emits a warning in a runbook. If a write fails app-only, run the site version policy modes **interactively / locally** with a SharePoint Administrator account, or use the tenant-level `Set-PnPTenant` settings. The default `Legacy` mode is unaffected.

## Site collection administrator requirement (delegated runs)

For **local / interactive** runs the connection is **delegated**: the effective rights are the
**intersection** of the app registration's delegated scope (`AllSites.FullControl` /
`Sites.FullControl.All`) **and** the signed-in user's own rights **on each target site**. A
full-control app scope is therefore not enough on its own — the signed-in account must also be a
**site collection administrator** on every site you process. Being a tenant **SharePoint
Administrator** grants management of the tenant and the admin center, but it does **not** by itself
grant content access to an arbitrary site collection.

When the account lacks rights on a site, SharePoint returns *"Attempted to perform an unauthorized
operation"*. The script does **not** treat this as a transient error or as a policy drift: it
records the site with a dedicated **`AccessDenied`** outcome, prints an actionable warning, and
**continues with the other sites**. At the end of the run a summary line reports the access-denied
count and the HTML report flags those sites (orange `AccessDenied` badge). To fix, add the account
as a **site collection administrator** on the affected sites (SharePoint admin center → *Sites* →
*Active sites* → select the site → *Membership* → *Site admins*, or `Set-PnPTenantSite -Url <site>
-Owners <upn>`) and re-run.

> Running as an **app-only** identity (Managed Identity or certificate) removes this
> intersection — the app itself is the identity — so the site-admin requirement does not apply
> there. It is specific to delegated (interactive/local) runs. An opt-in option to add the site
> collection administrator automatically is planned for a future release.

## Tenant-wide scope (SiteScope: All)

By default (`SiteScope: Selected`) the script only processes the sites listed in `SiteUrls`. Set `SiteScope: All` to apply a site version policy across **every site collection in the tenant**. The script connects to `TenantAdminUrl`, enumerates sites with `Get-PnPTenantSite` (OneDrive personal sites excluded), optionally narrowed by `SiteFilter`, and applies the policy to each — still gated by the per-site drift check, so unchanged sites are skipped.

> **⚠️ Tenant-wide impact:** `SiteScope: All` can touch **thousands** of site collections and, when `ApplyTo` includes `Existing`, submit a background version-trim job on each. **Always run with `"DryRun": true` first** to review the scope, and consider narrowing with `SiteFilter`. This scope is only supported with the site version policy modes (not `Legacy`), and requires **SharePoint Administrator** privileges plus a delegated (interactive/local) context.

```json
{
  "SiteScope": "All",
  "TenantAdminUrl": "https://contoso-admin.sharepoint.com",
  "SiteFilter": "Url -like 'sales'",
  "VersionPolicyMode": "ExpireAfter",
  "ExpireVersionsAfterDays": 180,
  "KeepMajorVersions": 100,
  "ApplyTo": "Both",
  "DryRun": true
}
```

## Logging and reports

Every run produces a summary of what happened per site (**Applied** / **WouldApply** / **Skipped** / **Compliant** / **AccessDenied** / **Failed**). The final summary line reports the counts, e.g. `--- SPSCleanVersions finished: N site(s), M result(s) — X would apply, Y skipped/compliant, Z access-denied, W failed ---`.

- **Local execution:** a transcript is written to a `Logs/` folder and a self-contained HTML report (summary cards + a filterable table) to a `Results/` folder, both next to the script. A machine-readable JSON (`SPSCleanVersions-<timestamp>.json`) is written alongside the HTML. Files older than `LogRetentionDays` (default 180) are pruned automatically. Set `"EnableReport": false` to skip the HTML report.
- **Azure Automation:** there is no persistent filesystem, so the **HTML report is not produced**. The per-site actions are visible in the job output (`Write-Output`/`Write-Warning`) and the run ends with a summary line (`--- SPSCleanVersions finished: ... ---`).

The report values are HTML-encoded, and a `DryRun` badge is shown when the run is a simulation. Rows with the `Failed` or `AccessDenied` outcome are highlighted so problem sites stand out.

### Resilience: single sign-in, throttling and permission handling

- **Single interactive sign-in.** A local run signs in **once**. That interactive connection drives the tenant enumeration (`Get-PnPTenantSite` for `SiteScope: All`) and its **SharePoint-audience delegated token** is reused for every site, so you are prompted a single time — not once per site — on every platform, including macOS.
- **Throttling-aware retry.** SharePoint calls are wrapped with a retry that honours the server `Retry-After` hint on HTTP 429/503 (capped at 300s) and otherwise uses exponential backoff — important for tenant-scale runs.
- **Fail fast on structural errors.** Authentication/token failures and permission (`AccessDenied`) failures are **not** retried (retrying cannot recover them); they are surfaced immediately with actionable guidance, and permission failures are recorded as `AccessDenied` per site (see [Site collection administrator requirement](#site-collection-administrator-requirement-delegated-runs)).

## Examples

### Apply an ExpireAfter site version policy

Versions expire after 180 days, keeping up to 100 major versions, applied to new and existing document libraries.

```json
{
  "SiteUrls": [
    "https://contoso.sharepoint.com/sites/News"
  ],
  "VersionPolicyMode": "ExpireAfter",
  "ExpireVersionsAfterDays": 180,
  "KeepMajorVersions": 100,
  "ApplyTo": "Both"
}
```

```powershell
.\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/News"],"VersionPolicyMode":"ExpireAfter","ExpireVersionsAfterDays":180,"KeepMajorVersions":100}'
```

### Inherit the tenant version policy

Clears any site-level override so document libraries follow the tenant default.

```json
{
  "SiteUrls": [
    "https://contoso.sharepoint.com/sites/News"
  ],
  "VersionPolicyMode": "InheritFromTenant"
}
```

### File-based configuration (local execution)

Save a JSON file (e.g. `Config/contoso-PROD.json`, ignored by git) and pass it with `-ConfigFile`. A ready-to-copy template is provided in [`Config/SPSCleanVersions.example.json`](https://github.com/luigilink/SPSCleanVersions/blob/main/Config/SPSCleanVersions.example.json).

```json
{
  "SiteUrls": [
    "https://contoso.sharepoint.com/sites/News",
    "https://contoso.sharepoint.com/sites/HR"
  ],
  "KeepMajorVersions": 50,
  "KeepMinorVersions": 0,
  "ClientId": "8cef7dae-500b-45ae-a717-b388ed2e7f69",
  "ForceDeleteOldVersions": false,
  "DryRun": true
}
```

```powershell
.\SPSCleanVersions.ps1 -ConfigFile '.\Config\contoso-PROD.json'
```

> **Tip:** Only `Config/*.example.json` is tracked in git. Real config files (`Config/*.json`) are ignored so your site URLs and Client IDs never land in version control.

### Minimal: Single site with defaults

Processes one site, keeping 50 major versions and 0 minor versions (defaults).

```json
{
  "SiteUrls": [
    "https://contoso.sharepoint.com/sites/News"
  ]
}
```

```powershell
.\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/News"]}'
```

### Multiple sites with custom retention

Processes two sites, keeping 100 major versions and 10 minor versions.

```json
{
  "SiteUrls": [
    "https://contoso.sharepoint.com/sites/News",
    "https://contoso.sharepoint.com/sites/HR"
  ],
  "KeepMajorVersions": 100,
  "KeepMinorVersions": 10
}
```

```powershell
.\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/News","https://contoso.sharepoint.com/sites/HR"],"KeepMajorVersions":100,"KeepMinorVersions":10}'
```

### Dry run (simulation mode)

Simulates the operation without making any changes. Ideal for testing.

```json
{
  "SiteUrls": [
    "https://contoso.sharepoint.com/sites/News"
  ],
  "KeepMajorVersions": 70,
  "DryRun": true
}
```

```powershell
.\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/News"],"KeepMajorVersions":70,"DryRun":true}'
```

### Force delete old versions (local only)

Submits a batch delete job to remove old file versions exceeding the configured limits. This only works with delegated user context (local/interactive login).

```json
{
  "SiteUrls": [
    "https://contoso.sharepoint.com/sites/News"
  ],
  "KeepMajorVersions": 50,
  "ForceDeleteOldVersions": true
}
```

```powershell
.\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/News"],"KeepMajorVersions":50,"ForceDeleteOldVersions":true}'
```

### Full configuration with Client ID

All properties specified, including a custom Client ID for authentication.

```json
{
  "SiteUrls": [
    "https://contoso.sharepoint.com/sites/News",
    "https://contoso.sharepoint.com/sites/HR",
    "https://contoso.sharepoint.com/sites/Finance"
  ],
  "KeepMajorVersions": 100,
  "KeepMinorVersions": 5,
  "ClientId": "8cef7dae-500b-45ae-a717-b388ed2e7f69",
  "ForceDeleteOldVersions": false,
  "DryRun": true
}
```

```powershell
.\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/News","https://contoso.sharepoint.com/sites/HR","https://contoso.sharepoint.com/sites/Finance"],"KeepMajorVersions":100,"KeepMinorVersions":5,"ClientId":"8cef7dae-500b-45ae-a717-b388ed2e7f69","ForceDeleteOldVersions":false,"DryRun":true}'
```

### Azure Automation Runbook

When running as an Azure Automation Runbook, paste the JSON string directly into the `InputJson` parameter field in the Azure portal:

```json
{
  "SiteUrls": [
    "https://contoso.sharepoint.com/sites/News"
  ],
  "KeepMajorVersions": 70,
  "ClientId": "8cef7dae-500b-45ae-a717-b388ed2e7f69",
  "DryRun": true
}
```

> **Tip:** The script automatically detects the Azure Automation environment and connects via Managed Identity. No interactive login is needed.
> **Note:** `ForceDeleteOldVersions` is automatically skipped in Azure Automation because the `New-PnPSiteFileVersionBatchDeleteJob` API requires delegated user context, which is not available with Managed Identity.

## Troubleshooting

### Error: `JSON property 'SiteUrls' is required` even though SiteUrls is present

When pasting the value into the Azure Automation **InputJson** field, paste the **raw JSON object only** — do **not** wrap it in the surrounding single quotes used on a PowerShell command line.

- ❌ Wrong (portal field): `'{"SiteUrls":["https://contoso.sharepoint.com/teams/CSSC"],"KeepMajorVersions":100}'`
- ✅ Correct (portal field): `{"SiteUrls":["https://contoso.sharepoint.com/teams/CSSC"],"KeepMajorVersions":100}`

Since v3.0.0 the script auto-strips a single wrapping pair of single quotes and validates that the parsed value is a JSON object, so this mistake now yields a clear message instead of the misleading `SiteUrls is required`.

### Error: `Invalid JSON input ... Invalid property identifier character`

The value contains **curly / smart quotes** (`“ ”`) instead of straight double quotes (`"`), typically after copying from Teams, Outlook, or Word. Retype the double quotes as straight quotes, or paste from a plain-text editor.

### Warning: `Access denied on <site> ... not a site collection administrator` (outcome `AccessDenied`)

A **delegated** run reports this when the signed-in account has no rights on that specific site,
even though the app registration carries `AllSites.FullControl`: delegated rights are the
intersection of the app scope **and** the user's rights on the site (see [Site collection
administrator requirement](#site-collection-administrator-requirement-delegated-runs)). Being a
tenant SharePoint Administrator is **not** sufficient by itself. The site is skipped (not applied,
not failed) and the run continues.

**Fix:** add the account as a **site collection administrator** on the affected site(s) and re-run:

```powershell
Set-PnPTenantSite -Url "https://contoso.sharepoint.com/sites/<site>" -Owners "<upn>"
```

or via the SharePoint admin center → *Sites* → *Active sites* → select the site → *Membership* →
*Site admins*.
