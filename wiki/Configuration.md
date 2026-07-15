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
  "ApplyTo": "<string>"
}
```

## Properties

| Property | Type | Required | Default | Description |
|---|---|---|---|---|
| `SiteUrls` | string[] | **Yes** | — | One or more SharePoint Site Collection URLs to process. |
| `KeepMajorVersions` | integer | No | `50` | Maximum number of major versions to retain. Maps to `-MajorVersions` in the site version policy modes. |
| `KeepMinorVersions` | integer | No | `0` | Maximum number of minor versions to retain. Set to `0` to disable minor versioning. Maps to `-MajorWithMinorVersions` in the site version policy modes. |
| `ClientId` | string | No | — | Azure AD App Registration Client ID used for authentication. Required for Interactive login (local) and optional for Managed Identity (Azure Automation). |
| `ForceDeleteOldVersions` | boolean | No | `false` | When `true`, submits a batch delete job via `New-PnPSiteFileVersionBatchDeleteJob` to remove file versions exceeding the configured limits. **Requires delegated user context** — automatically skipped in Azure Automation. |
| `DryRun` | boolean | No | `false` | When `true`, simulates all changes without applying them. Use this instead of `-WhatIf` when running as an Azure Automation Runbook. |
| `VersionPolicyMode` | string | No | `Legacy` | Version policy mechanism. `Legacy` keeps the per-library count-based `Set-PnPList` behaviour. `AutoExpiration`, `ExpireAfter`, `NoExpiration` and `InheritFromTenant` apply a site-level policy via `Set-PnPSiteVersionPolicy`. See [Version policy modes](#version-policy-modes). |
| `ExpireVersionsAfterDays` | integer | No | `0` | Number of days after which versions expire. Used by `ExpireAfter` (must be **>= 30**). `NoExpiration` forces `0`. |
| `ApplyTo` | string | No | `Both` | `New`, `Existing` or `Both` document libraries. Maps to `-ApplyToNewDocumentLibraries` / `-ApplyToExistingDocumentLibraries`. Site version policy modes only. |

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
