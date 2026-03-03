# Configuration

`SPSCleanVersions.ps1` accepts a single `-InputJson` parameter containing a JSON string with all configuration. This design ensures full compatibility with Azure Automation Runbooks, which have limitations with array, switch, and boolean parameter types.

## JSON Schema

```json
{
  "SiteUrls": ["<string>"],
  "KeepMajorVersions": <integer>,
  "KeepMinorVersions": <integer>,
  "ClientId": "<string>",
  "ForceDeleteOldVersions": <boolean>,
  "DryRun": <boolean>
}
```

## Properties

| Property | Type | Required | Default | Description |
|---|---|---|---|---|
| `SiteUrls` | string[] | **Yes** | — | One or more SharePoint Site Collection URLs to process. |
| `KeepMajorVersions` | integer | No | `50` | Maximum number of major versions to retain per document. |
| `KeepMinorVersions` | integer | No | `0` | Maximum number of minor versions to retain per document. Set to `0` to disable minor versioning. |
| `ClientId` | string | No | — | Azure AD App Registration Client ID used for authentication. Required for Interactive login (local) and optional for Managed Identity (Azure Automation). |
| `ForceDeleteOldVersions` | boolean | No | `false` | When `true`, submits a batch delete job via `New-PnPSiteFileVersionBatchDeleteJob` to remove file versions exceeding the configured limits. **Requires delegated user context** — automatically skipped in Azure Automation. |
| `DryRun` | boolean | No | `false` | When `true`, simulates all changes without applying them. Use this instead of `-WhatIf` when running as an Azure Automation Runbook. |

## Examples

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
