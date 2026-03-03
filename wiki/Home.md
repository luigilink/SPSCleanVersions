# SPSCleanVersions

SPSCleanVersions is a PowerShell script tool to clean Version History in your SharePoint Tenant. Optimize your storage costs by managing major and minor versions across libraries and lists.

## Key Features

* **Single JSON Input:** All configuration is passed via a single `-InputJson` string parameter, ensuring full compatibility with Azure Automation Runbooks.
* **Simulation Mode:** Use `"DryRun": true` in the JSON (or `-WhatIf` locally) to preview changes safely.
* **Flexible Retention:** Define custom thresholds for major and minor versions.
* **Multi-Site Processing:** Pass multiple Site Collection URLs in the `SiteUrls` JSON array to process them in a single execution.
* **Force Delete Old Versions:** Set `"ForceDeleteOldVersions": true` to trigger a batch delete job via `New-PnPSiteFileVersionBatchDeleteJob`. Requires delegated user context (automatically skipped in Azure Automation).
* **Azure Automation Ready:** Single string parameter avoids all runbook type limitations (arrays, switches, booleans).

For details on usage, configuration, and parameters, explore the links below:

* [Getting Started](./Getting-Started)
* [Configuration](./Configuration)
* [Usage](./Usage)
