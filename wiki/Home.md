# SPSCleanVersions

SPSCleanVersions is a PowerShell script tool to clean Version History in your SharePoint Tenant. Optimize your storage costs by managing major and minor versions across libraries and lists.

## Key Features

* **Two Config Sources:** Pass configuration inline via `-InputJson` (ideal for Azure Automation Runbooks) or from a local JSON file via `-ConfigFile` (ideal for local execution and testing). Both share the same JSON schema and validation.
* **Simulation Mode:** Use `"DryRun": true` in the JSON (or `-WhatIf` locally) to preview changes safely.
* **Flexible Retention:** Define custom thresholds for major and minor versions.
* **Site Version Policy:** Set `"VersionPolicyMode"` to `AutoExpiration`, `ExpireAfter`, `NoExpiration` or `InheritFromTenant` to apply a site-level policy via `Set-PnPSiteVersionPolicy` — including version expiration (`ExpireVersionsAfterDays`) and applying to new and/or existing libraries. The default `Legacy` mode keeps the per-library count-based behaviour.
* **Idempotent (drift-based):** in the site version policy modes, the policy is applied only where the current settings actually differ, so compliant sites are skipped.
* **Tenant-wide Scope:** Set `"SiteScope": "All"` with a `TenantAdminUrl` to enumerate and process every site collection via `Get-PnPTenantSite`.
* **Logging & HTML Report:** each run produces a per-site summary — a transcript and self-contained HTML report locally, or the report emitted to the job output stream in Azure Automation.
* **Multi-Site Processing:** Pass multiple Site Collection URLs in the `SiteUrls` JSON array to process them in a single execution.
* **Force Delete Old Versions:** Set `"ForceDeleteOldVersions": true` to trigger a batch delete job via `New-PnPSiteFileVersionBatchDeleteJob`. Requires delegated user context (automatically skipped in Azure Automation).
* **Azure Automation Ready:** Single string parameter avoids all runbook type limitations (arrays, switches, booleans).

For details on usage, configuration, and parameters, explore the links below:

* [Getting Started](./Getting-Started)
* [Configuration](./Configuration)
* [Usage](./Usage)
