# SPSCleanVersions

SPSCleanVersions is a PowerShell script tool to clean Version History in your SharePoint Tenant. Optimize your storage costs by managing major and minor versions across libraries and lists.

## Key Features

* **Simulation Mode:** Support for `-WhatIf` to preview changes and estimate potential storage gains safely.
* **Flexible Retention:** Define custom thresholds for major and minor versions.
* **Granular Targeting:** Process a single Site Collection or multiple sites via a list, avoiding tenant-wide timeouts.
* **Force Delete Old Versions:** Use `-ForceDeleteOldVersions` to trigger a batch delete job via `New-PnPSiteFileVersionBatchDeleteJob`, trimming file versions exceeding the configured major/minor count limits.

For details on usage, configuration, and parameters, explore the links below:

* [Getting Started](./Getting-Started)
* [Usage](./Usage)
