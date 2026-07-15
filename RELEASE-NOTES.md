# SPSCleanVersions - Release Notes

## [3.1.2] - 2026-07-15

### Changed

- SPSCleanVersions.ps1
  - The HTML report is now **local-only**. In Azure Automation it is no longer emitted
    into the job output stream (dumping the full HTML made the runbook log unreadable);
    only the run summary line is printed. Local execution is unchanged — the report is
    still written to `Results/`

A full list of changes in each version can be found in the [change log](CHANGELOG.md)
