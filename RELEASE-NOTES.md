# SPSCleanVersions - Release Notes

## [3.3.1] - 2026-09-30

### Fixed

- SPSCleanVersions.ps1
  - **JIT `AddSiteCollectionAdmin` — false-positive "already admin" fixed.** The pre-flight check now
    confirms the operator's UPN is actually in the site collection admin list before deciding it is
    "already admin" (reading the list is no longer treated as proof). Sites where the operator could
    read the admin list without being an effective admin (e.g. a Microsoft 365 group member) are no
    longer skipped and left in `Access denied`. As a safety net, a site that still denies a privileged
    operation despite looking already-administered is granted JIT admin and retried once, then revoked.
  - **Batch delete "already in progress" no longer wastes retries.** When a prior file-version
    batch-delete job is still running on a site, the rejection is detected and skipped (a benign
    "already in progress" message) instead of being retried 5× with exponential backoff (~5 min/site)
    and logged as `FAILED`.


A full list of changes in each version can be found in the [change log](CHANGELOG.md)
