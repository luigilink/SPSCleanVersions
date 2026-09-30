# SPSCleanVersions - Release Notes

## [Unreleased]

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

## [3.3.0] - 2026-09-29

### Added

- SPSCleanVersions.ps1
  - **JIT site collection admin (`AddSiteCollectionAdmin`)** — delegated runs can process sites the
    operator does not administer: the operator is temporarily added as site collection admin
    (through the admin center), the site is processed, then the grant is revoked. Already-admin
    sites are left untouched. Requires `TenantAdminUrl` + the SharePoint Administrator role (checked
    up front); ignored under app-only.
  - **`-CleanupAdminsOnly`** — revokes grants left by an interrupted run, replaying the grant state
    file (`Logs/SPSCleanVersions-admins-*.jsonl`). Idempotent; the run summary reports grant/revoke
    counts and warns on any revoke failure.

### Changed

- CI / Release
  - The release ZIP now contains `SPSCleanVersions.ps1` at its **root** (not under `scripts/`),
    aligned with SPSWakeUp packaging. No change to the script or the PowerShell Gallery package.


A full list of changes in each version can be found in the [change log](CHANGELOG.md)
