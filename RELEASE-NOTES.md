# SPSCleanVersions - Release Notes

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
