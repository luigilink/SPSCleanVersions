# SPSCleanVersions - Release Notes

## [3.5.0] - 2026-11-18

### Added

- SPSCleanVersions.ps1
  - **Trim control for large batches.** `TrimJobMode` (`Auto`, `Age`, `Count`, `Automatic`) chooses the trim
    job that `ForceDeleteOldVersions` submits, and `TrimOnly` submits only that job, without reading or
    changing the site version policy. A job applies one criterion and a site accepts one job at a time, so a
    run can do one pass per criterion (for example age first, then count).
  - **`-ExportJobStatus`.** Reads the trim job state of each site and writes a dated CSV (status, times,
    versions deleted, storage released, error), one row per site, including the sites that could not be
    read. Run it soon after the jobs finish: SharePoint does not keep the state of a completed job
    indefinitely. Local (delegated) runs only.
  - Trim job submissions now appear in the report and in the JSON results (`Submitted`, `WouldSubmit`,
    `InProgress`, `Failed`).

### Changed

- SPSCleanVersions.ps1
  - **The age-based trim job is refused when its cut-off falls before January 1, 2023**, because that job
    always deletes every version created before that date. The run stops before the sign-in and tells the
    largest value accepted.

### Fixed

- SPSCleanVersions.ps1
  - **`ForceDeleteOldVersions` trims existing versions according to the version policy mode.** It used to
    always submit a count-limit trim job, so with `ExpireAfter` the versions older than
    `ExpireVersionsAfterDays` were never removed (a policy only governs versions created afterwards).
    It now submits an **age-based** job for `ExpireAfter` and the **automatic** job for `AutoExpiration`;
    `Legacy`, `NoExpiration` and `InheritFromTenant` keep the count-limit job. The run log and the
    `DryRun` message tell which kind of job is submitted. See the new *Trimming existing versions*
    section of the wiki (permanent deletion, asynchronous jobs, final pass for age-based rules).

A full list of changes in each version can be found in the [change log](CHANGELOG.md)
