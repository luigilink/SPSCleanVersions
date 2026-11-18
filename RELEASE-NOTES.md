# SPSCleanVersions - Release Notes

## [Unreleased]

### Fixed

- SPSCleanVersions.ps1
  - **`ForceDeleteOldVersions` trims existing versions according to the version policy mode.** It used to
    always submit a count-limit trim job, so with `ExpireAfter` the versions older than
    `ExpireVersionsAfterDays` were never removed (a policy only governs versions created afterwards).
    It now submits an **age-based** job for `ExpireAfter` and the **automatic** job for `AutoExpiration`;
    `Legacy`, `NoExpiration` and `InheritFromTenant` keep the count-limit job. The run log and the
    `DryRun` message tell which kind of job is submitted. See the new *Trimming existing versions*
    section of the wiki (permanent deletion, asynchronous jobs, final pass for age-based rules).

## [3.4.0] - 2026-09-30

### Added

- SPSCleanVersions.ps1
  - **Colourized console output** — per-site lines carry coloured status tags: site (cyan),
    grant/revoke (`[jit ]`, magenta), applied (`[ ok ]`, green), skipped/compliant (`[skip]`, grey),
    `[warn]` / `[deny]` — plus a framed end-of-run **summary** block. Colour is ANSI via `$PSStyle`,
    shown **only** on an interactive VT-capable console; Azure Automation, redirected/piped output and
    non-VT hosts get a plain ASCII fallback, and `Start-Transcript` records the `.log` without ANSI
    (`OutputRendering = 'Host'`). `results.json` and the HTML report are unchanged, and the
    machine-readable `--- SPSCleanVersions finished: … ---` anchor line is kept.
  - **Per-site progress bar** (`Write-Progress`, `Site X of N` + percent) for local/interactive runs;
    suppressed under Azure Automation.


A full list of changes in each version can be found in the [change log](CHANGELOG.md)
