# SPSCleanVersions - Release Notes

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
