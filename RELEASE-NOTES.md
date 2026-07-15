# SPSCleanVersions - Release Notes

## [3.0.0] - 2026-07-15

### Added

- SPSCleanVersions.ps1
  - Add `-ConfigFile` parameter to load configuration from a local JSON file,
    complementing the inline `-InputJson` string used by Azure Automation
    Runbooks. Both input sources share the same JSON schema, parsing and
    validation, so defaults and required-property checks stay identical
  - `-InputJson` and `-ConfigFile` are mutually exclusive parameter sets, with
    `InlineJson` as the default set
  - Harden JSON parsing: trim the input, strip an accidental wrapping pair of
    single quotes (a common copy/paste mistake from a command line), and reject
    non-object JSON (string/scalar) with a clear, actionable message instead of
    the misleading `SiteUrls is required` error
- Config/SPSCleanVersions.example.json
  - Add example JSON configuration template
- Wiki Documentation
  - Document the file-based configuration workflow

### Changed

- .gitignore
  - Track only `Config/*.example.json`; ignore real configs and local run
    artifacts (`Logs/`, `Results/`)

A full list of changes in each version can be found in the [change log](CHANGELOG.md)
