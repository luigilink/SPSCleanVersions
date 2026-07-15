# SPSCleanVersions - Release Notes

## [3.1.1] - 2026-07-15

### Fixed

- SPSCleanVersions.ps1
  - **Azure Automation runbook compatibility:** remove the parameter sets introduced
    with `-ConfigFile`. Azure Automation rejects runbooks that use parameter sets
    (*"Parameter sets in runbooks are not supported in this release"*), which prevented
    the script from starting as a runbook. `-InputJson` and `-ConfigFile` are now plain
    optional parameters whose mutual exclusivity is validated in the body (exactly one
    is required)

A full list of changes in each version can be found in the [change log](CHANGELOG.md)
