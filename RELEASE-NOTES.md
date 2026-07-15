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
  - **Azure Automation runbook compatibility:** explicitly `Import-Module PnP.PowerShell`,
    because in a runbook the `#Requires -Modules` directive does not import the module and
    command auto-loading is unreliable in the sandbox (`Connect-PnPOnline` was *"not
    recognized"*). Harmless locally
- Wiki Documentation
  - Add a PowerShell / PnP.PowerShell version compatibility matrix (PnP 3.x needs
    PowerShell 7.4; use a 7.4 Runtime Environment in Azure Automation, or PnP 2.12.x
    on a 7.2 runbook)

A full list of changes in each version can be found in the [change log](CHANGELOG.md)
