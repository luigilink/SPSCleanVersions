# Getting Started

## Requirements

### PowerShell 7.2+ (Core)

Requires PowerShell 7.2 or later with PSEdition Core. [Installation guide](https://learn.microsoft.com/en-us/powershell/scripting/install/install-powershell?view=powershell-7.5).

### Module PnP.PowerShell (>= 2.12.0)

This tool relies on the PnP.PowerShell module version 2.12.0 or later. [Installation guide](https://pnp.github.io/powershell/articles/installation.html).

> **⚠️ PowerShell / PnP.PowerShell version matrix.** PnP.PowerShell **3.x requires PowerShell 7.4+**, while the 2.x line supports PowerShell 7.2. Make sure the two match, especially in **Azure Automation**:
>
> | Runbook / host runtime | Compatible PnP.PowerShell | Notes |
> |---|---|---|
> | PowerShell **7.4** (Runtime Environment) | **3.x** (e.g. 3.3.0) | Recommended. All modes work (Legacy + site version policy + tenant scope). |
> | PowerShell **7.2** (classic runbook) | **2.12.x** | Legacy mode works; the newer site version policy cmdlets may be unavailable. |
>
> A mismatch (e.g. PnP 3.x on a 7.2 runbook) fails at startup with *"The module requires a minimum PowerShell version of '7.4.0'"*. In Azure Automation, create a **PowerShell 7.4 Runtime Environment** and import PnP.PowerShell 3.x into it.

### Permissions

* **Role:** SharePoint Administrator or Global Administrator.
* **API Permissions:** `Sites.FullControl.All` (when using App Registration).

## Installation

1. [Download the latest release](https://github.com/luigilink/SPSCleanVersions/releases/latest) and unzip to a directory on your machine tool.

## Next Step

For JSON parameter configuration and examples, go to the [Configuration](./Configuration) page.
For usage details, go to the [Usage](./Usage) page.

## Change log

A full list of changes in each version can be found in the [change log](https://github.com/luigilink/SPSCleanVersions/blob/main/CHANGELOG.md).
