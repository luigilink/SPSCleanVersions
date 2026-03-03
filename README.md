# SPSCleanVersions

![Latest release date](https://img.shields.io/github/release-date/luigilink/SPSCleanVersions.svg?style=flat)
![Total downloads](https://img.shields.io/github/downloads/luigilink/SPSCleanVersions/total.svg?style=flat)  
![Issues opened](https://img.shields.io/github/issues/luigilink/SPSCleanVersions.svg?style=flat)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-2.1-4baaaa.svg)](code_of_conduct.md)
[![SPSCleanVersions CI Pester Tests](https://github.com/luigilink/SPSCleanVersions/actions/workflows/pester.yml/badge.svg)](https://github.com/luigilink/SPSCleanVersions/actions/workflows/pester.yml)

## Description

SPSCleanVersions is a PowerShell script tool to clean Version History in your SharePoint Tenant. Optimize your storage costs by managing major and minor versions across libraries and lists.

[**Download the latest release here!**](https://github.com/luigilink/SPSCleanVersions/releases/latest)

## Key Features

* **Single JSON Input:** All configuration is passed via a single `-InputJson` string parameter, ensuring full compatibility with Azure Automation Runbooks.
* **Simulation Mode:** Use `"DryRun": true` in the JSON (or `-WhatIf` locally) to preview changes safely.
* **Flexible Retention:** Define custom thresholds for major and minor versions.
* **Multi-Site Processing:** Pass multiple Site Collection URLs in the `SiteUrls` JSON array to process them in a single execution.
* **Force Delete Old Versions:** Set `"ForceDeleteOldVersions": true` to trigger a batch delete job via `New-PnPSiteFileVersionBatchDeleteJob`. Requires delegated user context (automatically skipped in Azure Automation).
* **Azure Automation Ready:** Single string parameter avoids all runbook type limitations (arrays, switches, booleans).

## Requirements

### PowerShell 7.2+ (Core)

Requires PowerShell 7.2 or later with PSEdition Core. [Installation guide](https://learn.microsoft.com/en-us/powershell/scripting/install/install-powershell?view=powershell-7.5).

### Module PnP.PowerShell (>= 2.12.0)

This tool relies on the PnP.PowerShell module version 2.12.0 or later. [Installation guide](https://pnp.github.io/powershell/articles/installation.html).

### Permissions

* **Role:** SharePoint Administrator or Global Administrator.
* **API Permissions:** `Sites.FullControl.All` (when using App Registration).

## Documentation

For detailed usage, configuration, and getting started information, visit the [SPSCleanVersions Wiki](https://github.com/luigilink/SPSCleanVersions/wiki)

## Changelog

A full list of changes in each version can be found in the [change log](CHANGELOG.md)
