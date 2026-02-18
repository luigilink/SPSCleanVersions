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

* **Simulation Mode:** Support for `-WhatIf` to preview changes and estimate potential storage gains safely.
* **Flexible Retention:** Define custom thresholds for major and minor versions.
* **Granular Targeting:** Process a single Site Collection or multiple sites via a list, avoiding tenant-wide timeouts.

## Requirements

### PowerShell 7.x

Required for performance and class-based resource management. [Installation guide](https://learn.microsoft.com/en-us/powershell/scripting/install/install-powershell?view=powershell-7.5).

### Module PnP.PowerShell

This tool relies on the PnP.PowerShell module. [Installation guide](https://pnp.github.io/powershell/articles/installation.html).

### Permissions

* **Role:** SharePoint Administrator or Global Administrator.
* **API Permissions:** `Sites.FullControl.All` (when using App Registration).

## Documentation

For detailed usage, configuration, and getting started information, visit the [SPSCleanVersions Wiki](https://github.com/luigilink/SPSCleanVersions/wiki)

## Changelog

A full list of changes in each version can be found in the [change log](CHANGELOG.md)
