<#PSScriptInfo
    .VERSION 2.0.0

    .GUID 7ecf4acd-17c4-4c50-be79-1fcf2b6611fe

    .AUTHOR luigilink (Jean-Cyril DROUHIN)

    .COPYRIGHT

    .TAGS
    script powershell sharepoint version history cleanup

    .LICENSEURI
    https://github.com/luigilink/SPSCleanVersions/blob/main/LICENSE

    .PROJECTURI
    https://github.com/luigilink/SPSCleanVersions

    .ICONURI

    .EXTERNALMODULEDEPENDENCIES

    .REQUIREDSCRIPTS

    .EXTERNALSCRIPTDEPENDENCIES

    .RELEASENOTES

    .PRIVATEDATA
#>

<#
    .SYNOPSIS
    SPSCleanVersions - Clean Version History in SharePoint Online.

    .DESCRIPTION
    A script tool to clean Version History in your SharePoint Tenant.
    Optimize your storage costs by managing major and minor versions across libraries and lists.
    Compatible with Local execution and Azure Automation Runbooks.
    Accepts a single JSON string parameter for full Azure Automation Runbook compatibility.

    .PARAMETER InputJson
    A JSON string containing all configuration. Supported properties:
      - SiteUrls              (string array, required) — Site Collection URLs to process.
      - KeepMajorVersions     (integer, optional, default: 50) — Number of major versions to keep.
      - KeepMinorVersions     (integer, optional, default: 0) — Number of minor versions to keep.
      - ClientId              (string, optional) — Azure AD App Registration Client ID.
      - ForceDeleteOldVersions (boolean, optional, default: false) — Trigger batch delete of old file versions.
      - DryRun                (boolean, optional, default: false) — Simulate changes without applying them.

    .EXAMPLE
    .\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/site1"],"KeepMajorVersions":100,"KeepMinorVersions":10}'
    Cleans version history for the specified site, keeping 100 major versions and 10 minor versions.

    .EXAMPLE
    .\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/site1","https://contoso.sharepoint.com/sites/site2"],"KeepMajorVersions":50,"DryRun":true}'
    Simulates the operation on multiple sites without making changes.

    .NOTES
    FileName:	SPSCleanVersions.ps1
    Author:		Jean-Cyril DROUHIN
    Date:		March 3, 2026
    Version:	2.0.0

    .LINK
    https://spjc.fr/
    https://github.com/luigilink/SPSCleanVersions
#>
#Requires -Version 7.2
#Requires -PSEdition Core
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '2.12.0' }

[CmdletBinding(SupportsShouldProcess)]
param
(
    [Parameter(Mandatory = $true, HelpMessage = "JSON string containing all configuration (SiteUrls, KeepMajorVersions, KeepMinorVersions, ClientId, ForceDeleteOldVersions, DryRun)")]
    [ValidateNotNullOrEmpty()]
    [System.String]
    $InputJson
)

#region --- Parse and validate JSON input ---
try {
    $config = $InputJson | ConvertFrom-Json -ErrorAction Stop
}
catch {
    throw "Invalid JSON input: $($_.Exception.Message)"
}

# Required: SiteUrls
if (-not $config.PSObject.Properties['SiteUrls'] -or
    $null -eq $config.SiteUrls -or
    @($config.SiteUrls).Count -eq 0) {
    throw "JSON property 'SiteUrls' is required and must contain at least one URL."
}
[string[]]$SiteUrls = @($config.SiteUrls)

# Optional with defaults
[int]$KeepMajorVersions      = if ($config.PSObject.Properties['KeepMajorVersions'])      { $config.KeepMajorVersions }      else { 50 }
[int]$KeepMinorVersions      = if ($config.PSObject.Properties['KeepMinorVersions'])      { $config.KeepMinorVersions }      else { 0 }
[string]$ClientId             = if ($config.PSObject.Properties['ClientId'])               { $config.ClientId }               else { '' }
[bool]$ForceDeleteOldVersions = if ($config.PSObject.Properties['ForceDeleteOldVersions']) { $config.ForceDeleteOldVersions } else { $false }
[bool]$DryRun                 = if ($config.PSObject.Properties['DryRun'])                 { $config.DryRun }                 else { $false }
#endregion

# When DryRun is specified, enable WhatIf mode so that ShouldProcess calls are simulated.
# This is required for Azure Automation Runbooks where -WhatIf common parameter is not supported.
if ($DryRun) {
    $WhatIfPreference = $true
}

Write-Output "--- Starting SPSCleanVersions ---"
if ($WhatIfPreference) {
    Write-Output "--- DryRun/WhatIf mode enabled: no changes will be applied ---"
}

# Disable PnP PowerShell update check to avoid interactive prompts in non-interactive environments (Azure Automation).
$env:PNPPOWERSHELL_UPDATECHECK = "false"

function Test-IsAzureAutomation {
    # In PS7.x, Azure Automation exposes several env vars (sandbox + managed identity endpoints).
    # Use multiple signals, not a single one.
    return (
        -not [string]::IsNullOrEmpty($env:AZUREPS_HOST_ENVIRONMENT) -or
        -not [string]::IsNullOrEmpty($env:AUTOMATION_ASSET_SANDBOX_ID) -or
        -not [string]::IsNullOrEmpty($env:AUTOMATION_ASSET_ENDPOINT) -or
        -not [string]::IsNullOrEmpty($env:MSI_ENDPOINT) -or
        -not [string]::IsNullOrEmpty($env:IDENTITY_ENDPOINT)
    )
}

foreach ($SiteUrl in $SiteUrls) {
    Write-Output "Processing Site: $SiteUrl"

    try {
        # Environment: Local vs Azure Automation
        if (Test-IsAzureAutomation) {
            Write-Output "Running in Azure Automation. Connecting via Managed Identity..."
            if (-not [string]::IsNullOrEmpty($ClientId)) {
                Connect-PnPOnline -Url $SiteUrl -ManagedIdentity -ClientId $ClientId
            }
            else {
                Connect-PnPOnline -Url $SiteUrl -ManagedIdentity
            }
        }
        else {
            Write-Output "Running locally. Connecting via Interactive login..."
            Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
        }

        # Get all Lists in the Site
        Write-Output "Retrieving lists from $SiteUrl..."
        $allLists = Get-PnPList
        $targetLists = $allLists | Where-Object {
            $_.Hidden -eq $false -and
            $_.EnableVersioning -eq $true -and
            $_.RootFolder.ServerRelativeUrl -notlike "*_catalogs*" -and
            $_.RootFolder.ServerRelativeUrl -notlike "*/SiteAssets*" -and
            $_.RootFolder.ServerRelativeUrl -notlike "*/SitePages*" -and
            $_.RootFolder.ServerRelativeUrl -notlike "*/Style Library*" -and
            $_.BaseTemplate -eq 101 # Document Libraries only
        }
        foreach ($list in $targetLists) {
            $minorDesired = ($KeepMinorVersions -gt 0)
            $changeNeeded = ($list.MajorVersionLimit -ne $KeepMajorVersions) -or
            ($list.EnableMinorVersions -ne $minorDesired) -or
            ($minorDesired -and ($list.MajorWithMinorVersionsLimit -ne $KeepMinorVersions)) -or
            (-not $minorDesired -and ($list.MajorWithMinorVersionsLimit -ne 0))

            if ($changeNeeded) {
                if ($PSCmdlet.ShouldProcess($list.Title, "Set versioning policy")) {
                    $p = @{
                        Identity         = "$($list.Title)"
                        EnableVersioning = $true
                        MajorVersions    = $KeepMajorVersions
                    }
                    if ($minorDesired) {
                        $p.EnableMinorVersions = $true
                        $p.MinorVersions = $KeepMinorVersions
                    }
                    else {
                        $p.EnableMinorVersions = $false
                    }
                    try {
                        Set-PnPList @p -ErrorAction Stop
                        Write-Output "`t$($list.Title) -> Major=$KeepMajorVersions; MinorEnabled=$minorDesired; MinorLimit=$KeepMinorVersions"
                    }
                    catch {
                        Write-Warning "`tFAILED $($list.Title): $($_.Exception.Message)"
                    }
                }
            }
            else {
                Write-Output "`t$($list.Title) already compliant"
            }
        }

        # Force deletion of old file version history
        if ($ForceDeleteOldVersions) {
            if ($PSCmdlet.ShouldProcess($SiteUrl, "Delete old file version history")) {
                if (Test-IsAzureAutomation) {
                    Write-Warning @"
Batch delete of file versions is NOT supported with app-only authentication.
This SharePoint API requires delegated user context.
Skipping New-PnPSiteFileVersionBatchDeleteJob for site: $SiteUrl
"@
                }
                else {
                    try {
                        Write-Output "`tStarting batch delete job for old file versions on $SiteUrl..."
                        $batchParams = @{
                            MajorVersionLimit           = $KeepMajorVersions
                            MajorWithMinorVersionsLimit = $KeepMinorVersions
                        }
                        New-PnPSiteFileVersionBatchDeleteJob @batchParams -Force -ErrorAction Stop
                        Write-Output "`tBatch delete job submitted successfully for $SiteUrl"
                    }
                    catch {
                        Write-Warning "`tFAILED to submit batch delete job for ${SiteUrl}: $($_.Exception.Message)"
                    }
                }
            }
        }
    }
    catch {
        Write-Error "Failed to process site $SiteUrl : $($_.Exception.Message)"
    }
    finally {
        Disconnect-PnPOnline
    }
}
