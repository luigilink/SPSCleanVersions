<#PSScriptInfo
    .VERSION 1.1.1

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

    .EXAMPLE
    .\SPSCleanVersions.ps1 -SiteUrls "https://contoso.sharepoint.com/sites/site1,https://contoso.sharepoint.com/sites/site2" -KeepMajorVersions 100 -KeepMinorVersions 10 -WhatIf
    Cleans version history for the specified site collections, keeping 100 major versions and 10 minor versions.

    .NOTES
    FileName:	SPSCleanVersions.ps1
    Author:		Jean-Cyril DROUHIN
    Date:		February 16, 2026
    Version:	1.1.1

    .LINK
    https://spjc.fr/
    https://github.com/luigilink/SPSCleanVersions
#>

[CmdletBinding(SupportsShouldProcess)]
param
(
    [Parameter(Position = 0, Mandatory = $true, HelpMessage = "List of Site Collection URLs separated by commas")]
    [System.String[]]
    $SiteUrls,

    [Parameter(Position = 1, Mandatory = $false)]
    [System.UInt32]
    $KeepMajorVersions = 50,

    [Parameter(Position = 2, Mandatory = $false)]
    [System.UInt32]
    $KeepMinorVersions = 0,

    [Parameter(Position = 3, Mandatory = $false)]
    [System.String]
    $ClientId,  # Default ClientId for Microsoft Graph and SharePoint Online Management Shell

    [Parameter(Position = 4, Mandatory = $false, HelpMessage = "Force deletion of old file version history using New-PnPSiteFileVersionBatchDeleteJob")]
    [switch]
    $ForceDeleteOldVersions
)

Write-Host "--- Starting SPSCleanVersions ---" -ForegroundColor Cyan

foreach ($SiteUrl in $SiteUrls) {
    Write-Host "Processing Site: $SiteUrl" -ForegroundColor Yellow
    
    try {
        # Environnement : Local vs Azure Automation
        if ($null -ne $env:AUTOMATION_ASSET_NAME) {
            Write-Output "Running in Azure Automation. Connecting via Managed Identity..."
            if ($null -ne $ClientId) {
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
                        Write-Host "`t$($list.Title) -> Major=$KeepMajorVersions; MinorEnabled=$minorDesired; MinorLimit=$KeepMinorVersions" -ForegroundColor Green
                    }
                    catch {
                        Write-Warning "`tFAILED $($list.Title): $($_.Exception.Message)"
                    }
                }
            }
            else {
                Write-Host "`t$($list.Title) already compliant" -ForegroundColor DarkGray
            }
        }

        # Force deletion of old file version history
        if ($ForceDeleteOldVersions) {
            if ($PSCmdlet.ShouldProcess($SiteUrl, "Delete old file version history")) {
                try {
                    Write-Host "`tStarting batch delete job for old file versions on $SiteUrl..." -ForegroundColor Yellow
                    $batchParams = @{
                        MajorVersionLimit              = $KeepMajorVersions
                        MajorWithMinorVersionsLimit    = $KeepMinorVersions
                    }
                    New-PnPSiteFileVersionBatchDeleteJob @batchParams -Force -ErrorAction Stop
                    Write-Host "`tBatch delete job submitted successfully for $SiteUrl" -ForegroundColor Green
                }
                catch {
                    Write-Warning "`tFAILED to submit batch delete job for ${SiteUrl}: $($_.Exception.Message)"
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
