<#PSScriptInfo
    .VERSION 1.0.0

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
    Date:		January 27, 2026
    Version:	1.0.0

    .LINK
    https://spjc.fr/
    https://github.com/luigilink/SPSCleanVersions
#>

param
(
    [Parameter(Position = 0, Mandatory=$true, HelpMessage="List of Site Collection URLs separated by commas")]
    [System.String[]]
    $SiteUrls,

    [Parameter(Position = 1, Mandatory=$false)]
    [System.UInt32]
    $KeepMajorVersions = 50,

    [Parameter(Position = 2, Mandatory=$false)]
    [System.UInt32]
    $KeepMinorVersions = 0,

    [Parameter(Position = 3, Mandatory=$false)]
    [System.String]
    $ManagedIdentityId,

    [Parameter(Position = 4, Mandatory=$false)]
    [Switch]
    $WhatIf
)

Process {
    Write-Host "--- Starting SPSCleanVersions ---" -ForegroundColor Cyan

    foreach ($SiteUrl in $SiteUrls) {
        Write-Host "Processing Site: $SiteUrl" -ForegroundColor Yellow
        
        try {
            # Environnement : Local vs Azure Automation
            if ($null -ne $env:AUTOMATION_ASSET_NAME) {
                Write-Output "Running in Azure Automation. Connecting via Managed Identity..."
                if ($null -ne $ManagedIdentityId) {
                    Connect-PnPOnline -Url $SiteUrl -ManagedIdentity -ClientId $ManagedIdentityId
                } else {
                    Connect-PnPOnline -Url $SiteUrl -ManagedIdentity
                }
            } 
            else {
                Write-Output "Running locally. Connecting via Interactive login..."
                Connect-PnPOnline -Url $SiteUrl -Interactive
            }
        }
        catch {
            Write-Error "Failed to process site $SiteUrl : $($_.Exception.Message)"
        }
        finally {
            Disconnect-PnPOnline
        }
    }
}
