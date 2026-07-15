<#PSScriptInfo
    .VERSION 3.1.0

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
    Configuration is provided either as an inline JSON string (-InputJson, ideal for
    Azure Automation Runbooks) or as a local JSON file (-ConfigFile, ideal for local
    execution and testing). Both sources share the same schema, parsing and validation.

    .PARAMETER InputJson
    A JSON string containing all configuration. Ideal for Azure Automation Runbooks,
    where the string is pasted directly into the runbook parameter field. Supported
    properties:
      - SiteUrls              (string array, required) — Site Collection URLs to process.
      - KeepMajorVersions     (integer, optional, default: 50) — Number of major versions to keep.
      - KeepMinorVersions     (integer, optional, default: 0) — Number of minor versions to keep.
      - ClientId              (string, optional) — Azure AD App Registration Client ID.
      - ForceDeleteOldVersions (boolean, optional, default: false) — Trigger batch delete of old file versions.
      - DryRun                (boolean, optional, default: false) — Simulate changes without applying them.
      - VersionPolicyMode     (string, optional, default: 'Legacy') — Version policy mechanism.
                              'Legacy' keeps the per-library count-based Set-PnPList behaviour.
                              'AutoExpiration', 'ExpireAfter', 'NoExpiration' and 'InheritFromTenant'
                              drive Set-PnPSiteVersionPolicy at the site level (modern model).
      - ExpireVersionsAfterDays (integer, optional, default: 0) — For 'ExpireAfter' (>= 30);
                              'NoExpiration' forces 0.
      - ApplyTo               (string, optional, default: 'Both') — 'New', 'Existing' or 'Both'
                              document libraries (site version policy modes only).

    .PARAMETER ConfigFile
    Path to a local JSON file containing the same configuration schema as -InputJson.
    Ideal for local execution and testing. The file is read and parsed with
    ConvertFrom-Json. Mutually exclusive with -InputJson. See
    Config/SPSCleanVersions.example.json for a template.

    .EXAMPLE
    .\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/site1"],"KeepMajorVersions":100,"KeepMinorVersions":10}'
    Cleans version history for the specified site, keeping 100 major versions and 10 minor versions.

    .EXAMPLE
    .\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/site1","https://contoso.sharepoint.com/sites/site2"],"KeepMajorVersions":50,"DryRun":true}'
    Simulates the operation on multiple sites without making changes.

    .EXAMPLE
    .\SPSCleanVersions.ps1 -ConfigFile '.\Config\contoso-PROD.json'
    Loads all configuration from a local JSON file. Ideal for local execution and testing.

    .EXAMPLE
    .\SPSCleanVersions.ps1 -InputJson '{"SiteUrls":["https://contoso.sharepoint.com/sites/site1"],"VersionPolicyMode":"ExpireAfter","ExpireVersionsAfterDays":180,"KeepMajorVersions":100}'
    Applies a site-level ExpireAfter version policy (versions expire after 180 days, 100 major versions) via Set-PnPSiteVersionPolicy.

    .NOTES
    FileName:	SPSCleanVersions.ps1
    Author:		Jean-Cyril DROUHIN
    Date:		July 15, 2026
    Version:	3.1.0

    .LINK
    https://spjc.fr/
    https://github.com/luigilink/SPSCleanVersions
#>
#Requires -Version 7.2
#Requires -PSEdition Core
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '2.12.0' }

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'InlineJson')]
param
(
    [Parameter(Mandatory = $true, ParameterSetName = 'InlineJson', HelpMessage = "JSON string containing all configuration (SiteUrls, KeepMajorVersions, KeepMinorVersions, ClientId, ForceDeleteOldVersions, DryRun)")]
    [ValidateNotNullOrEmpty()]
    [System.String]
    $InputJson,

    [Parameter(Mandatory = $true, ParameterSetName = 'ConfigFile', HelpMessage = "Path to a local JSON configuration file (same schema as -InputJson)")]
    [ValidateNotNullOrEmpty()]
    [System.String]
    $ConfigFile
)

#region --- Load and parse JSON input ---
# Configuration comes either from an inline JSON string (-InputJson, Azure Automation
# Runbooks) or from a local JSON file (-ConfigFile, local execution). Both converge on
# the same ConvertFrom-Json parsing and validation below.
if ($PSCmdlet.ParameterSetName -eq 'ConfigFile') {
    if (-not (Test-Path -Path $ConfigFile -PathType Leaf)) {
        throw "Configuration file not found: $ConfigFile"
    }
    $rawJson = Get-Content -Path $ConfigFile -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($rawJson)) {
        throw "Configuration file is empty: $ConfigFile"
    }
}
else {
    $rawJson = $InputJson
}

# Normalize the raw input to catch the most common copy/paste mistakes:
#   - surrounding whitespace,
#   - a wrapping pair of single quotes ('...') copied from a PowerShell command line,
#     which ConvertFrom-Json would otherwise silently parse as a JSON string value.
$rawJson = $rawJson.Trim()
if ($rawJson.Length -ge 2 -and $rawJson.StartsWith("'") -and $rawJson.EndsWith("'")) {
    $rawJson = $rawJson.Substring(1, $rawJson.Length - 2).Trim()
}

try {
    $config = $rawJson | ConvertFrom-Json -ErrorAction Stop
}
catch {
    throw "Invalid JSON input: $($_.Exception.Message). Ensure you pasted raw JSON (an object starting with '{'), " +
    "with straight double quotes and no surrounding single or curly quotes."
}

# ConvertFrom-Json accepts scalars and arrays at the root. A valid configuration must
# be a JSON object; anything else (string, number, array) means the input was wrapped
# in quotes or is otherwise not the expected shape, which would surface later as a
# misleading 'SiteUrls is required' error. Fail early with an actionable message.
if ($config -isnot [System.Management.Automation.PSCustomObject]) {
    throw "InputJson must be a JSON object (starting with '{'), but a $($config.GetType().Name) value was parsed. " +
    "Remove any surrounding single quotes or curly/smart quotes and paste the raw JSON object."
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

# Site version policy (Set-PnPSiteVersionPolicy) properties. VersionPolicyMode selects
# the mechanism: 'Legacy' keeps the per-library Set-PnPList behaviour; the other modes
# drive Set-PnPSiteVersionPolicy at the site level (the modern version-history model).
$validModes = @('Legacy', 'AutoExpiration', 'ExpireAfter', 'NoExpiration', 'InheritFromTenant')
[string]$VersionPolicyMode = if ($config.PSObject.Properties['VersionPolicyMode']) { [string]$config.VersionPolicyMode } else { 'Legacy' }
$matchedMode = $validModes | Where-Object { $_ -ieq $VersionPolicyMode }
if (-not $matchedMode) {
    throw "Invalid 'VersionPolicyMode' value '$VersionPolicyMode'. Allowed values: $($validModes -join ', ')."
}
$VersionPolicyMode = $matchedMode

[int]$ExpireVersionsAfterDays = if ($config.PSObject.Properties['ExpireVersionsAfterDays']) { $config.ExpireVersionsAfterDays } else { 0 }

$validApplyTo = @('New', 'Existing', 'Both')
[string]$ApplyTo = if ($config.PSObject.Properties['ApplyTo']) { [string]$config.ApplyTo } else { 'Both' }
$matchedApplyTo = $validApplyTo | Where-Object { $_ -ieq $ApplyTo }
if (-not $matchedApplyTo) {
    throw "Invalid 'ApplyTo' value '$ApplyTo'. Allowed values: $($validApplyTo -join ', ')."
}
$ApplyTo = $matchedApplyTo

# ExpireVersionsAfterDays must be 0 (NoExpiration) or >= 30 (ExpireAfter), per the
# Set-PnPSiteVersionPolicy contract. ExpireAfter additionally requires a value >= 30.
if ($ExpireVersionsAfterDays -ne 0 -and $ExpireVersionsAfterDays -lt 30) {
    throw "'ExpireVersionsAfterDays' must be 0 (no expiration) or greater than or equal to 30."
}
if ($VersionPolicyMode -eq 'ExpireAfter' -and $ExpireVersionsAfterDays -lt 30) {
    throw "VersionPolicyMode 'ExpireAfter' requires 'ExpireVersionsAfterDays' to be greater than or equal to 30."
}
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

function Test-SiteVersionPolicyDrift {
    <#
        .SYNOPSIS
        Returns $true when the current site version policy differs from the desired
        settings (a drift), $false when they already match.

        .DESCRIPTION
        Reads the current policy via Get-PnPSiteVersionPolicy and compares it to the
        desired mode/values. The comparison uses the fields returned by
        Get-PnPSiteVersionPolicy in PnP.PowerShell (DefaultTrimMode, DefaultExpireAfterDays,
        MajorVersionLimit) and is defensive: if the current policy cannot be read, or a
        field needed for the comparison is missing, the function returns $true (treat as
        drift and apply) rather than silently skipping a real change. Empty/blank fields
        mean no explicit site policy is set (the site inherits the tenant policy).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param
    (
        [Parameter(Mandatory = $true)] [ValidateSet('AutoExpiration', 'ExpireAfter', 'NoExpiration', 'InheritFromTenant')] [string] $Mode,
        [Parameter()] [int] $MajorVersions,
        [Parameter()] [int] $MajorWithMinorVersions,
        [Parameter()] [int] $ExpireAfterDays
    )

    try {
        $current = Get-PnPSiteVersionPolicy -ErrorAction Stop
    }
    catch {
        Write-Verbose "Test-SiteVersionPolicyDrift: unable to read current policy ($($_.Exception.Message)); treating as drift."
        return $true
    }

    if ($null -eq $current) {
        # No policy object returned: treat as inheriting the tenant policy.
        return ($Mode -ne 'InheritFromTenant')
    }

    # Read a property by any of several candidate names, or $null if absent.
    function Get-Prop($obj, [string[]]$names) {
        foreach ($n in $names) {
            $p = $obj.PSObject.Properties[$n]
            if ($null -ne $p) { return $p.Value }
        }
        return $null
    }

    # Get-PnPSiteVersionPolicy returns these as strings; empty/blank means no site policy.
    $curTrimMode = Get-Prop $current @('DefaultTrimMode')
    $curExpire = Get-Prop $current @('DefaultExpireAfterDays', 'ExpireVersionsAfterDays')
    $curMajor = Get-Prop $current @('MajorVersionLimit', 'MajorVersions')

    $hasSitePolicy = -not [string]::IsNullOrWhiteSpace([string]$curTrimMode)

    switch ($Mode) {
        'InheritFromTenant' {
            # Drift only if the site currently has an explicit policy to clear.
            return $hasSitePolicy
        }
        'AutoExpiration' {
            if (-not $hasSitePolicy) { return $true }
            return ("$curTrimMode" -ine 'AutoExpiration')
        }
        default {
            # ExpireAfter / NoExpiration: the trim mode and numeric limits must match.
            if (-not $hasSitePolicy) { return $true }
            if ("$curTrimMode" -ine $Mode) { return $true }
            if ([string]::IsNullOrWhiteSpace([string]$curMajor) -or [int]$curMajor -ne $MajorVersions) { return $true }
            # ExpireAfterDays is only meaningful for ExpireAfter; NoExpiration implies 0.
            $desiredExpire = if ($Mode -eq 'NoExpiration') { 0 } else { $ExpireAfterDays }
            $curExpireInt = if ([string]::IsNullOrWhiteSpace([string]$curExpire)) { 0 } else { [int]$curExpire }
            if ($curExpireInt -ne $desiredExpire) { return $true }
            return $false
        }
    }
}

function Set-SiteVersionPolicy {
    <#
        .SYNOPSIS
        Applies a site-level version policy via Set-PnPSiteVersionPolicy according to the
        requested mode, honouring ShouldProcess/WhatIf.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param
    (
        [Parameter(Mandatory = $true)] [string] $SiteUrl,
        [Parameter(Mandatory = $true)] [ValidateSet('AutoExpiration', 'ExpireAfter', 'NoExpiration', 'InheritFromTenant')] [string] $Mode,
        [Parameter()] [int] $MajorVersions,
        [Parameter()] [int] $MajorWithMinorVersions,
        [Parameter()] [int] $ExpireAfterDays,
        [Parameter()] [ValidateSet('New', 'Existing', 'Both')] [string] $ApplyTo = 'Both'
    )

    # Build the base parameter set for the requested mode.
    $params = @{}
    switch ($Mode) {
        'InheritFromTenant' {
            $params['InheritFromTenant'] = $true
        }
        'AutoExpiration' {
            $params['EnableAutoExpirationVersionTrim'] = $true
        }
        'ExpireAfter' {
            $params['EnableAutoExpirationVersionTrim'] = $false
            $params['ExpireVersionsAfterDays'] = $ExpireAfterDays
            $params['MajorVersions'] = $MajorVersions
        }
        'NoExpiration' {
            $params['EnableAutoExpirationVersionTrim'] = $false
            $params['ExpireVersionsAfterDays'] = 0
            $params['MajorVersions'] = $MajorVersions
        }
    }

    # Target new and/or existing document libraries. InheritFromTenant clears the site
    # setting so new libraries follow the tenant; the existing-libraries request is
    # still valid alongside it.
    $applyNew = ($ApplyTo -eq 'New' -or $ApplyTo -eq 'Both')
    $applyExisting = ($ApplyTo -eq 'Existing' -or $ApplyTo -eq 'Both')
    if ($applyNew) { $params['ApplyToNewDocumentLibraries'] = $true }
    if ($applyExisting) { $params['ApplyToExistingDocumentLibraries'] = $true }

    # MajorWithMinorVersions is only accepted when the request targets existing document
    # libraries. Set-PnPSiteVersionPolicy rejects it for a new-libraries-only request when
    # EnableAutoExpirationVersionTrim is $false. Only add it for ExpireAfter/NoExpiration
    # requests that include existing libraries.
    if ($applyExisting -and $MajorWithMinorVersions -gt 0 -and ($Mode -eq 'ExpireAfter' -or $Mode -eq 'NoExpiration')) {
        $params['MajorWithMinorVersions'] = $MajorWithMinorVersions
    }

    if ($PSCmdlet.ShouldProcess($SiteUrl, "Set site version policy ($Mode, ApplyTo=$ApplyTo)")) {
        Set-PnPSiteVersionPolicy @params -ErrorAction Stop
        Write-Output "`tSite version policy applied: Mode=$Mode; ApplyTo=$ApplyTo"
    }
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

        if ($VersionPolicyMode -eq 'Legacy') {
            # --- Legacy mode: per-library count-based limits via Set-PnPList ---
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
        }
        else {
            # --- Site version policy mode: Set-PnPSiteVersionPolicy at the site level ---
            # Get-/Set-PnPSiteVersionPolicy require a delegated user context that is site
            # collection administrator. In Azure Automation (Managed Identity / app-only)
            # there is no user context, so these calls may fail with an unauthorized error.
            # Warn once so the failure mode is clear.
            if (Test-IsAzureAutomation) {
                Write-Warning @"
Site version policy (Get-/Set-PnPSiteVersionPolicy) requires a delegated user context
that is site collection administrator. Running in Azure Automation (Managed Identity /
app-only) may not be supported and can fail with an unauthorized error for site: $SiteUrl
"@
            }
            Write-Output "Checking site version policy on $SiteUrl (Mode=$VersionPolicyMode)..."
            try {
                $hasDrift = Test-SiteVersionPolicyDrift -Mode $VersionPolicyMode `
                    -MajorVersions $KeepMajorVersions -MajorWithMinorVersions $KeepMinorVersions `
                    -ExpireAfterDays $ExpireVersionsAfterDays
                if ($hasDrift) {
                    Write-Output "`tDrift detected. Applying site version policy..."
                    Set-SiteVersionPolicy -SiteUrl $SiteUrl -Mode $VersionPolicyMode `
                        -MajorVersions $KeepMajorVersions -MajorWithMinorVersions $KeepMinorVersions `
                        -ExpireAfterDays $ExpireVersionsAfterDays -ApplyTo $ApplyTo
                }
                else {
                    Write-Output "`tNo drift. Site version policy already compliant; skipped."
                }
            }
            catch {
                Write-Warning "`tFAILED to apply site version policy on ${SiteUrl}: $($_.Exception.Message)"
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
