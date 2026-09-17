<#PSScriptInfo
    .VERSION 3.1.4

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
      - ClientId              (string, optional in Azure Automation; REQUIRED for local
                              execution) — Azure AD App Registration Client ID. Local runs
                              sign in interactively once via this app and reuse the
                              delegated token (auto-refreshed) across all sites.
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
      - SiteScope             (string, optional, default: 'Selected') — 'Selected' processes
                              SiteUrls; 'All' enumerates every site collection via Get-PnPTenantSite
                              (site version policy modes only; requires TenantAdminUrl).
      - TenantAdminUrl        (string, optional) — SharePoint admin center URL, required when
                              SiteScope is 'All' (e.g. https://contoso-admin.sharepoint.com).
      - SiteFilter            (string, optional) — server-side -Filter passed to Get-PnPTenantSite
                              to narrow the enumeration when SiteScope is 'All'.
      - EnableReport          (boolean, optional, default: true) — write a local HTML report
                              to Results/ (local execution only; not produced in Azure Automation).
      - LogRetentionDays      (integer, optional, default: 180) — prune Logs/ and Results/ files
                              older than this many days (local only). 0 disables pruning.

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
    Date:		September 16, 2026
    Version:	3.1.4

    .LINK
    https://spjc.fr/
    https://github.com/luigilink/SPSCleanVersions
#>
#Requires -Version 7.2
#Requires -PSEdition Core
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '2.12.0' }

# Azure Automation runbooks do not support parameter sets, so -InputJson and -ConfigFile
# are declared as plain optional parameters and their mutual exclusivity is validated in
# the body below (exactly one must be supplied).
[CmdletBinding(SupportsShouldProcess)]
param
(
    [Parameter(HelpMessage = "JSON string containing all configuration (SiteUrls, KeepMajorVersions, KeepMinorVersions, ClientId, ForceDeleteOldVersions, DryRun)")]
    [System.String]
    $InputJson,

    [Parameter(HelpMessage = "Path to a local JSON configuration file (same schema as -InputJson)")]
    [System.String]
    $ConfigFile
)

#region --- Load and parse JSON input ---
# Configuration comes either from an inline JSON string (-InputJson, Azure Automation
# Runbooks) or from a local JSON file (-ConfigFile, local execution). Exactly one must be
# supplied; both converge on the same ConvertFrom-Json parsing and validation below.
$hasInputJson = -not [string]::IsNullOrWhiteSpace($InputJson)
$hasConfigFile = -not [string]::IsNullOrWhiteSpace($ConfigFile)

if (-not $hasInputJson -and -not $hasConfigFile) {
    throw "Provide configuration via -InputJson (inline JSON string) or -ConfigFile (path to a JSON file)."
}
if ($hasInputJson -and $hasConfigFile) {
    throw "-InputJson and -ConfigFile are mutually exclusive; supply only one."
}

if ($hasConfigFile) {
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

# Site scope: 'Selected' processes the explicit SiteUrls; 'All' enumerates every site
# collection in the tenant via Get-PnPTenantSite (requires TenantAdminUrl).
$validScopes = @('Selected', 'All')
[string]$SiteScope = if ($config.PSObject.Properties['SiteScope']) { [string]$config.SiteScope } else { 'Selected' }
$matchedScope = $validScopes | Where-Object { $_ -ieq $SiteScope }
if (-not $matchedScope) {
    throw "Invalid 'SiteScope' value '$SiteScope'. Allowed values: $($validScopes -join ', ')."
}
$SiteScope = $matchedScope

[string]$TenantAdminUrl = if ($config.PSObject.Properties['TenantAdminUrl']) { [string]$config.TenantAdminUrl } else { '' }
[string]$SiteFilter     = if ($config.PSObject.Properties['SiteFilter'])     { [string]$config.SiteFilter }     else { '' }

# SiteUrls is required for 'Selected' scope; for 'All' it is optional (sites are
# enumerated from the tenant) and TenantAdminUrl becomes required instead.
if ($SiteScope -eq 'Selected') {
    if (-not $config.PSObject.Properties['SiteUrls'] -or
        $null -eq $config.SiteUrls -or
        @($config.SiteUrls).Count -eq 0) {
        throw "JSON property 'SiteUrls' is required and must contain at least one URL (or set 'SiteScope' to 'All')."
    }
    [string[]]$SiteUrls = @($config.SiteUrls)
}
else {
    if ([string]::IsNullOrWhiteSpace($TenantAdminUrl)) {
        throw "JSON property 'TenantAdminUrl' is required when 'SiteScope' is 'All' (e.g. https://contoso-admin.sharepoint.com)."
    }
    [string[]]$SiteUrls = @()
}

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

# 'SiteScope: All' only makes sense for the site version policy modes; the Legacy
# per-library path relies on an explicit SiteUrls list.
if ($SiteScope -eq 'All' -and $VersionPolicyMode -eq 'Legacy') {
    throw "'SiteScope' = 'All' is only supported with the site version policy modes (VersionPolicyMode: AutoExpiration, ExpireAfter, NoExpiration or InheritFromTenant), not 'Legacy'."
}

# Reporting / logging properties.
[bool]$EnableReport      = if ($config.PSObject.Properties['EnableReport'])     { $config.EnableReport }     else { $true }
[int]$LogRetentionDays   = if ($config.PSObject.Properties['LogRetentionDays']) { $config.LogRetentionDays } else { 180 }

# Multi-threading (local only). Threads > 1 splits the site list across that many child
# pwsh processes, all sharing a single delegated sign-in via a secured token file. Default
# 1 = the classic sequential behaviour. Azure Automation always runs sequentially.
[int]$Threads = if ($config.PSObject.Properties['Threads']) { [int]$config.Threads } else { 1 }
if ($Threads -lt 1) { $Threads = 1 }
if ($Threads -gt 16) {
    Write-Warning "Threads=$Threads is high and may trigger SharePoint throttling; capping at 16."
    $Threads = 16
}

# Internal worker-mode markers (set by the orchestrator when it spawns child processes;
# never set them by hand). Their presence puts this invocation in worker mode: it skips the
# interactive sign-in, authenticates each site from the shared token file, and writes its
# results as JSON for the parent to merge.
[string]$WorkerTokenFile   = if ($config.PSObject.Properties['_WorkerTokenFile'])   { [string]$config._WorkerTokenFile }   else { '' }
[string]$WorkerResultsFile = if ($config.PSObject.Properties['_WorkerResultsFile']) { [string]$config._WorkerResultsFile } else { '' }
[bool]$IsWorker = -not [string]::IsNullOrWhiteSpace($WorkerTokenFile)
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

# Explicitly import PnP.PowerShell. In Azure Automation runbooks the '#Requires -Modules'
# directive does not import the module, and command auto-loading is unreliable in the
# sandbox, so Connect-PnPOnline would otherwise be 'not recognized'. Harmless locally
# (the module is simply loaded if not already).
try {
    Import-Module -Name PnP.PowerShell -ErrorAction Stop
}
catch {
    throw "Unable to import the PnP.PowerShell module: $($_.Exception.Message). Ensure it is installed (locally) or imported into the Automation Account (Azure Automation)."
}

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

#region --- Throttling / retry helpers ---
# Adapted from luigilink/Philippe Entringer's SPO Storage Assessment toolkit
# (Get-RetryAfterDelay / Test-IsAuthError / Invoke-RetryCommand), MIT-licensed.
# SharePoint Online throttles aggressive callers (HTTP 429/503) with a Retry-After hint;
# at tenant scale (thousands of sites) honouring it is essential to avoid being blocked.

function Get-RetryAfterDelay {
    <#
        .SYNOPSIS
        Extracts a Retry-After delay (in seconds) from a throttling error, or 0 if none.

        .DESCRIPTION
        SharePoint Online throttling responses (HTTP 429/503) carry a Retry-After header.
        Depending on the failure it may surface on the exception's HttpResponseMessage
        (Headers.RetryAfter) or be embedded in the exception message. This checks both and
        returns 0 when no hint is found so the caller can fall back to exponential backoff.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param (
        [Parameter(Mandatory = $true)] [System.Management.Automation.ErrorRecord] $ErrorRecord
    )

    $ex = $ErrorRecord.Exception
    try {
        $response = if ($ex.PSObject.Properties['Response']) { $ex.Response } else { $null }
        $header = if ($response -and $response.PSObject.Properties['Headers']) { $response.Headers } else { $null }
        $ra = if ($header -and $header.PSObject.Properties['RetryAfter']) { $header.RetryAfter } else { $null }
        if ($ra) {
            if ($ra.Delta -and $ra.Delta.TotalSeconds -gt 0) {
                return [int][math]::Ceiling($ra.Delta.TotalSeconds)
            }
            if ($ra.Date) {
                $seconds = ([datetimeoffset]$ra.Date - [datetimeoffset]::UtcNow).TotalSeconds
                if ($seconds -gt 0) { return [int][math]::Ceiling($seconds) }
            }
        }
    }
    catch {
        Write-Verbose "No structured Retry-After header found: $($_.Exception.Message)"
    }

    if ($ex.Message -match 'Retry-After[:\s]+(\d+)') {
        return [int]$Matches[1]
    }
    return 0
}

function Test-IsAuthError {
    <#
        .SYNOPSIS
        Returns $true when an error looks like an authentication / token-expiry failure.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)] [System.Management.Automation.ErrorRecord] $ErrorRecord
    )
    $message = [string]$ErrorRecord.Exception.Message
    return [bool]($message -match '(?i)(\b401\b|unauthorized|invalid.?authentication|token (is )?expired|expired token|access token|AADSTS\d+|InvalidAuthenticationToken)')
}

function Invoke-RetryCommand {
    <#
        .SYNOPSIS
        Runs a script block with Retry-After-aware, exponential-backoff retry on failure.

        .DESCRIPTION
        Executes the supplied script block and, if it throws, retries. When the caught
        exception carries a Retry-After hint (HTTP 429/503 throttling from SharePoint
        Online) that server-provided delay is honoured (capped at 300s); otherwise it falls
        back to exponential backoff (BaseDelaySeconds * 2^attempt). Defence-in-depth on top
        of PnP.PowerShell's own retry, important for tenant-scale runs.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [scriptblock] $ScriptBlock,
        [ValidateRange(0, 20)] [int] $MaxRetries = 5,
        [ValidateRange(1, 300)] [int] $BaseDelaySeconds = 5,
        [string] $OperationName = 'operation'
    )

    $attempt = 0
    do {
        try {
            return & $ScriptBlock
        }
        catch {
            if ($attempt -ge $MaxRetries) { throw }
            $attempt++

            if (Test-IsAuthError -ErrorRecord $_) {
                Write-Warning "[$OperationName] attempt $attempt/$MaxRetries hit an authentication/token error: $($_.Exception.Message). Retrying; if this persists, check the ClientId / app registration."
            }

            $retryAfter = Get-RetryAfterDelay -ErrorRecord $_
            if ($retryAfter -gt 0) {
                $delay = [math]::Min($retryAfter, 300)
                Write-Warning "[$OperationName] attempt $attempt/$MaxRetries throttled: $($_.Exception.Message). Honouring Retry-After: ${delay}s."
            }
            else {
                $delay = [math]::Pow(2, $attempt) * $BaseDelaySeconds
                Write-Warning "[$OperationName] attempt $attempt/$MaxRetries failed: $($_.Exception.Message). Retrying in ${delay}s."
            }
            Start-Sleep -Seconds $delay
        }
    } while ($attempt -le $MaxRetries)
}
#endregion

#region --- Multi-thread orchestration (local only) ---
# For large local runs the site list can be processed by several child pwsh processes in
# parallel. All children share a SINGLE delegated sign-in: the parent signs in once, writes
# the access token to a permission-restricted temp file, and each child reads it to connect
# with -AccessToken (no extra prompts). The parent refreshes the token file while the
# children run so long batches never hit expiry. Multi-threading is LOCAL ONLY — Azure
# Automation always runs sequentially.

function Split-SitesIntoSlices {
    <#
        .SYNOPSIS
        Splits a list of site URLs into (at most) $Count contiguous slices, as evenly as
        possible. Empty slices are never returned.
    #>
    [CmdletBinding()]
    [OutputType([System.Collections.IEnumerable])]
    param
    (
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [string[]] $Sites,
        [Parameter(Mandatory = $true)] [ValidateRange(1, 64)] [int] $Count
    )
    $total = $Sites.Count
    if ($total -eq 0) { return @() }
    $n = [math]::Min($Count, $total)
    $base = [math]::Floor($total / $n)
    $remainder = $total % $n
    $slices = New-Object System.Collections.Generic.List[object]
    $index = 0
    for ($i = 0; $i -lt $n; $i++) {
        # Distribute the remainder one extra item at a time across the first slices.
        $size = $base + $(if ($i -lt $remainder) { 1 } else { 0 })
        $slice = [string[]]($Sites[$index..($index + $size - 1)])
        $slices.Add($slice)
        $index += $size
    }
    return , $slices.ToArray()
}

function Save-DelegatedTokenFile {
    <#
        .SYNOPSIS
        Writes the delegated access token to a user-only file. Hardens the permissions on an
        empty file FIRST, verifies them, and only then writes the token — so the secret is
        never briefly exposed with default permissions. Fails closed (throws, no token
        written) if the file cannot be locked down. Atomic move so readers never see a
        partial file.
    #>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)] [string] $Token,
        [Parameter(Mandatory = $true)] [string] $Path
    )
    $tmp = "$Path.tmp"
    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -WhatIf:$false }
    # 1. Create an EMPTY file, then restrict it before any secret touches disk.
    $null = New-Item -Path $tmp -ItemType File -Force -WhatIf:$false
    $hardened = $false
    try {
        if ($IsWindows) {
            $acl = Get-Acl -Path $tmp
            $acl.SetAccessRuleProtection($true, $false)
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                [System.Security.Principal.WindowsIdentity]::GetCurrent().Name,
                'FullControl', 'Allow')
            $acl.AddAccessRule($rule)
            Set-Acl -Path $tmp -AclObject $acl
            # Verify no inherited/extra identities remain beyond the current user.
            $current = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $others = @((Get-Acl -Path $tmp).Access | Where-Object { $_.IdentityReference.Value -ne $current })
            $hardened = ($others.Count -eq 0)
        }
        else {
            & chmod 600 $tmp 2>$null
            # Verify the mode really is user-only (no group/other bits).
            $mode = (Get-Item -LiteralPath $tmp).UnixMode
            $hardened = ($mode -match '^.rw-------')
        }
    }
    catch {
        $hardened = $false
    }
    if (-not $hardened) {
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue -WhatIf:$false
        throw "Refusing to write the shared token: could not restrict permissions on $tmp to the current user only."
    }
    # 2. Now write the token into the already-locked-down file, then publish atomically.
    Set-Content -Path $tmp -Value $Token -Encoding UTF8 -NoNewline -Force -WhatIf:$false
    Move-Item -Path $tmp -Destination $Path -Force -WhatIf:$false
}

function Get-DelegatedTokenFromFile {
    <#
        .SYNOPSIS
        Reads the delegated access token from the shared token file, tolerating a transient
        read while the parent atomically refreshes it.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param
    (
        [Parameter(Mandatory = $true)] [string] $Path
    )
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            $value = (Get-Content -Path $Path -Raw -ErrorAction Stop).Trim()
            if (-not [string]::IsNullOrWhiteSpace($value)) { return $value }
        }
        catch {
            Start-Sleep -Milliseconds 200
        }
    }
    throw "Unable to read the shared delegated token file: $Path"
}

function New-WorkerConfig {
    <#
        .SYNOPSIS
        Builds a per-thread configuration object from the parent config: the original
        settings, the thread's slice of sites, and the internal worker markers. Threads is
        forced to 1 so a worker never orchestrates recursively.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param
    (
        [Parameter(Mandatory = $true)] $BaseConfig,
        [Parameter(Mandatory = $true)] [AllowEmptyCollection()] [string[]] $Slice,
        [Parameter(Mandatory = $true)] [string] $TokenFile,
        [Parameter(Mandatory = $true)] [string] $ResultsFile
    )
    $worker = @{}
    foreach ($p in $BaseConfig.PSObject.Properties) { $worker[$p.Name] = $p.Value }
    # A worker always processes an explicit list of sites, never re-enumerates the tenant.
    $worker.Remove('SiteScope') | Out-Null
    $worker['SiteUrls'] = $Slice
    $worker['Threads'] = 1
    $worker['EnableReport'] = $false
    $worker['_WorkerTokenFile'] = $TokenFile
    $worker['_WorkerResultsFile'] = $ResultsFile
    return $worker
}
#endregion

#region --- Reporting helpers ---
# Per-site result records collected during the run and rendered into the report.
$script:RunResults = New-Object System.Collections.Generic.List[object]

function Add-RunResult {
    param(
        [Parameter(Mandatory = $true)] [string] $SiteUrl,
        [Parameter(Mandatory = $true)] [string] $Scope,
        [Parameter(Mandatory = $true)] [string] $Outcome,
        [Parameter()] [string] $Detail = ''
    )
    $script:RunResults.Add([PSCustomObject][ordered]@{
            Site    = $SiteUrl
            Scope   = $Scope
            Outcome = $Outcome
            Detail  = $Detail
        })
}

function ConvertTo-SPSHtmlEncoded {
    # HTML-encodes a value for safe insertion into the generated report.
    param([Parameter(ValueFromPipeline = $true)][AllowNull()][AllowEmptyString()][string] $Value)
    process {
        if ([string]::IsNullOrEmpty($Value)) { return '' }
        return [System.Net.WebUtility]::HtmlEncode($Value)
    }
}

function Export-SPSCleanVersionsReport {
    <#
        .SYNOPSIS
        Builds a self-contained (no CDN) HTML report from the collected run results and
        returns it as a string. Summary cards plus a filterable table.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param
    (
        [Parameter(Mandatory = $true)] [System.Collections.IEnumerable] $Results,
        [Parameter()] [string] $Title = 'SPSCleanVersions',
        [Parameter()] [string] $Version = '',
        [Parameter()] [bool] $DryRunMode = $false
    )

    $rows = @($Results)
    $total = $rows.Count
    $applied = @($rows | Where-Object { $_.Outcome -eq 'Applied' }).Count
    $wouldApply = @($rows | Where-Object { $_.Outcome -eq 'WouldApply' }).Count
    $skipped = @($rows | Where-Object { $_.Outcome -eq 'Skipped' -or $_.Outcome -eq 'Compliant' }).Count
    $failed = @($rows | Where-Object { $_.Outcome -eq 'Failed' }).Count
    $appliedLabel = if ($DryRunMode) { 'Would apply' } else { 'Applied' }
    $appliedValue = if ($DryRunMode) { $wouldApply } else { $applied }
    $generated = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $overall = if ($failed -gt 0) { 'ATTENTION' } else { 'OK' }
    $overallClass = if ($failed -gt 0) { 'kpi-alert' } else { 'kpi-ok' }
    $dryTag = if ($DryRunMode) { '<span class="kpi kpi-dry">DryRun</span>' } else { '' }

    $css = @'
:root{--ink:#1b1b1b;--muted:#6b7280;--bg:#f4f6f8;--card:#ffffff;--line:#e5e7eb;--brand:#2b5797;--brand-dark:#1e3f6f;--ok-bg:#bfff80;--ok-ink:#13300a;--alert-bg:#ff6464;--alert-ink:#3a0000;--dry-bg:#ffd54a;--dry-ink:#3a2e00;--zebra:#f7f9fb}
*{box-sizing:border-box}
body{margin:0;padding:0;background:var(--bg);color:var(--ink);font:14px/1.45 'Segoe UI','Aptos',Arial,sans-serif}
header.banner{position:sticky;top:0;z-index:10;padding:12px 20px;color:#fff;background:var(--brand);display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px}
header.banner h1{margin:0;font-size:16px;font-weight:600;display:flex;align-items:center;gap:8px}
.kpi{display:inline-block;padding:4px 12px;border-radius:6px;font-weight:700;font-size:12px;margin-left:6px}
.kpi-ok{background:var(--ok-bg);color:var(--ok-ink)}
.kpi-alert{background:var(--alert-bg);color:var(--alert-ink)}
.kpi-dry{background:var(--dry-bg);color:var(--dry-ink)}
.banner .meta{color:#e5e7eb;font-size:12px}
.layout{max-width:1100px;margin:16px auto;padding:0 16px}
.cards{display:flex;flex-wrap:wrap;gap:12px;margin:0 0 16px 0}
.card{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:14px 18px;min-width:140px;flex:1}
.card.accent{border-color:#f2b8b8;background:#fff5f5}
.card-value{font-size:26px;font-weight:700;color:var(--brand)}
.card.accent .card-value{color:#c0392b}
.card-label{font-size:12px;color:var(--muted);margin-top:2px}
section{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:12px 16px}
section h2{margin:0 0 10px 0;font-size:14px;color:var(--brand)}
.search{width:100%;padding:8px 10px;border:1px solid var(--line);border-radius:6px;font-size:13px;margin:0 0 12px 0}
.table-wrap{overflow:auto;max-height:70vh;border:1px solid var(--line);border-radius:6px}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{text-align:left;padding:7px 10px;border-bottom:1px solid var(--line);vertical-align:top}
thead th{position:sticky;top:0;background:#eef2f7;color:#10222e;font-weight:600;z-index:1}
tbody tr:nth-child(even){background:var(--zebra)}
tr.row-alert td{background:#fff5f5}
.badge{display:inline-block;padding:2px 10px;border-radius:999px;font-size:11px;font-weight:600;color:#fff}
.badge.Applied{background:var(--brand)}
.badge.WouldApply{background:#6f42c1}
.badge.Skipped,.badge.Compliant{background:#9aa4ad}
.badge.Failed{background:#c0392b}
footer{color:var(--muted);font-size:12px;text-align:center;padding:16px 0}
'@

    $sb = New-Object System.Text.StringBuilder
    foreach ($r in $rows) {
        $oc = ConvertTo-SPSHtmlEncoded ([string]$r.Outcome)
        $rowClass = if ($r.Outcome -eq 'Failed') { ' class="row-alert"' } else { '' }
        [void]$sb.Append("<tr$rowClass><td>" + (ConvertTo-SPSHtmlEncoded ([string]$r.Site)) + '</td>')
        [void]$sb.Append('<td>' + (ConvertTo-SPSHtmlEncoded ([string]$r.Scope)) + '</td>')
        [void]$sb.Append('<td><span class="badge ' + $oc + '">' + $oc + '</span></td>')
        [void]$sb.Append('<td>' + (ConvertTo-SPSHtmlEncoded ([string]$r.Detail)) + '</td></tr>')
    }

    $failedCardClass = if ($failed -gt 0) { 'card accent' } else { 'card' }
    $encTitle = ConvertTo-SPSHtmlEncoded $Title
    $encVer = ConvertTo-SPSHtmlEncoded $Version
    $page = @"
<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>$encTitle</title><style>$css</style></head><body>
<header class="banner">
  <h1>$encTitle <span class="kpi $overallClass">$overall</span> $dryTag</h1>
  <span class="meta">generated $generated &middot; v$encVer</span>
</header>
<div class="layout">
  <div class="cards">
    <div class="card"><div class="card-value">$total</div><div class="card-label">Sites processed</div></div>
    <div class="card"><div class="card-value">$appliedValue</div><div class="card-label">$appliedLabel</div></div>
    <div class="card"><div class="card-value">$skipped</div><div class="card-label">Skipped / compliant</div></div>
    <div class="$failedCardClass"><div class="card-value">$failed</div><div class="card-label">Failed</div></div>
  </div>
  <section>
    <h2>Per-site results</h2>
    <input id="spsSearch" class="search" type="search" placeholder="Filter rows...">
    <div class="table-wrap">
      <table><thead><tr><th>Site</th><th>Scope</th><th>Outcome</th><th>Detail</th></tr></thead><tbody id="spsBody">
$($sb.ToString())
      </tbody></table>
    </div>
  </section>
  <footer>Generated by SPSCleanVersions v$encVer &middot; $generated</footer>
</div>
<script>
(function(){var q=document.getElementById('spsSearch');q.addEventListener('input',function(){var t=q.value.toLowerCase();document.querySelectorAll('#spsBody tr').forEach(function(tr){tr.style.display=(t===''||tr.textContent.toLowerCase().indexOf(t)>-1)?'':'none';});});})();
</script>
</body></html>
"@
    return $page
}

function Clear-OldRunFiles {
    # Prune Logs/Results files older than the retention window (local only).
    param([string] $Path, [int] $Retention, [string] $Filter)
    if ($Retention -le 0 -or -not (Test-Path $Path)) { return }
    $cutoff = (Get-Date).AddDays(-$Retention)
    Get-ChildItem -Path $Path -Filter $Filter -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -le $cutoff } |
        ForEach-Object { Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue -WhatIf:$false }
}
#endregion

# Run context: local writes transcript + report files; Azure Automation emits the report
# into the output stream (no persistent filesystem).
$script:IsAzureAutomationRun = Test-IsAzureAutomation
$script:ScriptVersion = '3.1.4'
$script:RunTimestamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'
$script:LogsFolder = $null
$script:ResultsFolder = $null
$script:TranscriptStarted = $false

if (-not $script:IsAzureAutomationRun) {
    $scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $script:LogsFolder = Join-Path -Path $scriptRoot -ChildPath 'Logs'
    $script:ResultsFolder = Join-Path -Path $scriptRoot -ChildPath 'Results'
    foreach ($dir in @($script:LogsFolder, $script:ResultsFolder)) {
        if (-not (Test-Path -Path $dir)) { $null = New-Item -Path $dir -ItemType Directory -Force -WhatIf:$false }
    }
    Clear-OldRunFiles -Path $script:LogsFolder -Retention $LogRetentionDays -Filter '*.log'
    Clear-OldRunFiles -Path $script:ResultsFolder -Retention $LogRetentionDays -Filter '*.html'
    try {
        $transcriptPath = Join-Path -Path $script:LogsFolder -ChildPath ("SPSCleanVersions-$($script:RunTimestamp).log")
        Start-Transcript -Path $transcriptPath -IncludeInvocationHeader -WhatIf:$false | Out-Null
        $script:TranscriptStarted = $true
    }
    catch {
        Write-Warning "Unable to start transcript: $($_.Exception.Message)"
    }
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
        $current = Invoke-RetryCommand -OperationName 'Get-PnPSiteVersionPolicy' -ScriptBlock { Get-PnPSiteVersionPolicy -ErrorAction Stop }
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

    # MajorWithMinorVersions handling for ExpireAfter/NoExpiration (EnableAutoExpirationVersionTrim = $false):
    #   - For requests that include existing document libraries, SharePoint REQUIRES all three of
    #     ExpireVersionsAfterDays, MajorVersions and MajorWithMinorVersions to be specified — even
    #     when MajorWithMinorVersions is 0. Omitting it fails with "You must specify
    #     ExpireVersionsAfterDays, MajorVersions and MajorWithMinorVersions ... for document
    #     libraries that including existing ones."
    #   - It is rejected for a new-libraries-only request, so only add it when existing libraries
    #     are targeted.
    if ($applyExisting -and ($Mode -eq 'ExpireAfter' -or $Mode -eq 'NoExpiration')) {
        $params['MajorWithMinorVersions'] = $MajorWithMinorVersions
    }

    if ($PSCmdlet.ShouldProcess($SiteUrl, "Set site version policy ($Mode, ApplyTo=$ApplyTo)")) {
        Invoke-RetryCommand -OperationName "Set-PnPSiteVersionPolicy ($Mode)" -ScriptBlock { Set-PnPSiteVersionPolicy @params -ErrorAction Stop }
        Write-Output "`tSite version policy applied: Mode=$Mode; ApplyTo=$ApplyTo"
    }
}

function Resolve-EffectiveApplyTo {
    <#
        .SYNOPSIS
        Resolves the ApplyTo target actually usable in the current authentication context.

        .DESCRIPTION
        Applying a site version policy to EXISTING document libraries is not supported with
        app-only authentication (Azure Automation / Managed Identity): SharePoint answers
        "Cannot call this API with an app-only principal." In that context this function
        downgrades the target so the run does the app-only-capable work instead of failing:
          - 'Both'     -> 'New'  (keep the site default that governs new libraries)
          - 'Existing' -> 'None' (nothing can be applied app-only)
        'New' is unchanged, and local/delegated runs (IsAzureAutomation = $false) are never
        downgraded.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param
    (
        [Parameter(Mandatory = $true)] [ValidateSet('New', 'Existing', 'Both')] [string] $ApplyTo,
        [Parameter(Mandatory = $true)] [bool] $IsAzureAutomation
    )
    if ($IsAzureAutomation -and ($ApplyTo -eq 'Existing' -or $ApplyTo -eq 'Both')) {
        return $(if ($ApplyTo -eq 'Both') { 'New' } else { 'None' })
    }
    return $ApplyTo
}

function Get-TenantSiteUrls {
    <#
        .SYNOPSIS
        Connects to the tenant admin center and returns the URLs of all site collections
        (OneDrive excluded), optionally narrowed by a server-side filter.
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param
    (
        [Parameter(Mandatory = $true)] [string] $AdminUrl,
        [Parameter()] [string] $Filter = '',
        [Parameter()] [string] $ClientId = ''
    )

    # NOTE: this function returns the URL array, so it must not emit anything else to the
    # success stream — any Write-Output here would be captured into the returned value and
    # then processed as bogus 'sites'. Informational messages use Write-Verbose; the caller
    # logs the discovered count.
    Write-Verbose "Connecting to tenant admin center: $AdminUrl ..."
    if (Test-IsAzureAutomation) {
        if (-not [string]::IsNullOrEmpty($ClientId)) {
            Connect-PnPOnline -Url $AdminUrl -ManagedIdentity -ClientId $ClientId
        }
        else {
            Connect-PnPOnline -Url $AdminUrl -ManagedIdentity
        }
    }
    else {
        Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $ClientId
    }

    try {
        $getParams = @{ ErrorAction = 'Stop' }
        if (-not [string]::IsNullOrWhiteSpace($Filter)) { $getParams['Filter'] = $Filter }
        $sites = Invoke-RetryCommand -OperationName 'Get-PnPTenantSite' -ScriptBlock { Get-PnPTenantSite @getParams }
        $urls = @($sites | Where-Object { $null -ne $_.Url } | Select-Object -ExpandProperty Url)
        Write-Verbose "Discovered $($urls.Count) site collection(s) from the tenant."
        return , [string[]]$urls
    }
    finally {
        Disconnect-PnPOnline
    }
}

# Resolve the list of sites to process. For 'All' scope, enumerate the tenant first.
if ($SiteScope -eq 'All') {
    Write-Output "--- SiteScope=All: enumerating tenant site collections ---"
    Write-Output "Connecting to tenant admin center: $TenantAdminUrl ..."
    if ($WhatIfPreference) {
        Write-Warning "SiteScope=All applies the version policy across the whole tenant. Review the DryRun output carefully before a real run."
    }
    try {
        $SiteUrls = Get-TenantSiteUrls -AdminUrl $TenantAdminUrl -Filter $SiteFilter -ClientId $ClientId
    }
    catch {
        throw "Failed to enumerate tenant sites from ${TenantAdminUrl}: $($_.Exception.Message)"
    }
    Write-Output "Discovered $(@($SiteUrls).Count) site collection(s) to process."
    if (@($SiteUrls).Count -eq 0) {
        Write-Warning "No site collections were returned from the tenant; nothing to process."
    }
}

# --- Local batch auth: sign in ONCE and reuse a delegated token across all sites ---
# Interactive sign-in per site does not scale for batches: each Connect-PnPOnline -Interactive
# can re-prompt for a browser login, which is unusable for hundreds/thousands of sites. For
# local execution we therefore sign in ONCE (interactive) to an anchor site to establish a
# delegated MSAL context, then reuse its access token for every site. The token is a
# SharePoint token, which is valid tenant-wide, and it is re-read from the interactive
# connection on each iteration so MSAL refreshes it silently before it expires — no repeated
# browser prompts over a long run. If the single sign-in fails we fall back to the previous
# per-site interactive behaviour.
$script:DelegatedAuthConnection = $null
if (-not $script:IsAzureAutomationRun -and -not $IsWorker -and @($SiteUrls).Count -gt 0) {
    if ([string]::IsNullOrWhiteSpace($ClientId)) {
        throw "ClientId is required for local/interactive execution. Register an app once with 'Register-PnPEntraIDAppForInteractiveLogin' and pass its Client ID as the 'ClientId' config property."
    }
    # Anchor on a regular site (valid tenant-wide token); fall back to the admin URL if needed.
    $anchorUrl = if (@($SiteUrls).Count -gt 0) { @($SiteUrls)[0] }
    elseif (-not [string]::IsNullOrWhiteSpace($TenantAdminUrl)) { $TenantAdminUrl }
    else { $null }
    if ($null -ne $anchorUrl) {
        try {
            Write-Output "Signing in once (interactive) for the whole batch via: $anchorUrl ..."
            $script:DelegatedAuthConnection = Connect-PnPOnline -Url $anchorUrl -Interactive -ClientId $ClientId -ReturnConnection
            Write-Output "Interactive sign-in complete. The delegated token will be reused (auto-refreshed) for every site; no further prompts expected."
        }
        catch {
            Write-Warning "Single batch sign-in failed ($($_.Exception.Message)). Falling back to interactive login per site."
            $script:DelegatedAuthConnection = $null
        }
    }
}

# --- Multi-thread orchestration (local only) ---
# When Threads > 1 for a local run with more than one site, split the work across child pwsh
# processes that share this single sign-in via a secured token file, then merge their results.
$script:RunAsOrchestrator = $false
if ($Threads -gt 1 -and -not $script:IsAzureAutomationRun -and -not $IsWorker -and @($SiteUrls).Count -gt 1) {
    if ($null -eq $script:DelegatedAuthConnection) {
        Write-Warning "Multi-threading requires the single interactive sign-in, which is not available; falling back to sequential processing."
    }
    else {
        $script:RunAsOrchestrator = $true
        $runRoot = Join-Path -Path $script:ResultsFolder -ChildPath "multithread-$($script:RunTimestamp)"
        $tokenFile = Join-Path -Path $runRoot -ChildPath 'token.dat'
        $null = New-Item -Path $runRoot -ItemType Directory -Force -WhatIf:$false
        $selfPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
        $workers = $null

        try {
            $slices = Split-SitesIntoSlices -Sites ([string[]]@($SiteUrls)) -Count $Threads
            Save-DelegatedTokenFile -Token (Get-PnPAccessToken -Connection $script:DelegatedAuthConnection) -Path $tokenFile
            Write-Output "Multi-thread: $(@($SiteUrls).Count) site(s) across $($slices.Count) worker process(es)."

            # Track each worker with the slice it owns and the file it must produce, so we can
            # validate coverage and exit codes after the run.
            $workers = New-Object System.Collections.Generic.List[object]
            for ($i = 0; $i -lt $slices.Count; $i++) {
                $threadNo = $i + 1
                $threadFolder = Join-Path -Path $runRoot -ChildPath "Thread$threadNo"
                $null = New-Item -Path $threadFolder -ItemType Directory -Force -WhatIf:$false
                $resultsFile = Join-Path -Path $threadFolder -ChildPath 'results.json'
                $threadConfigPath = Join-Path -Path $threadFolder -ChildPath 'config.json'
                $workerCfg = New-WorkerConfig -BaseConfig $config -Slice ([string[]]$slices[$i]) -TokenFile $tokenFile -ResultsFile $resultsFile
                ($workerCfg | ConvertTo-Json -Depth 10) | Set-Content -Path $threadConfigPath -Encoding UTF8 -Force -WhatIf:$false
                # Quote the path arguments: Start-Process joins -ArgumentList with spaces and
                # does not preserve boundaries, so an install/config path containing spaces
                # would otherwise split and the worker would never load the script/config.
                $proc = Start-Process -FilePath 'pwsh' -PassThru -WindowStyle Hidden -ArgumentList @(
                    '-NoProfile', '-ExecutionPolicy', 'Bypass',
                    '-File', ('"{0}"' -f $selfPath),
                    '-ConfigFile', ('"{0}"' -f $threadConfigPath)
                )
                $workers.Add([PSCustomObject]@{ ThreadNo = $threadNo; Process = $proc; Slice = [string[]]$slices[$i]; ResultsFile = $resultsFile })
                Write-Output "  Worker $threadNo started (PID $($proc.Id)) for $(@($slices[$i]).Count) site(s)."
            }

            # Wait for all workers; refresh the shared token file periodically so long runs
            # never hit token expiry (the interactive connection refreshes it silently).
            $lastRefresh = Get-Date
            while (@($workers | Where-Object { -not $_.Process.HasExited }).Count -gt 0) {
                Start-Sleep -Seconds 5
                if (((Get-Date) - $lastRefresh).TotalMinutes -ge 20) {
                    try {
                        Save-DelegatedTokenFile -Token (Get-PnPAccessToken -Connection $script:DelegatedAuthConnection) -Path $tokenFile
                        $lastRefresh = Get-Date
                    }
                    catch {
                        Write-Warning "Token refresh failed: $($_.Exception.Message)"
                    }
                }
            }

            # Merge each worker's results, then validate exit codes and full slice coverage so
            # a crashed worker cannot silently drop its whole slice while the report shows OK.
            foreach ($w in $workers) {
                $reportedSites = New-Object System.Collections.Generic.HashSet[string]
                if (Test-Path -Path $w.ResultsFile) {
                    try {
                        $rows = Get-Content -Path $w.ResultsFile -Raw | ConvertFrom-Json
                        foreach ($row in @($rows)) {
                            Add-RunResult -SiteUrl ([string]$row.Site) -Scope ([string]$row.Scope) -Outcome ([string]$row.Outcome) -Detail ([string]$row.Detail)
                            $null = $reportedSites.Add([string]$row.Site)
                        }
                    }
                    catch {
                        Write-Warning "Could not read worker $($w.ThreadNo) results '$($w.ResultsFile)': $($_.Exception.Message)"
                    }
                }
                else {
                    Write-Warning "Worker $($w.ThreadNo) produced no results file (exit code $($w.Process.ExitCode))."
                }

                $exitCode = $w.Process.ExitCode
                if ($exitCode -ne 0) {
                    Write-Warning "Worker $($w.ThreadNo) exited with code $exitCode."
                }
                # Any assigned site that did not produce a row is recorded as a failure so the
                # consolidated report reflects the incomplete coverage instead of hiding it.
                foreach ($site in $w.Slice) {
                    if (-not $reportedSites.Contains($site)) {
                        Add-RunResult -SiteUrl $site -Scope "Thread$($w.ThreadNo)" -Outcome 'Failed' `
                            -Detail "No result reported by worker $($w.ThreadNo) (exit code $exitCode); the site may not have been processed."
                    }
                }
            }
        }
        finally {
            # On an exceptional path (e.g. a later Start-Process threw), stop any workers still
            # running before removing the shared token, so no orphaned worker keeps changing
            # sites after the parent has aborted.
            if ($null -ne $workers) {
                foreach ($w in $workers) {
                    try {
                        if ($w.Process -and -not $w.Process.HasExited) {
                            $w.Process.Kill()
                            $w.Process.WaitForExit(10000) | Out-Null
                        }
                    }
                    catch {
                        Write-Verbose "Could not stop worker $($w.ThreadNo): $($_.Exception.Message)"
                    }
                }
            }
            if (Test-Path -Path $tokenFile) { Remove-Item -Path $tokenFile -Force -ErrorAction SilentlyContinue -WhatIf:$false }
        }
    }
}

foreach ($SiteUrl in $(if ($script:RunAsOrchestrator) { @() } else { $SiteUrls })) {
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
        elseif ($IsWorker) {
            # Worker process: authenticate each site from the shared token file (re-read every
            # site so parent token refreshes are picked up). No interactive prompt.
            $accessToken = Get-DelegatedTokenFromFile -Path $WorkerTokenFile
            Connect-PnPOnline -Url $SiteUrl -AccessToken $accessToken
        }
        else {
            if ($null -ne $script:DelegatedAuthConnection) {
                # Reuse the single batch sign-in: read a fresh (silently MSAL-refreshed) token
                # from the interactive connection and connect to this site with it — no prompt.
                $accessToken = Get-PnPAccessToken -Connection $script:DelegatedAuthConnection
                Connect-PnPOnline -Url $SiteUrl -AccessToken $accessToken
            }
            else {
                Write-Output "Running locally. Connecting via Interactive login..."
                Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
            }
        }

        if ($VersionPolicyMode -eq 'Legacy') {
            # --- Legacy mode: per-library count-based limits via Set-PnPList ---
            # Get all Lists in the Site
            Write-Output "Retrieving lists from $SiteUrl..."
            $allLists = Invoke-RetryCommand -OperationName 'Get-PnPList' -ScriptBlock { Get-PnPList -ErrorAction Stop }
            $targetLists = $allLists | Where-Object {
                $_.Hidden -eq $false -and
                $_.EnableVersioning -eq $true -and
                $_.RootFolder.ServerRelativeUrl -notlike "*_catalogs*" -and
                $_.RootFolder.ServerRelativeUrl -notlike "*/SiteAssets*" -and
                $_.RootFolder.ServerRelativeUrl -notlike "*/SitePages*" -and
                $_.RootFolder.ServerRelativeUrl -notlike "*/Style Library*" -and
                $_.BaseTemplate -eq 101 # Document Libraries only
            }
            $legacyApplied = 0; $legacyCompliant = 0; $legacyFailed = 0
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
                            Invoke-RetryCommand -OperationName "Set-PnPList ($($list.Title))" -ScriptBlock { Set-PnPList @p -ErrorAction Stop }
                            Write-Output "`t$($list.Title) -> Major=$KeepMajorVersions; MinorEnabled=$minorDesired; MinorLimit=$KeepMinorVersions"
                            $legacyApplied++
                        }
                        catch {
                            Write-Warning "`tFAILED $($list.Title): $($_.Exception.Message)"
                            $legacyFailed++
                        }
                    }
                }
                else {
                    Write-Output "`t$($list.Title) already compliant"
                    $legacyCompliant++
                }
            }
            $legacyOutcome = if ($legacyFailed -gt 0) { 'Failed' } elseif ($legacyApplied -gt 0) { 'Applied' } else { 'Compliant' }
            Add-RunResult -SiteUrl $SiteUrl -Scope "Legacy (Major=$KeepMajorVersions,Minor=$KeepMinorVersions)" -Outcome $legacyOutcome `
                -Detail "$legacyApplied applied, $legacyCompliant compliant, $legacyFailed failed across $(@($targetLists).Count) libraries"
        }
        else {
            # --- Site version policy mode: Set-PnPSiteVersionPolicy at the site level ---
            # Get-/Set-PnPSiteVersionPolicy read and the site default / NEW-libraries write
            # both work with app-only (Managed Identity) in Azure Automation. However,
            # applying the policy to EXISTING document libraries is NOT supported app-only:
            # SharePoint answers "Cannot call this API with an app-only principal." So in
            # Azure Automation we drop the existing-libraries target (App-only can still set
            # the site default that governs new libraries) and tell the user to run the
            # existing-libraries pass locally / interactively with a SharePoint Administrator.
            $effectiveApplyTo = Resolve-EffectiveApplyTo -ApplyTo $ApplyTo -IsAzureAutomation ([bool](Test-IsAzureAutomation))
            if ($effectiveApplyTo -ne $ApplyTo) {
                Write-Warning @"
App-only (Managed Identity) cannot apply the version policy to EXISTING document libraries
("Cannot call this API with an app-only principal"). Existing libraries are skipped in
Azure Automation. Run VersionPolicyMode '$VersionPolicyMode' (ApplyTo=Existing) locally /
interactively with a SharePoint Administrator to cover existing libraries for: $SiteUrl
"@
            }

            Write-Output "Checking site version policy on $SiteUrl (Mode=$VersionPolicyMode)..."
            if ($effectiveApplyTo -eq 'None') {
                Write-Output "`tApp-only cannot target existing libraries; nothing to apply here. Skipped."
                Add-RunResult -SiteUrl $SiteUrl -Scope "$VersionPolicyMode (ApplyTo=$ApplyTo)" -Outcome 'Skipped' `
                    -Detail 'App-only: existing document libraries require a delegated context; run locally/interactively.'
            }
            else {
                # Note appended to output/results when the existing-libraries target was dropped for app-only.
                $existingNote = if ($effectiveApplyTo -ne $ApplyTo) { ' (existing libraries skipped: app-only)' } else { '' }
                try {
                    $hasDrift = Test-SiteVersionPolicyDrift -Mode $VersionPolicyMode `
                        -MajorVersions $KeepMajorVersions -MajorWithMinorVersions $KeepMinorVersions `
                        -ExpireAfterDays $ExpireVersionsAfterDays
                    if ($hasDrift) {
                        if ($WhatIfPreference) {
                            Write-Output "`tDrift detected. Would apply site version policy (DryRun; no change made).$existingNote"
                            Set-SiteVersionPolicy -SiteUrl $SiteUrl -Mode $VersionPolicyMode `
                                -MajorVersions $KeepMajorVersions -MajorWithMinorVersions $KeepMinorVersions `
                                -ExpireAfterDays $ExpireVersionsAfterDays -ApplyTo $effectiveApplyTo
                            Add-RunResult -SiteUrl $SiteUrl -Scope "$VersionPolicyMode (ApplyTo=$effectiveApplyTo)" -Outcome 'WouldApply' `
                                -Detail "DryRun: would set Major=$KeepMajorVersions; ExpireAfterDays=$ExpireVersionsAfterDays$existingNote"
                        }
                        else {
                            Write-Output "`tDrift detected. Applying site version policy...$existingNote"
                            Set-SiteVersionPolicy -SiteUrl $SiteUrl -Mode $VersionPolicyMode `
                                -MajorVersions $KeepMajorVersions -MajorWithMinorVersions $KeepMinorVersions `
                                -ExpireAfterDays $ExpireVersionsAfterDays -ApplyTo $effectiveApplyTo
                            Add-RunResult -SiteUrl $SiteUrl -Scope "$VersionPolicyMode (ApplyTo=$effectiveApplyTo)" -Outcome 'Applied' `
                                -Detail "Major=$KeepMajorVersions; ExpireAfterDays=$ExpireVersionsAfterDays$existingNote"
                        }
                    }
                    else {
                        Write-Output "`tNo drift. Site version policy already compliant; skipped."
                        Add-RunResult -SiteUrl $SiteUrl -Scope "$VersionPolicyMode (ApplyTo=$effectiveApplyTo)" -Outcome 'Skipped' -Detail 'No drift; already compliant'
                    }
                }
                catch {
                    Write-Warning "`tFAILED to apply site version policy on ${SiteUrl}: $($_.Exception.Message)"
                    Add-RunResult -SiteUrl $SiteUrl -Scope "$VersionPolicyMode (ApplyTo=$effectiveApplyTo)" -Outcome 'Failed' -Detail $_.Exception.Message
                }
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
                        Invoke-RetryCommand -OperationName 'New-PnPSiteFileVersionBatchDeleteJob' -ScriptBlock { New-PnPSiteFileVersionBatchDeleteJob @batchParams -Force -ErrorAction Stop }
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
        Add-RunResult -SiteUrl $SiteUrl -Scope $VersionPolicyMode -Outcome 'Failed' -Detail $_.Exception.Message
    }
    finally {
        Disconnect-PnPOnline
    }
}

#region --- Report output ---
# Worker processes do not write an HTML report; they hand their results back to the parent
# as JSON, which the orchestrator merges and renders into the single consolidated report.
if ($IsWorker) {
    try {
        $payload = @($script:RunResults) | ConvertTo-Json -Depth 6
        if ([string]::IsNullOrWhiteSpace($payload)) { $payload = '[]' }
        Set-Content -Path $WorkerResultsFile -Value $payload -Encoding UTF8 -Force -WhatIf:$false
    }
    catch {
        Write-Warning "Worker could not write its results file '$WorkerResultsFile': $($_.Exception.Message)"
    }
}

# The HTML report is a local artifact only. In Azure Automation there is no persistent
# filesystem and dumping the HTML into the job output stream makes the log unreadable, so
# the report is simply not produced there (the run summary below is still printed).
if ($EnableReport -and -not $script:IsAzureAutomationRun -and $script:RunResults.Count -gt 0) {
    $reportHtml = Export-SPSCleanVersionsReport -Results $script:RunResults `
        -Title 'SPSCleanVersions' -Version $script:ScriptVersion -DryRunMode:$WhatIfPreference
    try {
        $reportPath = Join-Path -Path $script:ResultsFolder -ChildPath ("SPSCleanVersions-$($script:RunTimestamp).html")
        Set-Content -Path $reportPath -Value $reportHtml -Encoding UTF8 -Force -WhatIf:$false
        Write-Output "HTML report written to: $reportPath"
    }
    catch {
        Write-Warning "Unable to write HTML report: $($_.Exception.Message)"
    }
}

# Run summary line.
$sumApplied = @($script:RunResults | Where-Object { $_.Outcome -eq 'Applied' }).Count
$sumWouldApply = @($script:RunResults | Where-Object { $_.Outcome -eq 'WouldApply' }).Count
$sumSkipped = @($script:RunResults | Where-Object { $_.Outcome -eq 'Skipped' -or $_.Outcome -eq 'Compliant' }).Count
$sumFailed = @($script:RunResults | Where-Object { $_.Outcome -eq 'Failed' }).Count
$appliedPart = if ($WhatIfPreference) { "$sumWouldApply would apply" } else { "$sumApplied applied" }
Write-Output "--- SPSCleanVersions finished: $($script:RunResults.Count) site(s) — $appliedPart, $sumSkipped skipped/compliant, $sumFailed failed ---"

if ($script:TranscriptStarted) {
    try { Stop-Transcript -WhatIf:$false | Out-Null } catch { }
}
#endregion
