BeforeAll {
    $scriptPath = Join-Path $PSScriptRoot '..' 'scripts' 'SPSCleanVersions.ps1'
    $scriptContent = Get-Content -Path $scriptPath -Raw
}

Describe 'SPSCleanVersions Script' {

    Context 'Script file validation' {

        It 'Script file should exist' {
            $scriptPath | Should -Exist
        }

        It 'Script should have valid PowerShell syntax' {
            $errors = $null
            [System.Management.Automation.PSParser]::Tokenize($scriptContent, [ref]$errors)
            $errors.Count | Should -Be 0
        }

        It 'Script should be a valid PowerShell file' {
            $scriptPath | Should -Not -BeNullOrEmpty
            (Get-Item $scriptPath).Extension | Should -Be '.ps1'
        }
    }

    Context 'Script metadata' {

        It 'Should contain a VERSION in PSScriptInfo' {
            $scriptContent | Should -Match '\.VERSION\s+\d+\.\d+\.\d+'
        }

        It 'Should contain an AUTHOR in PSScriptInfo' {
            $scriptContent | Should -Match '\.AUTHOR'
        }

        It 'Should contain a SYNOPSIS' {
            $scriptContent | Should -Match '\.SYNOPSIS'
        }

        It 'Should contain a DESCRIPTION' {
            $scriptContent | Should -Match '\.DESCRIPTION'
        }

        It 'Should contain an EXAMPLE' {
            $scriptContent | Should -Match '\.EXAMPLE'
        }
    }

    Context 'Parameters' {

        BeforeAll {
            $ast = [System.Management.Automation.Language.Parser]::ParseInput(
                $scriptContent,
                [ref]$null,
                [ref]$null
            )
            $paramBlock = $ast.ParamBlock
        }

        It 'Should define a param block' {
            $paramBlock | Should -Not -BeNullOrEmpty
        }

        It 'Should define InputJson as an optional parameter' {
            $inputJsonParam = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'InputJson' }
            $inputJsonParam | Should -Not -BeNullOrEmpty
        }

        It 'Should have InputJson typed as System.String' {
            $inputJsonParam = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'InputJson' }
            $typeAttr = $inputJsonParam.Attributes | Where-Object {
                $_ -is [System.Management.Automation.Language.TypeConstraintAst]
            }
            $typeAttr.TypeName.FullName | Should -Be 'System.String'
        }

        It 'Should define ConfigFile as an optional parameter' {
            $configFileParam = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'ConfigFile' }
            $configFileParam | Should -Not -BeNullOrEmpty
        }

        It 'Should have ConfigFile typed as System.String' {
            $configFileParam = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'ConfigFile' }
            $typeAttr = $configFileParam.Attributes | Where-Object {
                $_ -is [System.Management.Automation.Language.TypeConstraintAst]
            }
            $typeAttr.TypeName.FullName | Should -Be 'System.String'
        }

        It 'Should NOT use parameter sets (unsupported in Azure Automation runbooks)' {
            $scriptContent | Should -Not -Match 'ParameterSetName'
            $scriptContent | Should -Not -Match 'DefaultParameterSetName'
        }

        It 'Should validate that exactly one config source is supplied' {
            $scriptContent | Should -Match 'mutually exclusive; supply only one'
            $scriptContent | Should -Match 'Provide configuration via -InputJson'
        }

        It 'Should have exactly two parameters in the param block' {
            $paramBlock.Parameters.Count | Should -Be 2
        }
    }

    Context 'JSON parsing and validation' {

        It 'Should parse InputJson with ConvertFrom-Json' {
            $scriptContent | Should -Match 'ConvertFrom-Json'
        }

        It 'Should validate that SiteUrls property is required' {
            $scriptContent | Should -Match "SiteUrls.*required"
        }

        It 'Should apply default value of 50 for KeepMajorVersions' {
            $scriptContent | Should -Match 'KeepMajorVersions.*50'
        }

        It 'Should apply default value of 0 for KeepMinorVersions' {
            $scriptContent | Should -Match 'KeepMinorVersions.*0'
        }

        It 'Should apply default value of false for ForceDeleteOldVersions' {
            $scriptContent | Should -Match 'ForceDeleteOldVersions.*\$false'
        }

        It 'Should apply default value of false for DryRun' {
            $scriptContent | Should -Match 'DryRun.*\$false'
        }

        It 'Should throw on invalid JSON input' {
            $scriptContent | Should -Match 'Invalid JSON input'
        }

        It 'Should read the config file with Get-Content when ConfigFile is used' {
            $scriptContent | Should -Match 'Get-Content\s+-Path\s+\$ConfigFile\s+-Raw'
        }

        It 'Should validate the config file exists' {
            $scriptContent | Should -Match 'Configuration file not found'
        }

        It 'Should select the config file source when ConfigFile is supplied' {
            $scriptContent | Should -Match '\$hasConfigFile'
        }

        It 'Should trim the raw input before parsing' {
            $scriptContent | Should -Match '\$rawJson\s*=\s*\$rawJson\.Trim\(\)'
        }

        It 'Should strip a wrapping pair of single quotes' {
            $scriptContent | Should -Match 'StartsWith'
            $scriptContent | Should -Match 'EndsWith'
            $scriptContent | Should -Match 'Substring\(1, \$rawJson\.Length - 2\)'
        }

        It 'Should reject non-object JSON input' {
            $scriptContent | Should -Match 'InputJson must be a JSON object'
        }

        It 'Should check the parsed config is a PSCustomObject' {
            $scriptContent | Should -Match '\$config\s+-isnot\s+\[System\.Management\.Automation\.PSCustomObject\]'
        }
    }

    Context 'JSON parsing behaviour (functional)' {

        BeforeAll {
            # Reproduce the script's normalization + parsing + object guard in isolation,
            # so the behaviour is validated without importing PnP.PowerShell.
            function Invoke-ParseConfig {
                param([string]$rawJson)
                $rawJson = $rawJson.Trim()
                if ($rawJson.Length -ge 2 -and $rawJson.StartsWith("'") -and $rawJson.EndsWith("'")) {
                    $rawJson = $rawJson.Substring(1, $rawJson.Length - 2).Trim()
                }
                try {
                    $config = $rawJson | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    throw "Invalid JSON input: $($_.Exception.Message)."
                }
                if ($config -isnot [System.Management.Automation.PSCustomObject]) {
                    throw "InputJson must be a JSON object, but a $($config.GetType().Name) value was parsed."
                }
                return $config
            }
        }

        It 'Parses a clean JSON object' {
            $c = Invoke-ParseConfig '{"SiteUrls":["https://contoso.sharepoint.com/teams/CSSC"],"KeepMajorVersions":100}'
            @($c.SiteUrls).Count | Should -Be 1
        }

        It 'Auto-corrects input wrapped in single quotes' {
            $wrapped = "'{`"SiteUrls`":[`"https://contoso.sharepoint.com/teams/CSSC`"]}'"
            $c = Invoke-ParseConfig $wrapped
            @($c.SiteUrls).Count | Should -Be 1
        }

        It 'Tolerates leading and trailing whitespace' {
            $c = Invoke-ParseConfig '   {"SiteUrls":["https://x/teams/CSSC"]}   '
            @($c.SiteUrls).Count | Should -Be 1
        }

        It 'Rejects curly/smart quotes with an Invalid JSON message' {
            $l = [char]0x201C; $r = [char]0x201D
            $curly = "{$($l)SiteUrls$($r):[$($l)https://x/teams/CSSC$($r)]}"
            { Invoke-ParseConfig $curly } | Should -Throw -ExpectedMessage 'Invalid JSON input*'
        }

        It 'Rejects a bare JSON string (non-object) with a clear message' {
            { Invoke-ParseConfig '"just a string"' } | Should -Throw -ExpectedMessage 'InputJson must be a JSON object*'
        }
    }

    Context 'Script requirements' {

        It 'Should require PowerShell 7.2' {
            $scriptContent | Should -Match '#Requires\s+-Version\s+7\.2'
        }

        It 'Should require PSEdition Core' {
            $scriptContent | Should -Match '#Requires\s+-PSEdition\s+Core'
        }

        It 'Should require PnP.PowerShell module' {
            $scriptContent | Should -Match "#Requires\s+-Modules.*PnP\.PowerShell"
        }
    }

    Context 'CmdletBinding' {

        It 'Should support ShouldProcess (WhatIf)' {
            $scriptContent | Should -Match 'SupportsShouldProcess'
        }

        It 'Should have CmdletBinding attribute' {
            $scriptContent | Should -Match '\[CmdletBinding\('
        }

        It 'Should set WhatIfPreference when DryRun is specified' {
            $scriptContent | Should -Match 'if\s*\(\$DryRun\)\s*\{\s*\$WhatIfPreference\s*=\s*\$true'
        }
    }

    Context 'Azure Automation detection' {

        It 'Should define Test-IsAzureAutomation function' {
            $scriptContent | Should -Match 'function\s+Test-IsAzureAutomation'
        }

        It 'Should check multiple environment signals for Azure Automation' {
            $scriptContent | Should -Match 'AZUREPS_HOST_ENVIRONMENT'
            $scriptContent | Should -Match 'IDENTITY_ENDPOINT'
        }

        It 'Should disable PnP PowerShell update check' {
            $scriptContent | Should -Match 'PNPPOWERSHELL_UPDATECHECK'
        }

        It 'Should explicitly import the PnP.PowerShell module (runbook autoloading is unreliable)' {
            $scriptContent | Should -Match 'Import-Module\s+-Name\s+PnP\.PowerShell'
        }
    }

    Context 'Core logic patterns' {

        It 'Should iterate over SiteUrls with foreach' {
            $scriptContent | Should -Match 'foreach\s*\(\$SiteUrl\s+in\s+\$SiteUrls\)'
        }

        It 'Should connect to PnP Online' {
            $scriptContent | Should -Match 'Connect-PnPOnline'
        }

        It 'Should disconnect from PnP Online in finally block' {
            $scriptContent | Should -Match 'Disconnect-PnPOnline'
        }

        It 'Should retrieve lists with Get-PnPList' {
            $scriptContent | Should -Match 'Get-PnPList'
        }

        It 'Should filter to Document Libraries (BaseTemplate 101)' {
            $scriptContent | Should -Match 'BaseTemplate\s+-eq\s+101'
        }

        It 'Should exclude hidden lists' {
            $scriptContent | Should -Match 'Hidden\s+-eq\s+\$false'
        }

        It 'Should only process lists with versioning enabled' {
            $scriptContent | Should -Match 'EnableVersioning\s+-eq\s+\$true'
        }

        It 'Should use Set-PnPList to apply versioning changes' {
            $scriptContent | Should -Match 'Set-PnPList'
        }

        It 'Should handle Azure Automation with Managed Identity' {
            $scriptContent | Should -Match 'ManagedIdentity'
            $scriptContent | Should -Match 'Test-IsAzureAutomation'
        }

        It 'Should handle local execution with Interactive login' {
            $scriptContent | Should -Match '-Interactive'
        }

        It 'Should exclude system libraries from processing' {
            $scriptContent | Should -Match '_catalogs'
            $scriptContent | Should -Match 'SiteAssets'
            $scriptContent | Should -Match 'SitePages'
            $scriptContent | Should -Match 'Style Library'
        }
    }

    Context 'ForceDeleteOldVersions feature' {

        It 'Should call New-PnPSiteFileVersionBatchDeleteJob when ForceDeleteOldVersions is set' {
            $scriptContent | Should -Match 'New-PnPSiteFileVersionBatchDeleteJob'
        }

        It 'Should pass MajorVersionLimit to batch delete job' {
            $scriptContent | Should -Match 'MajorVersionLimit'
        }

        It 'Should pass MajorWithMinorVersionsLimit to batch delete job' {
            $scriptContent | Should -Match 'MajorWithMinorVersionsLimit'
        }

        It 'Should skip batch delete when running in Azure Automation (app-only context)' {
            $scriptContent | Should -Match 'Test-IsAzureAutomation'
        }

        It 'Should warn when batch delete is skipped due to app-only auth' {
            $scriptContent | Should -Match 'NOT supported with app-only authentication'
        }
    }

    Context 'Local single sign-in (token reuse)' {

        It 'Should establish a single delegated connection before the site loop' {
            $scriptContent | Should -Match '\$script:DelegatedAuthConnection'
            $scriptContent | Should -Match 'Connect-PnPOnline\s+-Url\s+\$anchorUrl\s+-Interactive\s+-ClientId\s+\$ClientId\s+-ReturnConnection'
        }

        It 'Should reuse a fresh access token per site via Get-PnPAccessToken' {
            $scriptContent | Should -Match 'Get-PnPAccessToken\s+-Connection\s+\$script:DelegatedAuthConnection'
            $scriptContent | Should -Match 'Connect-PnPOnline\s+-Url\s+\$SiteUrl\s+-AccessToken\s+\$accessToken'
        }

        It 'Should fall back to per-operation interactive login when the single sign-in fails' {
            $scriptContent | Should -Match 'Falling back to interactive login per operation'
            $scriptContent | Should -Match 'Connect-PnPOnline\s+-Url\s+\$SiteUrl\s+-Interactive\s+-ClientId\s+\$ClientId'
        }

        It 'Should require ClientId for local/interactive execution' {
            $scriptContent | Should -Match "ClientId is required for local/interactive execution"
        }

        It 'Should not attempt the single sign-in in Azure Automation' {
            $scriptContent | Should -Match 'if \(-not \$script:IsAzureAutomationRun\) \{'
        }

        It 'Should sign in before tenant enumeration and reuse the connection for it (single prompt)' {
            # Sign-in happens before the SiteScope=All enumeration and the connection is passed
            # to Get-TenantSiteUrls so enumeration does not trigger a second interactive prompt.
            $scriptContent | Should -Match 'Get-TenantSiteUrls -AdminUrl \$TenantAdminUrl -Filter \$SiteFilter -ClientId \$ClientId -Connection \$script:DelegatedAuthConnection'
            $scriptContent | Should -Match 'if \(\$null -ne \$Connection\) \{ \$getParams\[''Connection''\] = \$Connection \}'
        }
    }

    Context 'Site version policy feature' {

        It 'Should define the Set-SiteVersionPolicy helper function' {
            $scriptContent | Should -Match 'function\s+Set-SiteVersionPolicy'
        }
        It 'Should call Set-PnPSiteVersionPolicy' {
            $scriptContent | Should -Match 'Set-PnPSiteVersionPolicy'
        }

        It 'Should default VersionPolicyMode to Legacy' {
            $scriptContent | Should -Match "VersionPolicyMode.*'Legacy'"
        }

        It 'Should validate VersionPolicyMode against the allowed set' {
            $scriptContent | Should -Match "AutoExpiration"
            $scriptContent | Should -Match "ExpireAfter"
            $scriptContent | Should -Match "NoExpiration"
            $scriptContent | Should -Match "InheritFromTenant"
        }

        It 'Should branch to the legacy Set-PnPList path when VersionPolicyMode is Legacy' {
            $scriptContent | Should -Match "\`$VersionPolicyMode\s+-eq\s+'Legacy'"
        }

        It 'Should map AutoExpiration to EnableAutoExpirationVersionTrim' {
            $scriptContent | Should -Match 'EnableAutoExpirationVersionTrim'
        }

        It 'Should support ExpireVersionsAfterDays' {
            $scriptContent | Should -Match 'ExpireVersionsAfterDays'
        }

        It 'Should validate ExpireVersionsAfterDays is 0 or >= 30' {
            $scriptContent | Should -Match "must be 0 \(no expiration\) or greater than or equal to 30"
        }

        It 'Should map ApplyTo to ApplyToNew/ExistingDocumentLibraries' {
            $scriptContent | Should -Match 'ApplyToNewDocumentLibraries'
            $scriptContent | Should -Match 'ApplyToExistingDocumentLibraries'
        }

        It 'Should support InheritFromTenant' {
            $scriptContent | Should -Match 'InheritFromTenant'
        }
    }

    Context 'Resolve-EffectiveApplyTo (functional)' {

        BeforeAll {
            # Extract and dot-source the real Resolve-EffectiveApplyTo from the script AST so
            # we test the actual downgrade decision, not a source-text pattern. Regression for
            # #35: app-only cannot target existing document libraries, so in Azure Automation
            # Both -> New and Existing -> None; local/delegated runs are never downgraded.
            $sp = Join-Path $PSScriptRoot '..' 'scripts' 'SPSCleanVersions.ps1'
            $ast = [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path $sp), [ref]$null, [ref]$null)
            $fn = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Resolve-EffectiveApplyTo' }, $true) | Select-Object -First 1
            . ([ScriptBlock]::Create($fn.Extent.Text))
        }

        It 'Azure Automation + Both downgrades to New' {
            Resolve-EffectiveApplyTo -ApplyTo 'Both' -IsAzureAutomation $true | Should -Be 'New'
        }

        It 'Azure Automation + Existing downgrades to None' {
            Resolve-EffectiveApplyTo -ApplyTo 'Existing' -IsAzureAutomation $true | Should -Be 'None'
        }

        It 'Azure Automation + New stays New' {
            Resolve-EffectiveApplyTo -ApplyTo 'New' -IsAzureAutomation $true | Should -Be 'New'
        }

        It 'Local (delegated) never downgrades Both' {
            Resolve-EffectiveApplyTo -ApplyTo 'Both' -IsAzureAutomation $false | Should -Be 'Both'
        }

        It 'Local (delegated) never downgrades Existing' {
            Resolve-EffectiveApplyTo -ApplyTo 'Existing' -IsAzureAutomation $false | Should -Be 'Existing'
        }
    }

    Context 'App-only downgrade wiring' {

        It 'Passes the resolved effective target to the setter and warns only on a real downgrade' {
            # The loop computes $effectiveApplyTo via Resolve-EffectiveApplyTo, passes it to
            # Set-SiteVersionPolicy, warns only when it differs from the requested ApplyTo, and
            # records Skipped (without invoking the setter) when it resolves to 'None'.
            $scriptContent | Should -Match 'Resolve-EffectiveApplyTo\s+-ApplyTo\s+\$ApplyTo\s+-IsAzureAutomation'
            $scriptContent | Should -Match 'if\s*\(\s*\$effectiveApplyTo\s+-ne\s+\$ApplyTo\s*\)'
            $scriptContent | Should -Match '-ApplyTo\s+\$effectiveApplyTo'
            $scriptContent | Should -Match "if\s*\(\s*\`$effectiveApplyTo\s+-eq\s+'None'\s*\)"
        }
    }

    Context 'Site version policy (functional)' {

        BeforeAll {
            # Reproduce the ExpireVersionsAfterDays validation rules in isolation.
            function Test-ExpireDays {
                param([string]$Mode, [int]$Days)
                if ($Days -ne 0 -and $Days -lt 30) {
                    throw "'ExpireVersionsAfterDays' must be 0 (no expiration) or greater than or equal to 30."
                }
                if ($Mode -eq 'ExpireAfter' -and $Days -lt 30) {
                    throw "VersionPolicyMode 'ExpireAfter' requires 'ExpireVersionsAfterDays' to be greater than or equal to 30."
                }
                return $true
            }
        }

        It 'Accepts 0 (no expiration)' {
            Test-ExpireDays -Mode 'NoExpiration' -Days 0 | Should -BeTrue
        }

        It 'Accepts a value >= 30' {
            Test-ExpireDays -Mode 'ExpireAfter' -Days 180 | Should -BeTrue
        }

        It 'Rejects a value between 1 and 29' {
            { Test-ExpireDays -Mode 'ExpireAfter' -Days 10 } | Should -Throw
        }

        It 'Rejects ExpireAfter with 0 days' {
            { Test-ExpireDays -Mode 'ExpireAfter' -Days 0 } | Should -Throw -ExpectedMessage "*requires 'ExpireVersionsAfterDays'*"
        }
    }

    Context 'Set-SiteVersionPolicy bound parameters (functional)' {

        BeforeAll {
            # Extract and dot-source the REAL Set-SiteVersionPolicy from the script AST (same
            # approach as the report tests) so we assert the parameters actually bound to
            # Set-PnPSiteVersionPolicy, not a copy of the logic. Regression for #33:
            # MajorWithMinorVersions must be bound (including 0) for existing libraries and
            # omitted for a new-libraries-only request.
            $sp = Join-Path $PSScriptRoot '..' 'scripts' 'SPSCleanVersions.ps1'
            $ast = [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path $sp), [ref]$null, [ref]$null)
            # Set-SiteVersionPolicy now calls Invoke-RetryCommand (which uses Get-RetryAfterDelay
            # and Test-IsAuthError), so dot-source those helpers too or the call would fail.
            $wanted = 'Set-SiteVersionPolicy', 'Invoke-RetryCommand', 'Get-RetryAfterDelay', 'Test-IsAuthError'
            $funcs = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $wanted -contains $n.Name }, $true)
            foreach ($f in $funcs) { . ([ScriptBlock]::Create($f.Extent.Text)) }

            # Local stub so Set-PnPSiteVersionPolicy is mockable without importing
            # PnP.PowerShell. Its parameters mirror what the script splats, so the mock's
            # ParameterFilter can inspect the bound values.
            function Set-PnPSiteVersionPolicy {
                [CmdletBinding()]
                param(
                    [switch] $EnableAutoExpirationVersionTrim,
                    [int] $ExpireVersionsAfterDays,
                    [int] $MajorVersions,
                    [int] $MajorWithMinorVersions,
                    [switch] $ApplyToNewDocumentLibraries,
                    [switch] $ApplyToExistingDocumentLibraries,
                    [switch] $InheritFromTenant
                )
            }
        }

        BeforeEach {
            Mock Set-PnPSiteVersionPolicy { }
        }

        It 'ExpireAfter + Existing with 0 minor: binds MajorWithMinorVersions = 0' {
            Set-SiteVersionPolicy -SiteUrl 'https://x/sites/A' -Mode 'ExpireAfter' -MajorVersions 100 -MajorWithMinorVersions 0 -ExpireAfterDays 365 -ApplyTo 'Existing'
            Should -Invoke Set-PnPSiteVersionPolicy -Times 1 -Exactly -ParameterFilter {
                $null -ne $MajorWithMinorVersions -and $MajorWithMinorVersions -eq 0 -and $ApplyToExistingDocumentLibraries
            }
        }

        It 'ExpireAfter + Both with 0 minor: binds MajorWithMinorVersions = 0' {
            Set-SiteVersionPolicy -SiteUrl 'https://x/sites/A' -Mode 'ExpireAfter' -MajorVersions 100 -MajorWithMinorVersions 0 -ExpireAfterDays 365 -ApplyTo 'Both'
            Should -Invoke Set-PnPSiteVersionPolicy -Times 1 -Exactly -ParameterFilter {
                $null -ne $MajorWithMinorVersions -and $MajorWithMinorVersions -eq 0
            }
        }

        It 'NoExpiration + Existing with 0 minor: binds MajorWithMinorVersions = 0' {
            Set-SiteVersionPolicy -SiteUrl 'https://x/sites/A' -Mode 'NoExpiration' -MajorVersions 100 -MajorWithMinorVersions 0 -ExpireAfterDays 0 -ApplyTo 'Existing'
            Should -Invoke Set-PnPSiteVersionPolicy -Times 1 -Exactly -ParameterFilter {
                $null -ne $MajorWithMinorVersions -and $MajorWithMinorVersions -eq 0
            }
        }

        It 'ExpireAfter + Existing with a positive minor count: binds that value' {
            Set-SiteVersionPolicy -SiteUrl 'https://x/sites/A' -Mode 'ExpireAfter' -MajorVersions 100 -MajorWithMinorVersions 5 -ExpireAfterDays 365 -ApplyTo 'Existing'
            Should -Invoke Set-PnPSiteVersionPolicy -Times 1 -Exactly -ParameterFilter {
                $MajorWithMinorVersions -eq 5
            }
        }

        It 'ExpireAfter + New only: omits MajorWithMinorVersions and the existing target' {
            Set-SiteVersionPolicy -SiteUrl 'https://x/sites/A' -Mode 'ExpireAfter' -MajorVersions 100 -MajorWithMinorVersions 0 -ExpireAfterDays 365 -ApplyTo 'New'
            Should -Invoke Set-PnPSiteVersionPolicy -Times 1 -Exactly -ParameterFilter {
                ($null -eq $MajorWithMinorVersions) -and ($null -eq $ApplyToExistingDocumentLibraries) -and $ApplyToNewDocumentLibraries
            }
        }
    }

    Context 'Site version policy drift detection' {

        It 'Should define the Test-SiteVersionPolicyDrift helper function' {
            $scriptContent | Should -Match 'function\s+Test-SiteVersionPolicyDrift'
        }

        It 'Should read the current policy with Get-PnPSiteVersionPolicy' {
            $scriptContent | Should -Match 'Get-PnPSiteVersionPolicy'
        }

        It 'Should compare against the DefaultTrimMode field' {
            $scriptContent | Should -Match 'DefaultTrimMode'
        }

        It 'Should only apply the site policy when a drift is detected' {
            $scriptContent | Should -Match '\$hasDrift'
            $scriptContent | Should -Match 'No drift'
        }

        It 'Should gate MajorWithMinorVersions on existing libraries, not on a minor count > 0' {
            # Regression for #33: for existing document libraries in ExpireAfter/NoExpiration
            # mode, SharePoint requires MajorWithMinorVersions even when it is 0, so the
            # parameter must be gated on $applyExisting (plus the mode), never on a > 0 guard.
            $scriptContent | Should -Match '\$applyExisting\s+-and\s+\(\$Mode\s+-eq'
            $scriptContent | Should -Not -Match '\$MajorWithMinorVersions\s+-gt\s+0'
        }

        It 'Should warn about the app-only limitation in Azure Automation' {
            # #35: the warning now targets the specific unsupported operation (existing
            # document libraries) rather than a blanket delegated-context message.
            $scriptContent | Should -Match 'cannot apply the version policy to EXISTING'
            $scriptContent | Should -Match 'Cannot call this API with an app-only principal'
        }

        It 'Should define the Get-TenantSiteUrls helper function' {
            $scriptContent | Should -Match 'function\s+Get-TenantSiteUrls'
        }

        It 'Should enumerate tenant sites with Get-PnPTenantSite when SiteScope is All' {
            $scriptContent | Should -Match 'Get-PnPTenantSite'
        }

        It 'Should default SiteScope to Selected' {
            $scriptContent | Should -Match "SiteScope.*'Selected'"
        }

        It 'Should require TenantAdminUrl when SiteScope is All' {
            $scriptContent | Should -Match "'TenantAdminUrl' is required when 'SiteScope' is 'All'"
        }

        It 'Should reject SiteScope All with Legacy mode' {
            $scriptContent | Should -Match "'SiteScope' = 'All' is only supported"
        }

        It 'Should make SiteUrls optional when SiteScope is All' {
            $scriptContent | Should -Match "or set 'SiteScope' to 'All'"
        }

        It 'Should support a server-side SiteFilter for enumeration' {
            $scriptContent | Should -Match 'SiteFilter'
        }

        It 'Should not pollute the enumerated site list with informational Write-Output' {
            # Get-TenantSiteUrls must use Write-Verbose for status so its return value stays
            # clean (a Write-Output there would be captured as bogus site URLs).
            $scriptContent | Should -Match 'Write-Verbose "Connecting to tenant admin center'
            $scriptContent | Should -Match 'return , \[string\[\]\]\$urls'
        }

        It 'Should record a simulated outcome in DryRun instead of Applied' {
            $scriptContent | Should -Match "Outcome 'WouldApply'"
            $scriptContent | Should -Match 'Would apply site version policy'
        }

        It 'Should report a would-apply count in the DryRun summary' {
            $scriptContent | Should -Match 'would apply'
        }

        Context 'Drift comparison (functional, real field shapes)' {

            BeforeAll {
                # Reproduce the drift comparison in isolation using the exact field shapes
                # returned by Get-PnPSiteVersionPolicy (captured from a live tenant).
                $script:mockPolicy = $null
                function Get-PnPSiteVersionPolicy { param() ; return $script:mockPolicy }

                function Test-Drift {
                    param(
                        [string] $Mode, [int] $MajorVersions, [int] $ExpireAfterDays
                    )
                    try { $current = Get-PnPSiteVersionPolicy -ErrorAction Stop } catch { return $true }
                    if ($null -eq $current) { return ($Mode -ne 'InheritFromTenant') }
                    $curTrimMode = $current.PSObject.Properties['DefaultTrimMode'].Value
                    $curExpire = $current.PSObject.Properties['DefaultExpireAfterDays'].Value
                    $curMajor = $current.PSObject.Properties['MajorVersionLimit'].Value
                    $hasSitePolicy = -not [string]::IsNullOrWhiteSpace([string]$curTrimMode)
                    switch ($Mode) {
                        'InheritFromTenant' { return $hasSitePolicy }
                        'AutoExpiration' { if (-not $hasSitePolicy) { return $true }; return ("$curTrimMode" -ine 'AutoExpiration') }
                        default {
                            if (-not $hasSitePolicy) { return $true }
                            if ("$curTrimMode" -ine $Mode) { return $true }
                            if ([string]::IsNullOrWhiteSpace([string]$curMajor) -or [int]$curMajor -ne $MajorVersions) { return $true }
                            $desiredExpire = if ($Mode -eq 'NoExpiration') { 0 } else { $ExpireAfterDays }
                            $curExpireInt = if ([string]::IsNullOrWhiteSpace([string]$curExpire)) { 0 } else { [int]$curExpire }
                            if ($curExpireInt -ne $desiredExpire) { return $true }
                            return $false
                        }
                    }
                }

                $script:noPolicy = [PSCustomObject]@{ Url = 'x'; DefaultTrimMode = ''; DefaultExpireAfterDays = ''; MajorVersionLimit = ''; Description = 'No Site Level Policy Set for new document libraries' }
                $script:expireAfter = [PSCustomObject]@{ Url = 'x'; DefaultTrimMode = 'ExpireAfter'; DefaultExpireAfterDays = '180'; MajorVersionLimit = '100'; Description = 'Site has Manual settings...' }
            }

            It 'No site policy + InheritFromTenant = no drift' {
                $script:mockPolicy = $script:noPolicy
                Test-Drift -Mode 'InheritFromTenant' | Should -BeFalse
            }

            It 'No site policy + ExpireAfter = drift' {
                $script:mockPolicy = $script:noPolicy
                Test-Drift -Mode 'ExpireAfter' -MajorVersions 100 -ExpireAfterDays 180 | Should -BeTrue
            }

            It 'Matching ExpireAfter 180/100 = no drift' {
                $script:mockPolicy = $script:expireAfter
                Test-Drift -Mode 'ExpireAfter' -MajorVersions 100 -ExpireAfterDays 180 | Should -BeFalse
            }

            It 'ExpireAfter with different days = drift' {
                $script:mockPolicy = $script:expireAfter
                Test-Drift -Mode 'ExpireAfter' -MajorVersions 100 -ExpireAfterDays 90 | Should -BeTrue
            }

            It 'ExpireAfter with different major count = drift' {
                $script:mockPolicy = $script:expireAfter
                Test-Drift -Mode 'ExpireAfter' -MajorVersions 50 -ExpireAfterDays 180 | Should -BeTrue
            }

            It 'Explicit policy present + InheritFromTenant = drift' {
                $script:mockPolicy = $script:expireAfter
                Test-Drift -Mode 'InheritFromTenant' | Should -BeTrue
            }

            It 'Unreadable policy = drift (fail-safe)' {
                function Get-PnPSiteVersionPolicy { throw 'unauthorized' }
                Test-Drift -Mode 'ExpireAfter' -MajorVersions 100 -ExpireAfterDays 180 | Should -BeTrue
            }
        }
    }

    Context 'Logging and HTML report' {

        It 'Should write local artifacts even in DryRun (WhatIf bypass)' {
            # #39: DryRun sets $WhatIfPreference globally, which also suppressed the tool's own
            # artifact writes. Folder creation, transcript, HTML report and retention pruning
            # must use -WhatIf:$false so a local DryRun still produces a report/transcript.
            $scriptContent | Should -Match 'New-Item -Path \$dir -ItemType Directory -Force -WhatIf:\$false'
            $scriptContent | Should -Match 'Start-Transcript -Path \$transcriptPath -IncludeInvocationHeader -WhatIf:\$false'
            $scriptContent | Should -Match 'Set-Content -Path \$reportPath .* -WhatIf:\$false'
            $scriptContent | Should -Match 'Remove-Item -Path \$_\.FullName -Force -ErrorAction SilentlyContinue -WhatIf:\$false'
            $scriptContent | Should -Match 'Stop-Transcript -WhatIf:\$false'
        }

        It 'Should define the Export-SPSCleanVersionsReport function' {
            $scriptContent | Should -Match 'function\s+Export-SPSCleanVersionsReport'
        }

        It 'Should define the ConvertTo-SPSHtmlEncoded helper' {
            $scriptContent | Should -Match 'function\s+ConvertTo-SPSHtmlEncoded'
        }

        It 'Should collect per-site results with Add-RunResult' {
            $scriptContent | Should -Match 'function\s+Add-RunResult'
        }

        It 'Add-RunResult carries the structured Library/Major/Minor/ExpireAfterDays fields' {
            $scriptContent | Should -Match "Add-RunResult\b[\s\S]*Library\s*=\s*\`$Library"
            $scriptContent | Should -Match 'Major\s*=\s*\$Major'
            $scriptContent | Should -Match 'Minor\s*=\s*\$Minor'
            $scriptContent | Should -Match 'ExpireAfterDays\s*=\s*\$ExpireAfterDays'
        }

        It 'Legacy mode emits one result row per library' {
            $scriptContent | Should -Match "Add-RunResult -SiteUrl \`$SiteUrl -Scope 'Legacy' -Library \`$list\.Title -Outcome 'Applied'"
            $scriptContent | Should -Match "Add-RunResult -SiteUrl \`$SiteUrl -Scope 'Legacy' -Library \`$list\.Title -Outcome 'Compliant'"
            $scriptContent | Should -Match "Add-RunResult -SiteUrl \`$SiteUrl -Scope 'Legacy' -Library \`$list\.Title -Outcome 'WouldApply'"
        }

        It 'Site-policy mode can enumerate in-scope libraries when EnumerateLibraries is set' {
            $scriptContent | Should -Match "config.PSObject.Properties\['EnumerateLibraries'\]"
            # Only enumerate when the effective target actually includes existing libraries.
            $scriptContent | Should -Match '\$EnumerateLibraries -and \(\$effectiveApplyTo -eq ''Both'' -or \$effectiveApplyTo -eq ''Existing''\)'
            $scriptContent | Should -Match "-Outcome 'InScope'"
        }

        It 'Writes a machine-readable JSON alongside the HTML report' {
            $scriptContent | Should -Match 'SPSCleanVersions-\$\(\$script:RunTimestamp\)\.json'
            $scriptContent | Should -Match 'JSON results written to:'
            # Guard against the "Argument types do not match" failure: serialize the List via
            # .ToArray() / -InputObject, never `@($List) | ConvertTo-Json`.
            $scriptContent | Should -Match 'ConvertTo-Json -InputObject \$jsonRows'
            $scriptContent | Should -Not -Match '@\(\$script:RunResults\) \| ConvertTo-Json'
            # JSON emission is outside the non-empty HTML gate (empty run still yields []).
            $scriptContent | Should -Match "if \(\`$EnableReport -and -not \`$script:IsAzureAutomationRun\) \{"
            # results.json is covered by retention pruning.
            $scriptContent | Should -Match "Clear-OldRunFiles -Path \`$script:ResultsFolder -Retention \`$LogRetentionDays -Filter '\*\.json'"
        }

        It 'Legacy real-mutation path keeps the ShouldProcess (-Confirm) gate' {
            $scriptContent | Should -Match "ShouldProcess\(\`$list\.Title, 'Set versioning policy'\)"
        }

        It 'HTML report card counts distinct sites, not rows' {
            $scriptContent | Should -Match '\$distinctSites = @\(\$rows'
            $scriptContent | Should -Match '<div class="card-value">\$distinctSites</div><div class="card-label">Sites processed</div>'
        }

        It 'Should start a transcript for local runs' {
            $scriptContent | Should -Match 'Start-Transcript'
            $scriptContent | Should -Match 'Stop-Transcript'
        }

        It 'Should keep the HTML report local-only (not dumped in Azure Automation)' {
            $scriptContent | Should -Not -Match 'BEGIN SPSCleanVersions HTML report'
            $scriptContent | Should -Match '-not \$script:IsAzureAutomationRun'
        }

        It 'Should default EnableReport to true' {
            $scriptContent | Should -Match "EnableReport.*\`$true"
        }

        It 'Should prune old files with Clear-OldRunFiles' {
            $scriptContent | Should -Match 'function\s+Clear-OldRunFiles'
        }

        Context 'Report generation (functional)' {

            BeforeAll {
                $sp = Join-Path $PSScriptRoot '..' 'scripts' 'SPSCleanVersions.ps1'
                $ast = [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path $sp), [ref]$null, [ref]$null)
                $wanted = 'ConvertTo-SPSHtmlEncoded', 'Export-SPSCleanVersionsReport'
                $funcs = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $wanted -contains $n.Name }, $true)
                foreach ($f in $funcs) { . ([ScriptBlock]::Create($f.Extent.Text)) }

                $script:sample = New-Object System.Collections.Generic.List[object]
                $script:sample.Add([PSCustomObject]@{ Site = 'https://x/sites/A'; Scope = 'Legacy'; Library = 'Documents'; Outcome = 'Applied'; Major = '100'; Minor = '0'; ExpireAfterDays = ''; Detail = 'Set Major=100' })
                $script:sample.Add([PSCustomObject]@{ Site = 'https://x/sites/B'; Scope = 'ExpireAfter'; Library = ''; Outcome = 'WouldApply'; Major = '100'; Minor = ''; ExpireAfterDays = '365'; Detail = 'DryRun' })
                $script:sample.Add([PSCustomObject]@{ Site = 'https://x/sites/<C&D>'; Scope = 'Legacy'; Library = 'Lib<x>'; Outcome = 'Failed'; Major = '50'; Minor = '0'; ExpireAfterDays = ''; Detail = 'boom "q" <t>' })
            }

            It 'Produces a self-contained HTML document' {
                $html = Export-SPSCleanVersionsReport -Results $script:sample -Version '3.1.0'
                $html | Should -Match '<!DOCTYPE html>'
                $html | Should -Match 'Sites processed'
            }

            It 'Includes the Library, Major, Minor and ExpireAfterDays columns' {
                $html = Export-SPSCleanVersionsReport -Results $script:sample -Version '3.1.0'
                $html | Should -Match '<th>Library</th>'
                $html | Should -Match '<th>Major</th>'
                $html | Should -Match '<th>Minor</th>'
                $html | Should -Match '<th>ExpireAfterDays</th>'
                # A per-library value and an ExpireAfterDays value are rendered.
                $html | Should -Match '>Documents<'
                $html | Should -Match '>365<'
            }

            It 'Uses the house-style sticky brand banner and centered layout' {
                $html = Export-SPSCleanVersionsReport -Results $script:sample -Version '3.1.0'
                $html | Should -Match '<header class="banner">'
                $html | Should -Match '<div class="layout">'
            }

            It 'Highlights failed rows with the alert row class' {
                $html = Export-SPSCleanVersionsReport -Results $script:sample -Version '3.1.0'
                $html | Should -Match 'class="row-alert"'
            }

            It 'HTML-encodes dangerous values (no injection)' {
                $html = Export-SPSCleanVersionsReport -Results $script:sample -Version '3.1.0'
                $html | Should -Match '&lt;C&amp;D&gt;'
                $html | Should -Match '&quot;q&quot;'
                $html | Should -Not -Match '<C&D>'
            }

            It 'Shows a DryRun badge when DryRunMode is set' {
                $html = Export-SPSCleanVersionsReport -Results $script:sample -Version '3.1.0' -DryRunMode:$true
                $html | Should -Match 'DryRun'
            }

            It 'Uses a Would apply card in DryRun mode' {
                $wa = New-Object System.Collections.Generic.List[object]
                $wa.Add([PSCustomObject]@{ Site = 'https://x/sites/A'; Scope = 'ExpireAfter'; Outcome = 'WouldApply'; Detail = 'DryRun' })
                $html = Export-SPSCleanVersionsReport -Results $wa -Version '3.1.3' -DryRunMode:$true
                $html | Should -Match 'Would apply'
                $html | Should -Match 'badge WouldApply'
            }

            It 'Marks the overall status ATTENTION when a site failed' {
                $html = Export-SPSCleanVersionsReport -Results $script:sample -Version '3.1.0'
                $html | Should -Match 'ATTENTION'
            }
        }
    }

    Context 'Throttling / retry helpers (functional)' {

        BeforeAll {
            $sp = Join-Path $PSScriptRoot '..' 'scripts' 'SPSCleanVersions.ps1'
            $ast = [System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path $sp), [ref]$null, [ref]$null)
            $wanted = 'Invoke-RetryCommand', 'Get-RetryAfterDelay', 'Test-IsAuthError'
            $funcs = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $wanted -contains $n.Name }, $true)
            foreach ($f in $funcs) { . ([ScriptBlock]::Create($f.Extent.Text)) }
        }

        It 'Defines the three throttling helpers' {
            $scriptContent | Should -Match 'function\s+Get-RetryAfterDelay'
            $scriptContent | Should -Match 'function\s+Test-IsAuthError'
            $scriptContent | Should -Match 'function\s+Invoke-RetryCommand'
        }

        It 'Wraps the SharePoint calls with Invoke-RetryCommand' {
            $scriptContent | Should -Match "Invoke-RetryCommand -OperationName 'Get-PnPSiteVersionPolicy'"
            $scriptContent | Should -Match 'Invoke-RetryCommand -OperationName "Set-PnPSiteVersionPolicy'
            $scriptContent | Should -Match "Invoke-RetryCommand -OperationName 'Get-PnPTenantSite'"
            $scriptContent | Should -Match "Invoke-RetryCommand -OperationName 'Get-PnPList'"
            $scriptContent | Should -Match 'Invoke-RetryCommand -OperationName "Set-PnPList'
            $scriptContent | Should -Match "Invoke-RetryCommand -OperationName 'New-PnPSiteFileVersionBatchDeleteJob'"
        }

        It 'Invoke-RetryCommand returns the script block result without retrying on success' {
            $script:calls = 0
            $result = Invoke-RetryCommand -OperationName 'ok' -ScriptBlock { $script:calls++; 'value' }
            $result | Should -Be 'value'
            $script:calls | Should -Be 1
        }

        It 'Invoke-RetryCommand retries then succeeds' {
            $script:calls = 0
            $result = Invoke-RetryCommand -OperationName 'transient' -BaseDelaySeconds 1 -MaxRetries 3 -ScriptBlock {
                $script:calls++
                if ($script:calls -lt 2) { throw 'transient failure' }
                'ok'
            }
            $result | Should -Be 'ok'
            $script:calls | Should -Be 2
        }

        It 'Invoke-RetryCommand rethrows after exhausting retries' {
            $script:calls = 0
            { Invoke-RetryCommand -OperationName 'always' -BaseDelaySeconds 1 -MaxRetries 2 -ScriptBlock {
                    $script:calls++; throw 'boom'
                } } | Should -Throw
            $script:calls | Should -Be 3
        }

        It 'Invoke-RetryCommand honours a Retry-After hint (capped) instead of backoff' {
            # Mock Start-Sleep so the test is fast and we can assert the delay used. The first
            # attempt throws a throttling error carrying "Retry-After: 120"; the retry must
            # sleep for that server-provided value (120s), not the exponential backoff.
            $script:sleptFor = $null
            Mock -CommandName Start-Sleep -MockWith { param($Seconds) $script:sleptFor = $Seconds }
            $script:calls = 0
            $result = Invoke-RetryCommand -OperationName 'throttled' -BaseDelaySeconds 5 -MaxRetries 3 -ScriptBlock {
                $script:calls++
                if ($script:calls -lt 2) { throw 'Request was throttled. Retry-After: 120' }
                'ok'
            }
            $result | Should -Be 'ok'
            $script:sleptFor | Should -Be 120
        }

        It 'Invoke-RetryCommand caps a very large Retry-After at 300s' {
            $script:sleptFor = $null
            Mock -CommandName Start-Sleep -MockWith { param($Seconds) $script:sleptFor = $Seconds }
            $script:calls = 0
            $null = Invoke-RetryCommand -OperationName 'throttled-big' -BaseDelaySeconds 5 -MaxRetries 3 -ScriptBlock {
                $script:calls++
                if ($script:calls -lt 2) { throw 'Throttled. Retry-After: 999' }
                'ok'
            }
            $script:sleptFor | Should -Be 300
        }

        It 'Get-RetryAfterDelay reads a Retry-After hint from the message' {
            $err = try { throw 'Request throttled. Retry-After: 42' } catch { $_ }
            Get-RetryAfterDelay -ErrorRecord $err | Should -Be 42
        }

        It 'Get-RetryAfterDelay returns 0 when no hint is present' {
            $err = try { throw 'some unrelated error' } catch { $_ }
            Get-RetryAfterDelay -ErrorRecord $err | Should -Be 0
        }

        It 'Test-IsAuthError detects auth failures and ignores others' {
            $auth = try { throw 'AADSTS700082 token is expired' } catch { $_ }
            $other = try { throw 'file not found' } catch { $_ }
            Test-IsAuthError -ErrorRecord $auth | Should -BeTrue
            Test-IsAuthError -ErrorRecord $other | Should -BeFalse
        }
    }

    Context 'Error handling' {

        It 'Should use try/catch/finally pattern' {
            $scriptContent | Should -Match 'try\s*\{'
            $scriptContent | Should -Match 'catch\s*\{'
            $scriptContent | Should -Match 'finally\s*\{'
        }

        It 'Should use -ErrorAction Stop for Set-PnPList' {
            $scriptContent | Should -Match 'Set-PnPList\s+@p\s+-ErrorAction\s+Stop'
        }

        It 'Should write errors for failed site processing' {
            $scriptContent | Should -Match 'Write-Error'
        }

        It 'Should write warnings for failed list updates' {
            $scriptContent | Should -Match 'Write-Warning'
        }
    }
}
