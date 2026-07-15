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

        It 'Should not pass MajorWithMinorVersions for a new-libraries-only request' {
            $scriptContent | Should -Match '\$applyExisting\s+-and\s+\$MajorWithMinorVersions'
        }

        It 'Should warn about the app-only limitation in Azure Automation' {
            $scriptContent | Should -Match 'require a delegated user context'
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

        It 'Should define the Export-SPSCleanVersionsReport function' {
            $scriptContent | Should -Match 'function\s+Export-SPSCleanVersionsReport'
        }

        It 'Should define the ConvertTo-SPSHtmlEncoded helper' {
            $scriptContent | Should -Match 'function\s+ConvertTo-SPSHtmlEncoded'
        }

        It 'Should collect per-site results with Add-RunResult' {
            $scriptContent | Should -Match 'function\s+Add-RunResult'
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
                $script:sample.Add([PSCustomObject]@{ Site = 'https://x/sites/A'; Scope = 'ExpireAfter'; Outcome = 'Applied'; Detail = 'Major=100' })
                $script:sample.Add([PSCustomObject]@{ Site = 'https://x/sites/B'; Scope = 'ExpireAfter'; Outcome = 'Skipped'; Detail = 'No drift' })
                $script:sample.Add([PSCustomObject]@{ Site = 'https://x/sites/<C&D>'; Scope = 'Legacy'; Outcome = 'Failed'; Detail = 'boom "q" <t>' })
            }

            It 'Produces a self-contained HTML document' {
                $html = Export-SPSCleanVersionsReport -Results $script:sample -Version '3.1.0'
                $html | Should -Match '<!DOCTYPE html>'
                $html | Should -Match 'Sites processed'
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
