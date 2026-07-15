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

        It 'Should have InputJson as a mandatory parameter' {
            $inputJsonParam = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'InputJson' }
            $inputJsonParam | Should -Not -BeNullOrEmpty

            $mandatoryAttr = $inputJsonParam.Attributes | Where-Object {
                $_ -is [System.Management.Automation.Language.AttributeAst] -and
                $_.TypeName.Name -eq 'Parameter'
            }
            $mandatoryArg = $mandatoryAttr.NamedArguments | Where-Object { $_.ArgumentName -eq 'Mandatory' }
            $mandatoryArg | Should -Not -BeNullOrEmpty
        }

        It 'Should have InputJson typed as System.String' {
            $inputJsonParam = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'InputJson' }
            $typeAttr = $inputJsonParam.Attributes | Where-Object {
                $_ -is [System.Management.Automation.Language.TypeConstraintAst]
            }
            $typeAttr.TypeName.FullName | Should -Be 'System.String'
        }

        It 'Should have InputJson with ValidateNotNullOrEmpty' {
            $inputJsonParam = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'InputJson' }
            $validateAttr = $inputJsonParam.Attributes | Where-Object {
                $_ -is [System.Management.Automation.Language.AttributeAst] -and
                $_.TypeName.Name -eq 'ValidateNotNullOrEmpty'
            }
            $validateAttr | Should -Not -BeNullOrEmpty
        }

        It 'Should have ConfigFile as a mandatory parameter' {
            $configFileParam = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'ConfigFile' }
            $configFileParam | Should -Not -BeNullOrEmpty

            $mandatoryAttr = $configFileParam.Attributes | Where-Object {
                $_ -is [System.Management.Automation.Language.AttributeAst] -and
                $_.TypeName.Name -eq 'Parameter'
            }
            $mandatoryArg = $mandatoryAttr.NamedArguments | Where-Object { $_.ArgumentName -eq 'Mandatory' }
            $mandatoryArg | Should -Not -BeNullOrEmpty
        }

        It 'Should have ConfigFile typed as System.String' {
            $configFileParam = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'ConfigFile' }
            $typeAttr = $configFileParam.Attributes | Where-Object {
                $_ -is [System.Management.Automation.Language.TypeConstraintAst]
            }
            $typeAttr.TypeName.FullName | Should -Be 'System.String'
        }

        It 'Should place InputJson and ConfigFile in separate parameter sets' {
            $inputJsonParam = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'InputJson' }
            $configFileParam = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'ConfigFile' }

            $inputJsonSet = ($inputJsonParam.Attributes | Where-Object {
                    $_ -is [System.Management.Automation.Language.AttributeAst] -and $_.TypeName.Name -eq 'Parameter'
                }).NamedArguments | Where-Object { $_.ArgumentName -eq 'ParameterSetName' }
            $configFileSet = ($configFileParam.Attributes | Where-Object {
                    $_ -is [System.Management.Automation.Language.AttributeAst] -and $_.TypeName.Name -eq 'Parameter'
                }).NamedArguments | Where-Object { $_.ArgumentName -eq 'ParameterSetName' }

            $inputJsonSet | Should -Not -BeNullOrEmpty
            $configFileSet | Should -Not -BeNullOrEmpty
            $inputJsonSet.Argument.Value | Should -Not -Be $configFileSet.Argument.Value
        }

        It 'Should declare a DefaultParameterSetName on CmdletBinding' {
            $scriptContent | Should -Match 'DefaultParameterSetName'
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

        It 'Should branch on the ConfigFile parameter set' {
            $scriptContent | Should -Match "ParameterSetName\s+-eq\s+'ConfigFile'"
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
