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

        It 'Should have InputJson as the only mandatory parameter' {
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

        It 'Should have exactly one parameter in the param block' {
            $paramBlock.Parameters.Count | Should -Be 1
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
