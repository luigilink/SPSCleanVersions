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

        It 'Should have SiteUrls as a mandatory parameter' {
            $siteUrlsParam = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'SiteUrls' }
            $siteUrlsParam | Should -Not -BeNullOrEmpty

            $mandatoryAttr = $siteUrlsParam.Attributes | Where-Object {
                $_ -is [System.Management.Automation.Language.AttributeAst] -and
                $_.TypeName.Name -eq 'Parameter'
            }
            $mandatoryArg = $mandatoryAttr.NamedArguments | Where-Object { $_.ArgumentName -eq 'Mandatory' }
            $mandatoryArg | Should -Not -BeNullOrEmpty
        }

        It 'Should have SiteUrls typed as System.String[]' {
            $siteUrlsParam = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'SiteUrls' }
            $typeAttr = $siteUrlsParam.Attributes | Where-Object {
                $_ -is [System.Management.Automation.Language.TypeConstraintAst]
            }
            $typeAttr.TypeName.FullName | Should -Be 'System.String[]'
        }

        It 'Should have KeepMajorVersions parameter with default value 50' {
            $param = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'KeepMajorVersions' }
            $param | Should -Not -BeNullOrEmpty
            $param.DefaultValue.ToString() | Should -Be '50'
        }

        It 'Should have KeepMinorVersions parameter with default value 0' {
            $param = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'KeepMinorVersions' }
            $param | Should -Not -BeNullOrEmpty
            $param.DefaultValue.ToString() | Should -Be '0'
        }

        It 'Should have ClientId as an optional string parameter' {
            $param = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'ClientId' }
            $param | Should -Not -BeNullOrEmpty

            $mandatoryAttr = $param.Attributes | Where-Object {
                $_ -is [System.Management.Automation.Language.AttributeAst] -and
                $_.TypeName.Name -eq 'Parameter'
            }
            $mandatoryArg = $mandatoryAttr.NamedArguments | Where-Object {
                $_.ArgumentName -eq 'Mandatory' -and $_.ExpressionOmitted -eq $false
            }
            if ($mandatoryArg) {
                $mandatoryArg.Argument.SafeGetValue() | Should -Be $false
            }
        }

        It 'Should have ForceDeleteOldVersions as a switch parameter' {
            $param = $paramBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'ForceDeleteOldVersions' }
            $param | Should -Not -BeNullOrEmpty

            $typeAttr = $param.Attributes | Where-Object {
                $_ -is [System.Management.Automation.Language.TypeConstraintAst]
            }
            $typeAttr.TypeName.Name | Should -Be 'switch'
        }
    }

    Context 'CmdletBinding' {

        It 'Should support ShouldProcess (WhatIf)' {
            $scriptContent | Should -Match 'SupportsShouldProcess'
        }

        It 'Should have CmdletBinding attribute' {
            $scriptContent | Should -Match '\[CmdletBinding\('
        }
    }

    Context 'Core logic patterns' {

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
            $scriptContent | Should -Match 'AUTOMATION_ASSET_NAME'
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
