BeforeDiscovery {
    $script:ModuleDir = (Get-Item $PSCommandPath).Directory.Parent.FullName
    $script:ModuleName = (Split-Path -Leaf $PSCommandPath) -Replace ".Module.Tests.ps1"
    $script:ModuleFile =     "$ModuleDir/$ModuleName.psm1"
    $script:ModuleManifest = "$ModuleDir/$ModuleName.psd1"
    $script:CodeFiles = Get-ChildItem -Path "$ModuleDir" -Filter *.ps1 -Recurse
    $script:ScriptAnalyzerRules = Get-ScriptAnalyzerRule
    $script:ScriptAnalyzerResults = Invoke-ScriptAnalyzer -Path $ModuleFile -ExcludeRule PSUseBOMForUnicodeEncodedFile,PSReviewUnusedParameter
}

Describe "Module Structure and Validation Tests" -Tag Unit -WarningAction SilentlyContinue {
    Context "Module File <ModuleFile>" {
        It "has the root module <ModuleName>" {
            "$ModuleFile" | Should -Exist
        }

        It "has the a manifest file of <ModuleName>" {
            "$ModuleManifest" | Should -Exist
        }

        It "<ModuleFile> contains valid PowerShell code" {
            $psFile = Get-Content -Path $ModuleFile -ErrorAction Stop
            $errors = $null
            $null = [System.Management.Automation.PSParser]::Tokenize($psFile, [ref]$errors)
            $errors.Count | Should -Be 0
        }
    }

    Context "Code Validation <file>" -ForEach $CodeFiles {
        It "<_> is valid PowerShell code" {
            $psFile = $_
            $errors = $null
            [void][System.Management.Automation.Language.Parser]::ParseFile($psFile, [ref]$null, [ref]$errors)
            $errors.Count | Should -Be 0
        }
    }

    Context "Module Manifest of <ModuleName>" {
        It "should not throw an exception in import" {
            { Import-Module -Name $ModuleManifest -Force -ErrorAction Stop } | Should -Not -Throw
        }
    }

    Context "Testing module <ModuleName> against PSSA rules" -ForEach $ScriptAnalyzerRules {
        BeforeAll {
            Get-ScriptAnalyzerRule | Where-Object RuleName -eq $_ | Select-Object -ExpandProperty CommonName -OutVariable commonName
        }
        It "should pass rule '<commonName>'" {
            If ($ScriptAnalyzerResults.RuleName -contains $_) {
                $ScriptAnalyzerResults | Where-Object RuleName -eq $_ | Select-Object Message -OutVariable err
                $err.Message | Should -BeNullOrEmpty
            }
        }
    }

}