BeforeAll {
    Import-Module ./PSTableauFiles -Force
    # Import-Module Assert
    # . ./Tests/Test.Functions.ps1
    # InModuleScope 'PSTableauFiles' { $script:VerbosePreference = 'Continue' } # display verbose output of module functions
    $script:VerbosePreference = 'Continue' # display verbose output of the tests
}
BeforeDiscovery {
    # $script:ConfigFiles = Get-ChildItem -Path "Tests/Config" -Filter "test_*.json" -Recurse
    # $script:twbFiles  = Get-ChildItem -Recurse -Path "Tests/Files" -Filter *.twb  -Exclude invalid.*
    $script:twbxFiles = Get-ChildItem -Recurse -Path "Tests/Files" -Filter *.twbx -Exclude invalid.*
    # $script:tdsFiles  = Get-ChildItem -Recurse -Path "Tests/Files" -Filter *.tds  -Exclude invalid.*
    $script:tdsxFiles = Get-ChildItem -Recurse -Path "Tests/Files" -Filter *.tdsx -Exclude invalid.*
}

Describe "Functional Tests for Update-TableauFile" -Tag Functional {
    Context "Update XML inside .twbx file" -ForEach $twbxFiles {
        It "Copy .twbx file to the test drive" {
            Copy-Item -Path $_ -Destination TestDrive:
            $script:testFilePath = Join-Path $TestDrive ([System.IO.Path]::GetFileName($_))
            $testFilePath | Should -Not -BeNullOrEmpty
            Get-Item -Path $testFilePath | Should -Not -BeNullOrEmpty
        }
        It "Update .twbx file (XML)" {
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml((Get-TableauFileXml $testFilePath))
            $xml | Should -Not -BeNullOrEmpty
            $script:testXmlContent = $xml.OuterXml
            $xml.SelectNodes("/workbook") | ForEach-Object {
                $_.SetAttribute("source-platform", "unknown");
            }
            Update-TableauFile -Path $testFilePath -DocumentXml $xml | Should -BeTrue
        }
        It "Get updated XML from .twbx file" {
            $updatedXml = Get-TableauFileXml $testFilePath
            $updatedXml | Should -BeOfType String
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml($updatedXml)
            $xml.OuterXml | Should -Not -Be $testXmlContent
        }
    }
    Context "Update XML inside .tdsx file" -ForEach $tdsxFiles {
        It "Copy .tdsx file to the test drive" {
            Copy-Item -Path $_ -Destination TestDrive:
            $script:testFilePath = Join-Path $TestDrive ([System.IO.Path]::GetFileName($_))
            $testFilePath | Should -Not -BeNullOrEmpty
            Get-Item -Path $testFilePath | Should -Not -BeNullOrEmpty
        }
        It "Update .tdsx file (XML)" {
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml((Get-TableauFileXml $testFilePath))
            $xml | Should -Not -BeNullOrEmpty
            $script:testXmlContent = $xml.OuterXml
            $xml.SelectNodes("/datasource") | ForEach-Object {
                $_.SetAttribute("source-platform", "unknown");
            }
            Update-TableauFile -Path $testFilePath -DocumentXml $xml | Should -BeTrue
        }
        It "Get updated XML from .tdsx file" {
            $updatedXml = Get-TableauFileXml $testFilePath
            $updatedXml | Should -BeOfType String
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml($updatedXml)
            $xml.OuterXml | Should -Not -Be $testXmlContent
        }
    }
}
