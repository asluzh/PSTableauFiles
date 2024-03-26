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

Describe "Functional Tests for Update-TableauZipFile" -Tag Integration {
    Context "Update XML inside .twbx file" -ForEach $twbxFiles {
        It "Copy .twbx file to the test drive" {
            Copy-Item -Path $_ -Destination TestDrive:
            $script:testFilePath = Join-Path $TestDrive ([System.IO.Path]::GetFileName($_))
            $testFilePath | Should -Not -BeNullOrEmpty
            Get-Item -Path $testFilePath | Should -Not -BeNullOrEmpty
        }
        It "Update main XML in .twbx file" {
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml((Get-TableauFileXml $testFilePath))
            $xml | Should -Not -BeNullOrEmpty
            $script:testXmlContent = $xml.OuterXml
            $xml.SelectNodes("/workbook") | ForEach-Object {
                $_.SetAttribute("source-platform", "unknown");
            }
            Update-TableauZipFile -Path $testFilePath -DocumentXml $xml | Should -BeTrue
        }
        It "Get updated XML from .twbx file" {
            $updatedXml = Get-TableauFileXml $testFilePath
            $updatedXml | Should -BeOfType String
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml($updatedXml)
            $xml.OuterXml | Should -Not -Be $testXmlContent
        }
    }
    Context "Update data files inside Tableau archives" {
        It "Copy files to the test drive" {
            Copy-Item -Path "Tests/Files/Superstore-2022-3.twbx" -Destination TestDrive:
            $script:testFilePath1 = Join-Path $TestDrive "Superstore-2022-3.twbx"
            Get-Item -Path $testFilePath1 | Should -Not -BeNullOrEmpty
            Copy-Item -Path "Tests/Files/Employee-Csv-Live.tdsx" -Destination TestDrive:
            $script:testFilePath2 = Join-Path $TestDrive "Employee-Csv-Live.tdsx"
            Get-Item -Path $testFilePath2 | Should -Not -BeNullOrEmpty
            Copy-Item -Path "Tests/Files/Employee-Xlsx-Live.tdsx" -Destination TestDrive:
            $script:testFilePath3 = Join-Path $TestDrive "Employee-Xlsx-Live.tdsx"
            Get-Item -Path $testFilePath3 | Should -Not -BeNullOrEmpty
        }
        It "Update and verify data files (1)" {
            $oldEntry = [System.IO.Compression.ZipFile]::OpenRead($testFilePath1).Entries | Where-Object { $_.Name -eq "Sales Commission.csv" } | Select-Object -First 1
            Update-TableauZipFile -Path $testFilePath1 -DataFile "Tests/Files/Sales Commission.csv" | Should -BeTrue
            $newEntry = [System.IO.Compression.ZipFile]::OpenRead($testFilePath1).Entries | Where-Object { $_.Name -eq "Sales Commission.csv" } | Select-Object -First 1
            $newEntry.Length | Should -Not -Be $oldEntry.Length
        }
        It "Update and verify data files (2)" {
            $oldEntry = [System.IO.Compression.ZipFile]::OpenRead($testFilePath2).Entries | Where-Object { $_.Name -eq "Employee+Data.csv" } | Select-Object -First 1
            Update-TableauZipFile -Path $testFilePath2 -DataFile "Tests/Files/Employee+Data.csv" | Should -BeTrue
            $newEntry = [System.IO.Compression.ZipFile]::OpenRead($testFilePath2).Entries | Where-Object { $_.Name -eq "Employee+Data.csv" } | Select-Object -First 1
            $newEntry.Length | Should -Not -Be $oldEntry.Length
        }
        It "Update and verify data files (3)" {
            $oldEntry = [System.IO.Compression.ZipFile]::OpenRead($testFilePath3).Entries | Where-Object { $_.Name -eq "Employee+Master.xlsx" } | Select-Object -First 1
            Update-TableauZipFile -Path $testFilePath3 -DataFile "Tests/Files/Employee+Master.xlsx" | Should -BeTrue
            $newEntry = [System.IO.Compression.ZipFile]::OpenRead($testFilePath3).Entries | Where-Object { $_.Name -eq "Employee+Master.xlsx" } | Select-Object -First 1
            $newEntry.Length | Should -Not -Be $oldEntry.Length
        }
        It "Update main XML in .twbx file and copy the test result (1)" -Skip {
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml((Get-TableauFileXml $testFilePath1))
            $xml | Should -Not -BeNullOrEmpty
            $script:testXmlContent = $xml.OuterXml
            Update-TableauZipFile -Path $testFilePath1 -DocumentXml $xml | Should -BeTrue
            Rename-Item -Path $testFilePath1 -NewName 1.twbx
            Copy-Item -Path (Join-Path $TestDrive "1.twbx") -Destination "Tests/Files"
        }
        It "Update main XML in .twbx file and copy the test result (2)" -Skip {
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml((Get-TableauFileXml $testFilePath2))
            $xml | Should -Not -BeNullOrEmpty
            $script:testXmlContent = $xml.OuterXml
            Update-TableauZipFile -Path $testFilePath2 -DocumentXml $xml | Should -BeTrue
            Rename-Item -Path $testFilePath2 -NewName 2.tdsx
            Copy-Item -Path (Join-Path $TestDrive "2.tdsx") -Destination "Tests/Files"
        }
        It "Update main XML in .twbx file and copy the test result (3)" -Skip {
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml((Get-TableauFileXml $testFilePath3))
            $xml | Should -Not -BeNullOrEmpty
            $script:testXmlContent = $xml.OuterXml
            Update-TableauZipFile -Path $testFilePath3 -DocumentXml $xml | Should -BeTrue
            Rename-Item -Path $testFilePath3 -NewName 3.tdsx
            Copy-Item -Path (Join-Path $TestDrive "3.tdsx") -Destination "Tests/Files"
        }
    }
    Context "Update XML inside .tdsx file" -ForEach $tdsxFiles {
        It "Copy .tdsx file to the test drive" {
            Copy-Item -Path $_ -Destination TestDrive:
            $script:testFilePath = Join-Path $TestDrive ([System.IO.Path]::GetFileName($_))
            $testFilePath | Should -Not -BeNullOrEmpty
            Get-Item -Path $testFilePath | Should -Not -BeNullOrEmpty
        }
        It "Update main XML in .tdsx file" {
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml((Get-TableauFileXml $testFilePath))
            $xml | Should -Not -BeNullOrEmpty
            $script:testXmlContent = $xml.OuterXml
            $xml.SelectNodes("/datasource") | ForEach-Object {
                $_.SetAttribute("source-platform", "unknown");
            }
            Update-TableauZipFile -Path $testFilePath -DocumentXml $xml | Should -BeTrue
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
