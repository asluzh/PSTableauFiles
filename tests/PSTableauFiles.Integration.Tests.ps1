BeforeDiscovery {
    $script:twbFiles  = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.twb  -Exclude invalid.* | Resolve-Path -Relative
    $script:twbxFiles = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.twbx -Exclude invalid.* | Resolve-Path -Relative
    $script:tdsFiles  = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.tds  -Exclude invalid.* | Resolve-Path -Relative
    $script:tdsxFiles = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.tdsx -Exclude invalid.* | Resolve-Path -Relative
}
BeforeAll {
    Import-Module ./PSTableauFiles -Force
    #Requires -Modules PSTableauREST
    # Import-Module Assert
    . ./scripts/SecretStore.Functions.ps1
    # InModuleScope 'PSTableauFiles' { $script:VerbosePreference = 'Continue' } # display verbose output of module functions
    $script:VerbosePreference = 'Continue' # display verbose output of the tests
    InModuleScope 'PSTableauFiles' { $script:DebugPreference = 'Continue' } # display debug output of the module
    $script:DebugPreference = 'Continue' # display debug output of the tests
    # InModuleScope 'PSTableauFiles' { $script:ProgressPreference = 'SilentlyContinue' } # suppress progress for upload/download operations

    # Retrieve configuration of test Tableau Server for validation of content via publishing
    Get-ChildItem -Path "./tests/config" -Filter "test_*.json" | Select-Object -First 1 | ForEach-Object {
        $script:ConfigFile = Get-Content $_ | ConvertFrom-Json
        if ($ConfigFile.username) {
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "credential" -Value (New-Object System.Management.Automation.PSCredential($ConfigFile.username, (Get-SecurePassword -Namespace $ConfigFile.server -Username $ConfigFile.username)))
            Connect-TableauServer -Server $ConfigFile.server -Site $ConfigFile.site -Credential $ConfigFile.credential | Out-Null
        }
        if ($ConfigFile.pat_name) {
            $ConfigFile | Add-Member -MemberType NoteProperty -Name "pat_credential" -Value (New-Object System.Management.Automation.PSCredential($ConfigFile.pat_name, (Get-SecurePassword -Namespace $ConfigFile.server -Username $ConfigFile.pat_name)))
            Connect-TableauServer -Server $ConfigFile.server -Site $ConfigFile.site -Credential $ConfigFile.pat_credential -PersonalAccessToken | Out-Null
        }
        $project = New-TableauProject -Name (New-Guid)
        $script:testProjectId = $project.id
    }
}
AfterAll {
    if ($script:ConfigFile) {
        if ($script:testProjectId) {
            Remove-TableauProject -ProjectId $script:testProjectId | Out-Null
            $script:testProjectId = $null
        }
        Disconnect-TableauServer | Out-Null
    }
}

Describe "Functional Tests for Update-TableauZipFile" -Tag Integration {
    Context "Create Tableau packaged files" -ForEach $twbFiles {
        BeforeEach {
            $twbFile = $_
            Write-Debug "Processing $twbFile"
        }
        It "New Tableau zip file for <twbFile>" {
            $script:zipFilePath = (Join-Path $TestDrive ([System.IO.Path]::GetFileNameWithoutExtension($twbFile)))+".twbx"
            Write-Debug "Creating $zipFilePath"
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml((Get-TableauFileXml $twbFile))
            $xml | Should -Not -BeNullOrEmpty
            New-TableauZipFile -DocumentXml $xml -Path $zipFilePath
            Test-TableauZipFile -Path $zipFilePath | Should -BeTrue
        }
        It "Test publish Tableau zip file(s) on <ConfigFile.server>" {
            if ($script:ConfigFile) {
                Test-Path -Path $zipFilePath | Should -BeTrue
                $workbook = Publish-TableauWorkbook -InFile $zipFilePath -Name $twbFile -ProjectId $script:testProjectId -Overwrite
                $workbook.id | Should -BeOfType String
            } else {
                Set-ItResult -Skipped -Because "Test Tableau Server is not configured"
            }
        }
    }
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
            Copy-Item -Path "./tests/assets/Superstore-2022-3.twbx" -Destination TestDrive:
            $script:testFilePath1 = Join-Path $TestDrive "Superstore-2022-3.twbx"
            Get-Item -Path $testFilePath1 | Should -Not -BeNullOrEmpty
            Copy-Item -Path "./tests/assets/Employee-Csv-Live.tdsx" -Destination TestDrive:
            $script:testFilePath2 = Join-Path $TestDrive "Employee-Csv-Live.tdsx"
            Get-Item -Path $testFilePath2 | Should -Not -BeNullOrEmpty
            Copy-Item -Path "./tests/assets/Employee-Xlsx-Live.tdsx" -Destination TestDrive:
            $script:testFilePath3 = Join-Path $TestDrive "Employee-Xlsx-Live.tdsx"
            Get-Item -Path $testFilePath3 | Should -Not -BeNullOrEmpty
        }
        It "Update and verify data files (1)" {
            $oldEntry = [System.IO.Compression.ZipFile]::OpenRead($testFilePath1).Entries | Where-Object { $_.Name -eq "Sales Commission.csv" } | Select-Object -First 1
            Update-TableauZipFile -Path $testFilePath1 -DataFile "./tests/assets/Sales Commission.csv" | Should -BeTrue
            $newEntry = [System.IO.Compression.ZipFile]::OpenRead($testFilePath1).Entries | Where-Object { $_.Name -eq "Sales Commission.csv" } | Select-Object -First 1
            $newEntry.Length | Should -Not -Be $oldEntry.Length
        }
        It "Update and verify data files (2)" {
            $oldEntry = [System.IO.Compression.ZipFile]::OpenRead($testFilePath2).Entries | Where-Object { $_.Name -eq "Employee+Data.csv" } | Select-Object -First 1
            Update-TableauZipFile -Path $testFilePath2 -DataFile "./tests/assets/Employee+Data.csv" | Should -BeTrue
            $newEntry = [System.IO.Compression.ZipFile]::OpenRead($testFilePath2).Entries | Where-Object { $_.Name -eq "Employee+Data.csv" } | Select-Object -First 1
            $newEntry.Length | Should -Not -Be $oldEntry.Length
        }
        It "Update and verify data files (3)" {
            $oldEntry = [System.IO.Compression.ZipFile]::OpenRead($testFilePath3).Entries | Where-Object { $_.Name -eq "Employee+Master.xlsx" } | Select-Object -First 1
            Update-TableauZipFile -Path $testFilePath3 -DataFile "./tests/assets/Employee+Master.xlsx" | Should -BeTrue
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
            Copy-Item -Path (Join-Path $TestDrive "1.twbx") -Destination "./tests/assets"
        }
        It "Update main XML in .twbx file and copy the test result (2)" -Skip {
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml((Get-TableauFileXml $testFilePath2))
            $xml | Should -Not -BeNullOrEmpty
            $script:testXmlContent = $xml.OuterXml
            Update-TableauZipFile -Path $testFilePath2 -DocumentXml $xml | Should -BeTrue
            Rename-Item -Path $testFilePath2 -NewName 2.tdsx
            Copy-Item -Path (Join-Path $TestDrive "2.tdsx") -Destination "./tests/assets"
        }
        It "Update main XML in .twbx file and copy the test result (3)" -Skip {
            $xml = New-Object System.Xml.XmlDocument
            $xml.LoadXml((Get-TableauFileXml $testFilePath3))
            $xml | Should -Not -BeNullOrEmpty
            $script:testXmlContent = $xml.OuterXml
            Update-TableauZipFile -Path $testFilePath3 -DocumentXml $xml | Should -BeTrue
            Rename-Item -Path $testFilePath3 -NewName 3.tdsx
            Copy-Item -Path (Join-Path $TestDrive "3.tdsx") -Destination "./tests/assets"
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
