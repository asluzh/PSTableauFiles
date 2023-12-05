BeforeAll {
    Import-Module ./PSTableauFiles -Force
    # Import-Module Assert
    # . ./Tests/Test.Functions.ps1
    # InModuleScope 'PSTableauFiles' { $script:VerbosePreference = 'Continue' } # display verbose output of module functions
    $script:VerbosePreference = 'Continue' # display verbose output of the tests
}
BeforeDiscovery {
    # $script:ConfigFiles = Get-ChildItem -Path "Tests/Config" -Filter "test_*.json" -Recurse
    $script:twbFiles  = Get-ChildItem -Recurse -Path "Tests/Files" -Filter *.twb  -Exclude invalid.*
    $script:twbxFiles = Get-ChildItem -Recurse -Path "Tests/Files" -Filter *.twbx -Exclude invalid.*
    $script:tdsFiles  = Get-ChildItem -Recurse -Path "Tests/Files" -Filter *.tds  -Exclude invalid.*
    $script:tdsxFiles = Get-ChildItem -Recurse -Path "Tests/Files" -Filter *.tdsx -Exclude invalid.*
}

Describe "Unit Tests for Get-TableauFileXml" -Tag Unit {
    Context "Getting XML from .twb files" -ForEach $twbFiles {
        It "<_> content returned" {
            Get-TableauFileXml $_ | Should -Not -BeNullOrEmpty
        }
    }
    Context "Getting XML from .twbx files" -ForEach $twbxFiles {
        It "<_> content returned" {
            Get-TableauFileXml $_ | Should -Not -BeNullOrEmpty
        }
    }
    Context "Getting XML from .tds files" -ForEach $tdsFiles {
        It "<_> content returned" {
            Get-TableauFileXml -Path $_ | Should -Not -BeNullOrEmpty
        }
    }
    Context "Getting XML from .tdsx files" -ForEach $tdsxFiles {
        It "<_> content returned" {
            Get-TableauFileXml -Path $_ | Should -Not -BeNullOrEmpty
        }
    }
    Context "Exceptions" {
        It "Missing files should throw exception" {
            {Get-TableauFileXml "Tests/Files/missing.twbx"} | Should -Throw -ExpectedMessage "File not found*"
        }
        It "Unknown file types should throw exception" {
            {Get-TableauFileXml "Tests/Files/invalid.twby"} | Should -Throw -ExpectedMessage "Unknown file type*"
        }
        It "Invalid TWBX should throw exception" {
            {Get-TableauFileXml "Tests/Files/invalid.twbx"} | Should -Throw -ExpectedMessage "Main XML file not found*"
        }
        It "Invalid Zip file should throw exception" {
            $err = {Get-TableauFileXml "Tests/Files/invalid.zip.tdsx"} | Should -Throw -PassThru
            $err.Exception.InnerException.Message | Should -Be "End of Central Directory record could not be found."
        }
    }
    Context "Test Zip File .twbx" -ForEach $twbxFiles {
        It "<_> is valid Zip file" {
            InModuleScope PSTableauFiles -Parameters @{ file = $_ } {
                $file | Should -Not -BeNullOrEmpty
                Test-TableauZipFile $file | Should -BeTrue
            }
        }
    }
    Context "Test Zip File .tsdx" -ForEach $tdsxFiles {
        It "<_> is valid Zip file" {
            InModuleScope PSTableauFiles -Parameters @{ file = $_ } {
                $file | Should -Not -BeNullOrEmpty
                Test-TableauZipFile -Path $file | Should -BeTrue
            }
        }
    }
    It "Test Zip File - invalid" {
        InModuleScope PSTableauFiles {
            Test-TableauZipFile "Tests/Files/invalid.zip.tdsx" | Should -BeFalse
        }
    }
}
