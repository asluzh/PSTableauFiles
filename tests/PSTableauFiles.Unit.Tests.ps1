BeforeDiscovery {
    $script:ConfigFiles = Get-ChildItem -Path "./tests/config" -Filter "test_*.json" | Resolve-Path -Relative
    $script:twbFiles  = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.twb  -Exclude invalid.* | Resolve-Path -Relative
    $script:twbxFiles = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.twbx -Exclude invalid.* | Resolve-Path -Relative
    $script:tdsFiles  = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.tds  -Exclude invalid.* | Resolve-Path -Relative
    $script:tdsxFiles = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.tdsx -Exclude invalid.* | Resolve-Path -Relative
}
BeforeAll {
    Import-Module ./PSTableauFiles -Force
    # Import-Module Assert
    # . ./Tests/Test.Functions.ps1
    # InModuleScope 'PSTableauFiles' { $script:VerbosePreference = 'Continue' } # display verbose output of module functions
    $script:VerbosePreference = 'Continue' # display verbose output of the tests
}

Describe "Unit Tests for Get-TableauFileXml" -Tag Unit {
    Context "Getting XML from .twb files" -ForEach $twbFiles {
        It "<_> content returned" {
            Get-TableauFileXml $_ | Should -BeOfType String
        }
    }
    Context "Getting XML from .twb files (pipeline)" -ForEach $twbFiles {
        It "<_> content returned" {
            $_ | Get-TableauFileXml | Should -BeOfType String
        }
    }
    Context "Getting XML from .twbx files" -ForEach $twbxFiles {
        It "<_> content returned" {
            Get-TableauFileXml $_ | Should -BeOfType String
        }
    }
    Context "Getting XML from .twbx files (pipeline)" -ForEach $twbxFiles {
        It "<_> content returned" {
            $_ | Get-TableauFileXml | Should -BeOfType String
        }
    }
    Context "Getting XML from .tds files" -ForEach $tdsFiles {
        It "<_> content returned" {
            Get-TableauFileXml -Path $_ | Should -BeOfType String
        }
    }
    Context "Getting XML from .tdsx files" -ForEach $tdsxFiles {
        It "<_> content returned" {
            Get-TableauFileXml -Path $_ | Should -BeOfType String
        }
    }
    Context "Exceptions" {
        It "Missing files should throw exception" {
            {Get-TableauFileXml "./tests/assets/missing.twbx"} | Should -Throw -ExpectedMessage "File not found*"
        }
        It "Unknown file types should throw exception" {
            {Get-TableauFileXml "./tests/assets/invalid.twby"} | Should -Throw -ExpectedMessage "Unknown file type*"
        }
        It "Invalid TWBX should throw exception" {
            {Get-TableauFileXml "./tests/assets/invalid.twbx"} | Should -Throw -ExpectedMessage "Main XML file not found*"
        }
        It "Invalid Zip file should throw exception" {
            $err = {Get-TableauFileXml "./tests/assets/invalid.zip.tdsx"} | Should -Throw -PassThru
            $err.Exception.InnerException.Message | Should -Be "End of Central Directory record could not be found."
        }
    }
    Context "Test Zip File .twbx" -ForEach $twbxFiles {
        It "<_> is valid Zip file" {
            InModuleScope PSTableauFiles -Parameters @{ file = $_ } {
                # $file | Should -Not -BeNullOrEmpty
                Test-TableauZipFile $file | Should -BeTrue
            }
        }
    }
    Context "Test Zip File .tsdx" -ForEach $tdsxFiles {
        It "<_> is valid Zip file" {
            InModuleScope PSTableauFiles -Parameters @{ file = $_ } {
                $file | Test-TableauZipFile | Should -BeTrue
            }
        }
    }
    It "Test Zip File - invalid" {
        InModuleScope PSTableauFiles {
            Test-TableauZipFile "./tests/assets/invalid.zip.tdsx" | Should -BeFalse
        }
    }
}
