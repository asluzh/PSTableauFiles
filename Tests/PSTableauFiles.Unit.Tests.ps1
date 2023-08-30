BeforeDiscovery {
    Get-Module PSTableauFiles | Remove-Module -Force
    Import-Module ./PSTableauFiles.psm1 -Force
}

Describe "Unit Tests for Get-TableauDocumentXml" -Tag Unit {
    InModuleScope PSTableauFiles {
        $twbFiles  = Get-ChildItem -Recurse -Path "Tests/Files" -Filter *.twb  -Exclude invalid.*
        $twbxFiles = Get-ChildItem -Recurse -Path "Tests/Files" -Filter *.twbx -Exclude invalid.*
        $tdsFiles  = Get-ChildItem -Recurse -Path "Tests/Files" -Filter *.tds  -Exclude invalid.*
        $tdsxFiles = Get-ChildItem -Recurse -Path "Tests/Files" -Filter *.tdsx -Exclude invalid.*
        Context "Getting XML from .twb files" -ForEach $twbFiles {
            It "<_> content returned" {
                Get-TableauDocumentXml $_ | Should -Not -BeNullOrEmpty
            }
        }
        Context "Getting XML from .twbx files" -ForEach $twbxFiles {
            It "<_> content returned" {
                Get-TableauDocumentXml $_ | Should -Not -BeNullOrEmpty
            }
        }
        Context "Getting XML from .tds files" -ForEach $tdsFiles {
            It "<_> content returned" {
                Get-TableauDocumentXml $_ | Should -Not -BeNullOrEmpty
            }
        }
        Context "Getting XML from .tdsx files" -ForEach $tdsxFiles {
            It "<_> content returned" {
                Get-TableauDocumentXml $_ | Should -Not -BeNullOrEmpty
            }
        }
        Context "Exceptions" {
            It "Unknown file types should throw exception" {
                {Get-TableauDocumentXml "Tests/Files/test_file.twby"} | Should -Throw -ExpectedMessage "Unknown file type*"
            }
            It "Missing files should throw exception" {
                {Get-TableauDocumentXml "Tests/Files/missing.twbx"} | Should -Throw -ExpectedMessage "File not found*"
            }
            It "Invalid TWBX should throw exception" {
                {Get-TableauDocumentXml "Tests/Files/invalid.twbx"} | Should -Throw -ExpectedMessage "Main XML file not found*"
            }
            It "Invalid Zip file should throw exception" {
                $err = {Get-TableauDocumentXml "Tests/Files/invalid.zip.tdsx"} | Should -Throw -PassThru
                $err.Exception.InnerException.Message | Should -Be "End of Central Directory record could not be found."
            }
        }
    }
}
