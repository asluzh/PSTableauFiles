BeforeDiscovery {
    $script:ModuleName = 'PSTableauFiles'
    $script:ConfigFiles = Get-ChildItem -Path "./tests/config" -Filter "test_*.json" | Resolve-Path -Relative
    $script:twbFiles  = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.twb  -Exclude invalid.* | Resolve-Path -Relative
    $script:twbxFiles = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.twbx -Exclude invalid.* | Resolve-Path -Relative
    $script:tdsFiles  = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.tds  -Exclude invalid.* | Resolve-Path -Relative
    $script:tdsxFiles = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.tdsx -Exclude invalid.* | Resolve-Path -Relative
    $script:workbookFiles  = $script:twbFiles + $script:twbxFiles
}
BeforeAll {
    Get-Module -Name $ModuleName -All | Remove-Module -Force -ErrorAction Ignore
    Import-Module ./$ModuleName -Force
    # Requires -Modules Assert
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
}

Describe "Unit Tests for Test-TableauZipFile" -Tag Unit {
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
    Context "Exceptions" {
        It "Test Zip File - invalid" {
            InModuleScope PSTableauFiles {
                Test-TableauZipFile "./tests/assets/invalid.zip.tdsx" | Should -BeFalse
            }
        }
    }
}

Describe "Unit Tests for Get-TableauFileStructure" -Tag Unit {
    Context "Validate structure of workbook files" -Tag Struc -ForEach $workbookFiles {
        It "Workbook structure contains workbook element - <_>" {
            $result = Get-TableauFileStructure -Path $_ -XmlPath '/workbook'
            $result.Length | Should -Be 1
            $result.XmlElement | Should -BeOfType System.Xml.XmlElement
            $result.Elements | Should -BeNullOrEmpty
            $result.Attributes | Should -BeNullOrEmpty
            $xml = Get-TableauFileXml -Path $_
            $result2 = Get-TableauFileStructure -DocumentXml $xml
            $result2.Length | Should -Be 1
            $result2.XmlElement | Should -BeOfType System.Xml.XmlElement
            $result2.Elements | Should -BeNullOrEmpty
            $result2.Attributes | Should -BeNullOrEmpty
        }
        It "Workbook structure contains datasources element - <_>" {
            $result = Get-TableauFileStructure -Path $_ -XmlPath '/workbook/datasources'
            $result.Length | Should -Be 1
            $result.XmlElement | Should -BeOfType System.Xml.XmlElement
            $result.Elements | Should -BeNullOrEmpty
            $result.Attributes | Should -BeNullOrEmpty
        }
        It "Workbook structure (workbook) contains known elements and attributes - <_>" {
            $result = Get-TableauFileStructure -Path $_ -XmlPath '/workbook' -XmlElements -XmlAttributes
            $result.Length | Should -Be 1
            $result.Elements | Should -BeOfType System.Xml.XmlElement
            $result.Elements.Length | Should -BeGreaterThan 1
            $result.Elements | Select-Object -ExpandProperty 'Name' | Should -Contain 'datasources'
            $result.Elements | Select-Object -ExpandProperty 'Name' | Should -Contain 'worksheets'
            # $result.Elements | Select-Object -ExpandProperty 'Name' | Should -Contain 'repository-location' # only for workbooks that have been ever published
            $result.Elements | Select-Object -ExpandProperty 'Name' | Should -Contain 'preferences'
            $known_elem = @('document-format-change-manifest','repository-location','preferences','style',
                'datasources','datasource-relationships','mapsources','shared-views','actions',
                'worksheets','dashboards','windows','datagraph','external','thumbnails',
                'referenced-extensions','explain-data',
                '_.fcp.ExplainData_AuthorControls.true...explain-data',
                '_.fcp.WorkbookOptimizerRuleConfig.true...workbook-optimizer',
                '_.fcp.AnimationOnByDefault.false...style')
            $result.Elements | Select-Object -ExpandProperty 'Name' | Should -BeIn $known_elem
            $result.Attributes | Should -BeOfType System.Xml.XmlAttribute
            $result.Attributes.Length | Should -BeGreaterThan 1
            $result.Attributes | Select-Object -ExpandProperty 'Name' | Should -Contain 'source-build'
            # $result.Attributes | Select-Object -ExpandProperty 'Name' | Should -Contain 'source-platform'
            $result.Attributes | Select-Object -ExpandProperty 'Name' | Should -Contain 'version'
            $known_attr = @('xml:base','xmlns:user','source-build','source-platform','version','original-version','upgrade-extracts')
            $result.Attributes | Select-Object -ExpandProperty 'Name' | Should -BeIn $known_attr
        }
        It "Workbook structure (workbook/datasources) contains known elements and attributes - <_>" {
            $result = Get-TableauFileStructure -Path $_ -XmlPath '/workbook/datasources' -XmlElements
            $result.Elements.Length | Should -BeGreaterOrEqual 1
            $result.Elements | ForEach-Object { $_.get_name() } | Should -Contain 'datasource'
            $result = Get-TableauFileStructure -Path $_ -XmlPath '/workbook/datasources/datasource' -XmlElements -XmlAttributes
            $result.Elements.Length | Should -BeGreaterOrEqual 1
            $result.Elements | ForEach-Object { $_.get_name() } | Should -Contain 'connection'
            $result.Elements | ForEach-Object { $_.get_name() } | Should -Contain 'column'
            $known_elem = @('connection','column','aliases','column-instance','semantic-values','group','filter',
                'drill-paths','default-sorts','field-sort-info','folders-common','date-options',
                'layout','style','overridable-settings',
                'extract','datasource-dependencies','object-graph','repository-location',
                '_.fcp.ObjectModelTableType.true...column',
                '_.fcp.ObjectModelEncapsulateLegacy.true...object-graph')
            $result.Elements | ForEach-Object { $_.get_name() } | Should -BeIn $known_elem
            $result.Attributes.Length | Should -BeGreaterOrEqual 1
            $result.Attributes | Should -BeOfType System.Xml.XmlAttribute
            $result.Attributes | Select-Object -ExpandProperty 'Name' | Should -Contain 'name'
            # $result.Attributes | Select-Object -ExpandProperty 'Name' | Should -Contain 'inline'
            $result.Attributes | Select-Object -ExpandProperty 'Name' | Should -Contain 'version'
            # $result.Attributes | Select-Object -ExpandProperty 'Name' | Should -Contain 'caption' # not for all datasources
            $known_attr = @('name','version','caption','inline')
            $result.Attributes | Select-Object -ExpandProperty 'Name' | Should -BeIn $known_attr
        }
        It "Workbook structure (workbook/datasources/column) contains known elements and attributes - <_>" {
            $result = Get-TableauFileStructure -Path $_ -XmlPath '/workbook/datasources/datasource/column' -XmlAttributes
            $result.Attributes.Length | Should -BeGreaterThan 1
            $result.Attributes | Should -BeOfType System.Xml.XmlAttribute
            $result.Attributes | Select-Object -ExpandProperty 'Name' | Should -Contain 'name'
            $result.Attributes | Select-Object -ExpandProperty 'Name' | Should -Contain 'role'
            $result.Attributes | Select-Object -ExpandProperty 'Name' | Should -Contain 'type'
            $result.Attributes | Select-Object -ExpandProperty 'Name' | Should -Contain 'datatype'
            # $result.Attributes | Select-Object -ExpandProperty 'Name' | Should -Contain 'caption' # not for all columns
        }
    }
}

Describe "Unit Tests for Get-TableauFileObject" -Tag Unit {
    Context "Validate structure of workbook files" -Tag Struc -ForEach $workbookFiles {
        It "Check workbook object from .twb(x) file - <_>" {
            $result = Get-TableauFileObject -Path $_
            $result.Length | Should -Be 1
            $result.FileName | Should -Not -BeNullOrEmpty
            $result.FileVersion | Should -Not -BeNullOrEmpty
            $result.BuildVersion | Should -Not -BeNullOrEmpty
            $result.Worksheets | Should -Not -BeNullOrEmpty
            $result.Datasources | Should -Not -BeNullOrEmpty
            $result.DocumentXml | Should -BeOfType System.Xml.XmlDocument
        }
    }
}
