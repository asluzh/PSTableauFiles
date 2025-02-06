# System.IO.Compression.FileSystem requires at least .NET 4.5
# [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression") | Out-Null

function New-TableauZipFile {
<#
.SYNOPSIS
Create a Tableau packaged file

.DESCRIPTION
Creates a Tableau packaged file (workbook: .twbx, data source: .tdsx) with the provided XML.

.PARAMETER Path
The file path to the output Tableau packaged file

.PARAMETER DocumentXml
The workbook/datasource XML.

.EXAMPLE
New-TableauZipFile -Path workbook.twbx -DocumentXml $xml
#>
[CmdletBinding(SupportsShouldProcess)]
Param(
    [Parameter(Mandatory)] [string]$Path,
    [Parameter(Mandatory)] [xml]$DocumentXml
)
    $fileType = [System.IO.Path]::GetExtension($Path).Substring(1)
    $entryName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    if ($fileType -eq "twbx") {
        $tableauObject = "workbook"
        $entryName += '.twb'
    } elseif ($fileType -eq "tdsx") {
        $tableauObject = "data source"
        $entryName += '.tds'
    } else {
        throw [System.IO.FileFormatException] "Unknown file type. Tableau document file types are expected."
    }
    if ($PSCmdlet.ShouldProcess($Path, "New packaged $tableauObject")) {
        $fileStream = $null
        $zipArchive = $null
        $entryStream = $null
        try {
            $fileStream = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::ReadWrite)
            $zipArchive = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Update)
            $entry = $zipArchive.CreateEntry($entryName, ([System.IO.Compression.CompressionLevel]::Optimal))
            $entryStream = $entry.Open()
            $DocumentXml.Save($entryStream)
            $entryStream.Close()
            return $true
        } finally {
            if ($entryStream) {
                $entryStream.Dispose()
            }
            if ($zipArchive) {
                $zipArchive.Dispose()
            }
            if ($fileStream) {
                $fileStream.Dispose()
            }
        }
    }
}

function Update-TableauZipFile {
<#
.SYNOPSIS
Update Tableau File Contents

.DESCRIPTION
Updates the supplied workbook/datasource XML inside the compressed Tableau file.
or
Updates the original data file inside the compressed Tableau file.

.PARAMETER Path
The file path of the compressed Tableau file.

.PARAMETER DocumentXml
(Optional) The workbook/datasource XML for update.

.PARAMETER DataFile
(Optional) The file path of the data file(s).

.EXAMPLE
$result = Update-TableauZipFile -Path $twbxFile -DocumentXml $xml

.EXAMPLE
$result = Update-TableauZipFile -Path $twbxFile -DataFile "Employee.xlsx"
#>
[CmdletBinding(SupportsShouldProcess)]
Param(
    [Parameter(Mandatory)] [string]$Path,
    [Parameter()] [xml]$DocumentXml,
    [Parameter()] [string[]]$DataFile
)
    if (Test-Path $Path) { # -and (Test-TableauZipFile $Path)
        $fileItem = Get-Item $Path
        $fileType = $fileItem.Extension.Substring(1)
        if ($fileType -eq "twbx") {
            $tableauObject = "workbook"
            $tableauDocumentType = '.twb'
        } elseif ($fileType -eq "tdsx") {
            $tableauObject = "data source"
            $tableauDocumentType = '.tds'
        } else {
            throw [System.IO.FileFormatException] "Unknown file type. Tableau document file types are expected."
        }
        if ($DocumentXml -and $PSCmdlet.ShouldProcess($Path, "Update $tableauDocumentType in packaged $tableauObject")) {
            $fileStream = $null
            $zipArchive = $null
            $entryStream = $null
            try {
                $fileStream = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
                $zipArchive = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Update)
                # note: $_.FullName -eq $_.Name is only true for root-level entries
                $entries = $zipArchive.Entries | Where-Object { $_.FullName -eq $_.Name -and [System.IO.Path]::GetExtension($_.Name) -eq $tableauDocumentType }
                if ($entries.Count -gt 1) {
                    throw [System.IO.FileLoadException] "More than one main XML files found."
                }
                $entry = $entries | Select-Object -First 1
                if ($entry) {
                    $entryName = $entry.Name
                    $entry.Delete()
                    $entry = $zipArchive.CreateEntry($entryName, ([System.IO.Compression.CompressionLevel]::Optimal))
                    $entryStream = $entry.Open()
                    $DocumentXml.Save($entryStream)
                    $entryStream.Close()
                    return $true
                } else {
                    throw [System.IO.FileNotFoundException] "Main XML file not found."
                }
            } finally {
                if ($entryStream) {
                    $entryStream.Dispose()
                }
                if ($zipArchive) {
                    $zipArchive.Dispose()
                }
                if ($fileStream) {
                    $fileStream.Dispose()
                }
            }
        }
        if ($DataFile) {
            foreach ($file in $DataFile) {
                $fileName = [System.IO.Path]::GetFileName($file)
                if ($PSCmdlet.ShouldProcess($Path, "Update data file $fileName in packaged $tableauObject")) {
                    $fileStream = $null
                    $zipArchive = $null
                    $entryStream = $null
                    $dataFileStream = $null
                    try {
                        $fileStream = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
                        $zipArchive = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Update)
                        $entries = $zipArchive.Entries | Where-Object { $_.Name -eq $fileName }
                        if ($entries.Count -gt 1) {
                            # TODO allow to specify the file (or subfolder) in archive
                            throw [System.IO.FileLoadException] "Duplicate data files found."
                        }
                        $entry = $entries #| Select-Object -First 1
                        if ($entry) {
                            $entryName = $entry.FullName
                            $entry.Delete()
                            $entry = $zipArchive.CreateEntry($entryName, ([System.IO.Compression.CompressionLevel]::Optimal))
                            $dataFileStream = [System.IO.File]::OpenRead($file)
                            $entryStream = $entry.Open()
                            $dataFileStream.CopyTo($entryStream)
                            $entryStream.Close()
                            $dataFileStream.Close()
                            return $true
                        } else {
                            throw [System.IO.FileNotFoundException] "Data file not found."
                        }
                    } finally {
                        if ($dataFileStream) {
                            $dataFileStream.Dispose()
                        }
                        if ($entryStream) {
                            $entryStream.Dispose()
                        }
                        if ($zipArchive) {
                            $zipArchive.Dispose()
                        }
                        if ($fileStream) {
                            $fileStream.Dispose()
                        }
                    }
                }
            }
        }
        return $false
    } else {
        throw [System.IO.FileNotFoundException] "Tableau file not found."
    }
}

function Test-TableauZipFile {
<#
.SYNOPSIS
Test Tableau zip file

.DESCRIPTION
Tests the Tableau packaged (zip) file for the magic zip file header.
The functions returns true for genuine zip file or false otherwise.

.PARAMETER Path
The path for the Tableau workbook/datasource file.

.EXAMPLE
$twbxFile | Test-TableauZipFile

.EXAMPLE
$result = Test-TableauZipFile datasource.tdsx

.NOTES
Source http://stackoverflow.com/a/1887113/31308
#>
Param(
    [Parameter(Mandatory,ValueFromPipeline)] [string]$Path
)
    process {
        $fileStream = $null
        $byteReader = $null
        try {
            $fileItem = Get-Item $Path
            $fileStream = New-Object System.IO.FileStream($fileItem.FullName, [System.IO.FileMode]::Open)
            $byteReader = New-Object System.IO.BinaryReader($fileStream)
            $bytes = $byteReader.ReadBytes(4)
            if ($bytes.Length -eq 4) {
                if ($bytes[0] -eq 80 -and
                    $bytes[1] -eq 75 -and
                    $bytes[2] -eq 3 -and
                    $bytes[3] -eq 4)
                {
                    return $true;
                }
            }
        } finally {
            if ($byteReader) {
                $byteReader.Close()
            }
            if ($fileStream) {
                $fileStream.Close()
            }
        }
        return $false
    }
}

function Get-TableauFileXml {
<#
.SYNOPSIS
Get Tableau Document Xml

.DESCRIPTION
Returns the workbook/datasource XML from a TWB(X)/TDS(X) file.
If the file is not compressed, the original XML contents are returned.

.PARAMETER Path
The file path to the Tableau document.

.EXAMPLE
$xml = Get-TableauFileXml -Path $filePath
#>
Param(
    [Parameter(Mandatory,ValueFromPipeline)] [string]$Path
)
    process {
        if (Test-Path $Path) {
            $fileItem = Get-Item $Path
            $fileType = $fileItem.Extension.Substring(1)
            if ($fileType -eq "twb" -or $fileType -eq "tds") {
                return (Get-Content $Path -Raw)
            }
            elseif ($fileType -eq "twbx" -or $fileType -eq "tdsx") {
                $fileStream = $null
                $zipArchive = $null
                $reader = $null
                try {
                    $fileStream = New-Object System.IO.FileStream($fileItem.FullName, [System.IO.FileMode]::Open)
                    $zipArchive = New-Object System.IO.Compression.ZipArchive($fileStream)
                    # note: $_.FullName -eq $_.Name is only true for root-level entries
                    $xmlFiles = $zipArchive.Entries | Where-Object {
                            $_.FullName -eq $_.Name -and ([System.IO.Path]::GetExtension($_.Name) -in @(".twb",".tds"))
                        } | Select-Object -First 1
                    if ($null -eq $xmlFiles) {
                        throw [System.IO.FileNotFoundException] "Main XML file not found."
                    }
                    $reader = New-Object System.IO.StreamReader $xmlFiles[0].Open()
                    $xml = $reader.ReadToEnd()
                    return $xml
                } finally {
                    if ($reader) {
                        $reader.Dispose()
                    }
                    if ($zipArchive) {
                        $zipArchive.Dispose()
                    }
                    if ($fileStream) {
                        $fileStream.Dispose()
                    }
                }
            } else {
                throw [System.IO.FileFormatException] "Unknown file type. Tableau document file types are expected."
            }
        } else {
            throw [System.IO.FileNotFoundException] "File not found."
        }
    }
}

function ConvertTo-TableauColumnDisplayName {
Param(
    [Parameter(Mandatory,Position=0,ValueFromPipeline)]
    [string]$Name
)
    process {
        if ($Name -eq '[:Measure Names]') {
            'Measure Names'
        } else {
            $Name -ireplace "^\[|\]$", ""
        }
    }
}

function ConvertTo-TableauDisplayFormula {
Param(
    [Parameter(Mandatory)]
    [string]$Formula,
    [Parameter(Position=1)]
    [hashtable]$Mapping
)
    if ($Mapping) {
        # TODO regex support for nested square brackets in the names e.g. https://stackoverflow.com/questions/69373904/regex-with-nested-square-brackets
        # TODO regex support for comments - don't replace matches within the commented parts
        # PS 5.1: [regex]::Replace($Formula,'(\[\w\]\.\[\w\])', {param($match) $Mapping[$match.Value] })
        # PS 6+ : $Formula -replace '(\[\w+\])', { if ($Mapping.ContainsKey($_.Value)) {$Mapping[$_.Value]} else {$_.Value} }
        return [regex]::Replace($Formula, '((\[[^]]+\]\.\[[^]]+\])|(\[[^]]+\]))', {
            param($match)
            if ($Mapping.ContainsKey($match.Value)) {
                $Mapping[$match.Value]
            } else {
                $match.Value
            } })
    } else {
        return $Formula
    }
}

# function ConvertFrom-Xml {
# <#
# .SYNOPSIS
# Converts am XML element to PSObject representation
# .EXAMPLE
# $xml = ConvertTo-Xml (get-content 1.json | ConvertFrom-Json) -Depth 4 -NoTypeInformation -as String
# .EXAMPLE
# ConvertFrom-Xml ([xml]($xml)).Objects.Object | ConvertTo-Json
# #>
# Param(
#     [Parameter(Mandatory,Position=0,ValueFromPipeline)][ValidateNotNullOrEmpty()]
#     [System.Xml.XmlElement]$Element
# )
#     if ($Element.Property) {
#         $PSObject = New-Object PSObject
#         foreach ($Property in @($Element.Property)) {
#             Write-Verbose $Property
#             if ($Property.Property.Name -like 'Property') {
#                 $PSObject | Add-Member NoteProperty $Property.Name ($Property.Property | ForEach-Object { ConvertFrom-Xml $_ })
#             } else {
#                 if ($Property.'#text') {
#                     $PSObject | Add-Member NoteProperty $Property.Name $Property.'#text'
#                 } else {
#                     if ($Property.Name) {
#                         $PSObject | Add-Member NoteProperty $Property.Name (ConvertFrom-Xml $Property)
#                     }
#                 }
#             }
#         }
#         $PSObject
#     }
# }

function ConvertFrom-XmlAttr {
Param(
    [Parameter(Mandatory,Position=0,ValueFromPipeline)][ValidateNotNullOrEmpty()]
    [System.Xml.XmlElement]$Element
)
    process {
        $PSObject = New-Object PSObject
        foreach ($attr in @($Element.Attributes)) {
            # convert attribute name from kebab case to pascal case
            $pc_name = [regex]::replace($attr.Name.ToLower(), '(^|-)(.)', { $args[0].Groups[2].Value.ToUpper()})
            $PSObject | Add-Member NoteProperty -Name $pc_name -Value $attr.Value
        }
        $PSObject
    }
}

function Get-TableauParamObject {
Param(
    [Parameter(Mandatory,Position=0,ValueFromPipeline)][ValidateNotNullOrEmpty()]
    [System.Xml.XmlElement]$Element
)
    process {
        $props = @{
            Name = $Element.Attributes['name'].Value;
            DisplayName = if ($Element.Attributes['caption']) { $Element.Attributes['caption'].Value } else { $Element.Attributes['name'].Value | ConvertTo-TableauColumnDisplayName };
            DataType = $Element.Attributes['datatype'].Value;
            DomainType = $Element.Attributes['param-domain-type'].Value;
            Hidden = if ($Element.Attributes['hidden']) { $Element.Attributes['hidden'].Value -eq 'true' } else { $false };
            Value = $Element.Attributes['value'].Value;
        }
        if ($props['DomainType'] -eq 'range') {
            if ($Element.range) {
                $props['Range'] = $Element.range | ConvertFrom-XmlAttr
            }
        } elseif ($props['DomainType'] -eq 'list') {
            $param_value_list = $Element | Select-Xml './members/member' | Select-Object -ExpandProperty Node | ForEach-Object {
                $_ | ConvertFrom-XmlAttr
            }
            if ($param_value_list) {
                $props['ValueList'] = $param_value_list
                # Write-Debug ($param_value_list | ConvertTo-Json)
            }
            # the "aliases" element is ignored, because the information is already contained in "members"
        }
        if ($Element.Attributes['_.fcp.ParameterDefaultValues.true...default-value-field']) {
            $props['DevaultValueField'] = $Element.Attributes['_.fcp.ParameterDefaultValues.true...default-value-field'].Value
        }
        if ($Element.Attributes['_.fcp.ParameterDefaultValues.true...source-field']) {
            $props['ValueListSourceField'] = $Element.Attributes['_.fcp.ParameterDefaultValues.true...source-field'].Value
        }
        $props
    }
}

function Get-TableauFileStructure {
<#
.SYNOPSIS
Get Tableau Document properties from an XML file

.DESCRIPTION
Returns Tableau Document metadata/properties for a Tableau document.
Either DocumentXml, Path, or LiteralPath should be provided.

.PARAMETER DocumentXml
(Optional) The Tableau workbook/datasource XML.

.PARAMETER Path
(Optional) The path for the Tableau workbook/datasource file.

.PARAMETER LiteralPath
(Optional) The literal path for the Tableau workbook/datasource file.

.PARAMETER XmlPath
(Optional) The XPath specification for the element to be selected.

.PARAMETER XmlAttributes
(Optional) If provided, the output will include the list of attributes in the selected element.

.PARAMETER XmlElements
(Optional) If provided, the output will include the list of sub-elements in the selected element.

.EXAMPLE
$result = Get-TableauFileStructure -Path $filename

.EXAMPLE
$result = Get-TableauFileStructure -DocumentXml $xml
#>
Param(
    [Parameter(Mandatory,ParameterSetName='Xml',Position=0,ValueFromPipeline)][ValidateNotNullOrEmpty()]
    [xml[]]$DocumentXml,

    [Parameter(Mandatory,ParameterSetName='Path',Position=0,ValueFromPipeline,ValueFromPipelineByPropertyName)][ValidateNotNullOrEmpty()]
    [string[]]$Path,

    [Parameter(Mandatory,ParameterSetName='LiteralPath',ValueFromPipeline,ValueFromPipelineByPropertyName)][Alias('FullName')][ValidateNotNullOrEmpty()]
    [string[]]$LiteralPath,

    [Parameter()]
    [string]$XmlPath='/workbook',
    [Parameter()]
    [switch]$XmlAttributes,
    [Parameter()]
    [switch]$XmlElements
)
    process {
        $needXml = $false
        if ($PSCmdlet.ParameterSetName -eq "Path") {
            $paths = Resolve-Path -Path $Path | Select-Object -ExpandProperty Path
            $needXml = $true
        }
        elseif ($PSCmdlet.ParameterSetName -eq "LiteralPath") {
            $paths = Resolve-Path -LiteralPath $LiteralPath | Select-Object -ExpandProperty Path
            $needXml = $true
        }

        if ($needXml) {
            $DocumentXml = $paths | ForEach-Object { Get-TableauFileXml $_ }
        }

        $i = 0
        foreach ($xml in $DocumentXml) {
            $props = @{
                FileVersion = ($xml | Select-Xml '/workbook/@version').Node.Value;
                FileOriginalVersion = ($xml | Select-Xml '/workbook/@original-version').Node.Value;
                BuildVersion = ($xml | Select-Xml '//comment()[1]').Node.Value;
                SourcePlatform = ($xml | Select-Xml '/workbook/@source-platform').Node.Value;
                SourceBuild = ($xml | Select-Xml '/workbook/@source-build').Node.Value;
            }
            $xml | Select-Xml $XmlPath | Select-Object -ExpandProperty Node | ForEach-Object {
                if ($XmlElements) {
                    $props['Elements'] = $_.GetEnumerator() | ForEach-Object { $_ }
                    # .get_name()
                }
                if ($XmlAttributes) {
                    $props['Attributes'] = $_.Attributes.GetEnumerator() | Write-Output # ForEach-Object { @{name=$_.name; value=$_.value} }
                    # .get_attributes().GetEnumerator() - name,value
                }
                $props['XmlElement'] = $_
            }

            if ($paths) {
                $props['FilePath'] = $paths | Select-Object -Index $i
                $props['FileName'] = [System.IO.Path]::GetFileName($props['FilePath'])
            }

            Write-Output (New-Object PSObject -Property $props)
            $i++
        }
    }
}

function Get-TableauFileObject {
<#
.SYNOPSIS
Get Tableau Document Object created from an XML file

.DESCRIPTION
Returns Tableau Document metadata information for a Tableau workbook(s) and/or data source(s).
Either DocumentXml, Path, or LiteralPath should be provided.s

.PARAMETER DocumentXml
(Optional) The Tableau workbook/datasource XML.

.PARAMETER Path
(Optional) The path for the Tableau workbook/datasource file.

.PARAMETER LiteralPath
(Optional) The literal path for the Tableau workbook/datasource file.

.EXAMPLE
$result = Get-TableauFileObject -DocumentXml $xml

.NOTES
Inspired by https://joshua.poehls.me/2013/tableaukit-a-powershell-module-for-tableau/
#>
Param(
    [Parameter(Mandatory,ParameterSetName='Xml',Position=0,ValueFromPipeline)][ValidateNotNullOrEmpty()]
    [xml[]]$DocumentXml,

    [Parameter(Mandatory,ParameterSetName='Path',Position=0,ValueFromPipeline,ValueFromPipelineByPropertyName)][ValidateNotNullOrEmpty()]
    [string[]]$Path,

    [Parameter(Mandatory,ParameterSetName='LiteralPath',ValueFromPipeline,ValueFromPipelineByPropertyName)][Alias('FullName')][ValidateNotNullOrEmpty()]
    [string[]]$LiteralPath
)
    process {
        $needXml = $false
        if ($PSCmdlet.ParameterSetName -eq "Path") {
            $paths = Resolve-Path -Path $Path | Select-Object -ExpandProperty Path
            $needXml = $true
        }
        elseif ($PSCmdlet.ParameterSetName -eq "LiteralPath") {
            $paths = Resolve-Path -LiteralPath $LiteralPath | Select-Object -ExpandProperty Path
            $needXml = $true
        }

        if ($needXml) {
            $DocumentXml = $paths | ForEach-Object { Get-TableauFileXml $_ }
        }

        $i = 0
        foreach ($xml in $DocumentXml) {

            $doc_type = $xml.DocumentElement.LocalName
            if ($doc_type -eq 'workbook') {
                $repository_location = $xml | Select-Xml '/workbook/repository-location' | Select-Object -ExpandProperty Node
                $datasources_xml = $xml | Select-Xml '/workbook/datasources/datasource' | Select-Object -ExpandProperty Node
            } elseif ($doc_type -eq 'datasource') {
                $repository_location = $xml | Select-Xml '/datasource/repository-location' | Select-Object -ExpandProperty Node
                $datasources_xml = $xml | Select-Xml '/datasource' | Select-Object -ExpandProperty Node
            }

            $parameters = @()
            $datasources = $datasources_xml | ForEach-Object {
                $datasource_name = $_.name
                if ($datasource_name -eq 'Parameters' -and $_.hasconnection -eq 'false' -and $_.inline -eq 'true') {
                    # this special data source contains the definition of parameters
                    $_ | Select-Xml './column' | Select-Object -ExpandProperty Node | ForEach-Object {
                        $parameters += New-Object PSObject -Property ($_ | Get-TableauParamObject)
                    }
                } else {
                    # Following sub-elements are present under datasource:
                    # connection > named-connections, relation, cols>map, metadata-records>metadata-record
                    # column-instance
                    # drill-paths > drill-path > field
                    # folders-common > folder > folder-item
                    # extract > connection
                    # layout
                    # style > style-rule > encoding
                    # semantic-values
                    # datasource-dependencies (for params dependency)
                    # object-graph > objects > object
                    # object-graph > relationships > relationship
                    $columns = $_ | Select-Xml './column' | Select-Object -ExpandProperty Node | ForEach-Object {
                        $column = @{
                            Name = $_.Attributes['name'].Value;
                            DisplayName = if ($_.Attributes['caption']) { $_.Attributes['caption'].Value } else { $_.Attributes['name'].Value | ConvertTo-TableauColumnDisplayName };
                            Role = $_.Attributes['role'].Value;
                            Type = $_.Attributes['type'].Value;
                            DataType = $_.Attributes['datatype'].Value;
                            Hidden = if ($_.Attributes['hidden']) { $_.Attributes['hidden'].Value -eq 'true' } else { $false };
                        }
                        if ($_.Attributes['datatype'].Value -eq 'table') {
                            # it appears Tableau Desktop always adds "(Count)" for record counters
                            $column['DisplayName'] = $column['DisplayName'] + ' (Count)'
                        }
                        if ($_.HasChildNodes) {
                            $column['Formula'] = ($_ | Select-Xml './calculation/@formula').Node.Value;
                        }
                        $column
                    }
                    $cols = $_ | Select-Xml './connection/cols/map' | Select-Object -ExpandProperty Node | ForEach-Object {
                        $_ | ConvertFrom-XmlAttr
                    }
                    $metadata_records = $_ | Select-Xml './connection/metadata-records/metadata-record' | Select-Object -ExpandProperty Node | ForEach-Object {
                        $_ | ConvertFrom-XmlAttr
                    }
                    # complement missing columns from worksheet metadata - only for "cols"
                    $cols | ForEach-Object {
                        $column_name = $_.Key
                        $column_found = $columns | Where-Object Name -eq $column_name
                        if (-Not $column_found) {
                            Write-Debug ("New column found: {0}" -f $column_name)
                            $xml | Select-Xml '/workbook/worksheets/worksheet/table/view/datasource-dependencies' | Select-Object -ExpandProperty Node | Where-Object datasource -eq $datasource_name | ForEach-Object {
                                $_ws_column = $_ | Select-Xml './column' | Select-Object -ExpandProperty Node | Where-Object name -eq $column_name | Select-Object -First 1
                                if ($_ws_column) {
                                    $column = @{
                                        Name = $_ws_column.Attributes['name'].Value;
                                        DisplayName = if ($_ws_column.Attributes['caption']) { $_ws_column.Attributes['caption'].Value } else { $_ws_column.Attributes['name'].Value | ConvertTo-TableauColumnDisplayName };
                                        Role = $_ws_column.Attributes['role'].Value;
                                        Type = $_ws_column.Attributes['type'].Value;
                                        DataType = $_ws_column.Attributes['datatype'].Value;
                                        Hidden = if ($_ws_column.Attributes['hidden']) { $_ws_column.Attributes['hidden'].Value -eq 'true' } else { $false };
                                    }
                                    # if ($_ws_column.Attributes['datatype'].Value -eq 'table') {
                                    #     $column['DisplayName'] = $column['DisplayName'] + ' (Count)'
                                    # }
                                    # if ($_ws_column.HasChildNodes) {
                                    #     $column['Formula'] = ($_ws_column | Select-Xml './calculation/@formula').Node.Value;
                                    # }
                                    $columns += $column
                                } else {
                                    Write-Debug "WS column definition not found"
                                }
                            }
                        }
                    }
                    $folders = $_ | Select-Xml './folders-common/folder' | Select-Object -ExpandProperty Node | ForEach-Object {
                        @{
                            Name = $_.name;
                            FolderItems = $_.'folder-item' | ForEach-Object {
                                $_ | ConvertFrom-XmlAttr;
                            }
                        }
                    }
                    $drillpaths = $_ | Select-Xml './drill-paths/drill-path' | Select-Object -ExpandProperty Node | ForEach-Object {
                        @{
                            Name = $_.name;
                            HierarchyItems = $_.field; #| ForEach-Object { $_ };
                        }
                    }
                    $encodings = $_ | Select-Xml './style/style-rule/encoding' | Select-Object -ExpandProperty Node | ForEach-Object {
                        $encoding = $_ | ConvertFrom-XmlAttr
                        if ($_.map) {
                            $encoding_mapping = $_ | Select-Xml './map' | Select-Object -ExpandProperty Node | ForEach-Object {
                                @{
                                    From = $_.bucket;
                                    To = $_.to;
                                }
                            }
                            $encoding | Add-Member NoteProperty -Name Mapping -Value $encoding_mapping
                        }
                        $encoding
                    }
                    $parameter_dependencies = $_ | Select-Xml './datasource-dependencies' | Select-Object -ExpandProperty Node | ForEach-Object {
                        if ($_.datasource -eq 'Parameters') {
                            $_ | Select-Xml './column' | Select-Object -ExpandProperty Node | ForEach-Object {
                                $_ | Get-TableauParamObject
                            }
                        } else {
                            # TODO
                            # $dependency = $_ | ConvertFrom-XmlAttr
                            # $datasource_dependencies += $dependency
                        }
                    }
                    $props = @{
                        Name = $_.name;
                        DisplayName = if ($_.caption) { $_.caption } else { $_.name };
                        ConnectionType = ($_ | Select-Xml './connection/@class').Node.Value;
                        # TODO: what is the outcome when the data sources has multiple connections?
                        # TODO: Include a "Connection" PSObject property with properties specific to the type of connection (i.e. file path for CSVs and server for SQL Server, etc).
                        ConnectionCols = $cols;
                        Columns = $columns;
                        Metadata = $metadata_records;
                        Folders = $folders;
                        Hierarchies = $drillpaths;
                        Encodings = $encodings;
                        ParameterDependencies = $parameter_dependencies;
                    }
                    $props
                }
            }

            # translate formula references into display names
            $ref_mapping = @{}
            foreach ($datasource in $datasources) {
                foreach ($column in $datasource.Columns) {
                    $ref_mapping[$column.Name] = '[' + $column.DisplayName + ']'
                    $ref_mapping['[' + $datasource.Name + '].' + $column.Name] = '[' + $datasource.DisplayName + '].[' + $column.DisplayName + ']'
                }
            }
            foreach ($parameter in $parameters) {
                # $ref_mapping[$parameter.Name] = '[' + $parameter.DisplayName + ']' # referencing by name without "[Parameters]." is not possible
                $ref_mapping['[Parameters].' + $parameter.Name] = '[Parameters].[' + $parameter.DisplayName + ']'
            }
            # $ref_mapping.GetEnumerator() | ForEach-Object {
            #     Write-Debug ($_.Key + '   ...   ' + $_.Value)
            # }
            foreach ($datasource in $datasources) {
                foreach ($column in $datasource.Columns) {
                    if ($column['Formula']) {
                        $column['DisplayFormula'] = ConvertTo-TableauDisplayFormula -Formula $column['Formula'] -Mapping $ref_mapping
                    }
                }
            }

            $props = @{
                FileVersion = ($xml | Select-Xml "/$doc_type/@version").Node.Value;
                FileOriginalVersion = ($xml | Select-Xml "/$doc_type/@original-version").Node.Value;
                BuildVersion = ($xml | Select-Xml '//comment()[1]').Node.Value;
                SourcePlatform = ($xml | Select-Xml "/$doc_type/@source-platform").Node.Value;
                RepositoryLocation = @{
                    DerivedFrom=($repository_location.Attributes | Where-Object Name -eq 'derived-from').Value;
                    Id=($repository_location.Attributes | Where-Object Name -eq 'id').Value;
                    Path=($repository_location.Attributes | Where-Object Name -eq 'path').Value;
                    Revision=($repository_location.Attributes | Where-Object Name -eq 'revision').Value;
                    Site=($repository_location.Attributes | Where-Object Name -eq 'site').Value;
                };
                DocumentXml = $xml;
            }

            if ($doc_type -eq 'workbook') {
                $props['SourceBuild'] = ($xml | Select-Xml '/workbook/@source-build').Node.Value;
                $props['Datasources'] = $datasources;
                $props['Parameters'] = $parameters;

                $preferences = $xml | Select-Xml '/workbook/preferences/preference' | Select-Object -ExpandProperty Node | ForEach-Object {
                    $_ | ConvertFrom-XmlAttr
                }
                $props['Preferences'] = $preferences

                $color_palettes = $xml | Select-Xml '/workbook/preferences/color-palette' | Select-Object -ExpandProperty Node | ForEach-Object {
                    $color_palette = $_ | ConvertFrom-XmlAttr
                    $color_palette | Add-Member NoteProperty -Name Colors -Value $_.color
                    $color_palette
                }
                $props['ColorPalettes'] = $color_palettes;

                $styles = $xml | Select-Xml '/workbook/style/style-rule' | Select-Object -ExpandProperty Node | ForEach-Object {
                    $style_rule = $_ | ConvertFrom-XmlAttr
                    if ($_.format) {
                        $style_rule | Add-Member NoteProperty -Name Format -Value ($_.format | ConvertFrom-XmlAttr)
                    }
                    $style_rule
                }
                $props['Styles'] = $styles;

                $worksheets = $xml | Select-Xml '/workbook/worksheets/worksheet' | Select-Object -ExpandProperty Node | ForEach-Object {
                    @{
                        Name = $_.Attributes['name'].Value;
                        DisplayName = if ($_.Attributes['caption']) { $_.Attributes['caption'].Value } else { $_.Attributes['name'].Value };
                        # TODO: Include list of referenced data sources.
                    }
                }
                $props['Worksheets'] = $worksheets;

                $dashboards = $xml | Select-Xml '/workbook/dashboards/dashboard' | Select-Object -ExpandProperty Node | ForEach-Object {
                    # TODO: This really slows down the whole cmdlet. Find a way to make the dashboards' Worksheets property lazy evaluated.
                    $dashboardWorksheets = $_ | Select-Xml './zones//zone' | Select-Object -ExpandProperty Node | Where-Object { $null -eq $_.Attributes['type'] -and $null -ne $_.Attributes['name'] } | ForEach-Object {
                        # Assume any zone with a @name but not a @type is a worksheet zone.
                        $zone = $_
                        $worksheets | Where-Object { $_.Name -eq $zone.Attributes['name'].Value }
                    }
                    @{
                        Name = $_.Attributes['name'].Value;
                        DisplayName = if ($_.Attributes['caption']) { $_.Attributes['caption'].Value } else { $_.Attributes['name'].Value };
                        Worksheets = $dashboardWorksheets;
                    }
                }
                $props['Dashboards'] = $dashboards;

                # $props['Stories'] = $stories;
                # $props['Actions'] = $actions;
                # $props['Windows'] = $windows;
                # $props['DataGraph'] = $datagraph;
                $shapes = $xml | Select-Xml '/workbook/external/shapes/shape' | Select-Object -ExpandProperty Node | ForEach-Object {
                    $shape = $_ | ConvertFrom-XmlAttr
                    $shape | Add-Member NoteProperty -Name Image -Value ([System.Convert]::FromBase64String($_.InnerText))
                    $shape
                }
                $props['Shapes'] = $shapes;

                $thumbnails = $xml | Select-Xml '/workbook/thumbnails/thumbnail' | Select-Object -ExpandProperty Node | ForEach-Object {
                    $thumbnail = $_ | ConvertFrom-XmlAttr
                    $thumbnail | Add-Member NoteProperty -Name Image -Value ([System.Convert]::FromBase64String($_.InnerText))
                    $thumbnail
                }
                $props['Thumbnails'] = $thumbnails;

                # /workbook/referenced-extensions
                $referenced_extensions = $xml | Select-Xml '/workbook/referenced-extensions/referenced-extension | /workbook/referenced-extensions/_.fcp.VizExtensions.true...referenced-extension' | Select-Object -ExpandProperty Node | ForEach-Object {
                    $dashboard_extension = $_ | Select-Xml './manifest/dashboard-extension' | Select-Object -ExpandProperty Node
                    $worksheet_extension = $_ | Select-Xml './manifest/worksheet-extension' | Select-Object -ExpandProperty Node
                    if ($_.Name -eq 'referenced-extension') {
                        $extension = $dashboard_extension
                    } else {
                        $extension = $worksheet_extension
                    }
                    $referenced_extension = New-Object PSObject -Property @{
                        Type = if ($worksheet_extension) { 'worksheet-extension' } else { 'dashboard-extension' };
                        ManifestVersion = ($_ | Select-Xml './manifest/@manifest-version').Node.Value;
                        ExtensionVersion = $extension.Attributes['extension-version'].Value;
                        ExtensionId = $extension.Attributes['id'].Value;
                        SourceLocation = ($extension | Select-Xml './source-location/url').Node.InnerText;
                        Name = ($extension | Select-Xml './name').Node.InnerText;
                        Description = ($extension | Select-Xml './description').Node.InnerText; #$extension.description;
                        Author = $extension.author | ConvertFrom-XmlAttr;
                        MinApiVersion = ($extension | Select-Xml './min-api-version').Node.InnerText;
                        DefaultLocale = ($extension | Select-Xml './default-locale').Node.InnerText;
                        ContextMenu = if ($extension | Select-Xml './context-menu/configure-context-menu-item') { $true } else { $false };
                        Icon = [System.Convert]::FromBase64String($extension.icon);
                        # Resources = '';
                        # ReferencedViews = '';
                    }
                    # Encoding for VizExtensions
                    if ($worksheet_extension) {
                        $encodings = $worksheet_extension | Select-Xml './encoding' | Select-Object -ExpandProperty Node | ForEach-Object {
                            @{
                                Id = $_.Attributes['id'].Value;
                                DisplayName = ($_ | Select-Xml './display-name').Node.InnerText;
                                RoleSpec = ($_ | Select-Xml './role-spec/role-type').Node.InnerText;
                                FieldsMaxCount = ($_ | Select-Xml './fields/@max-count').Node.Value;
                            }
                        }
                        $referenced_extension | Add-Member NoteProperty -Name Encoding -Value $encodings
                    }
                    # Write-Debug ($referenced_extension | ConvertTo-Json)
                    $referenced_extension
                }
                $props['ReferencedExtensions'] = $referenced_extensions;

                # explain-data
                # <_.fcp.ExplainData_AuthorControls.true...explain-data>

            } elseif ($doc_type -eq 'datasource') {
                $props['Datasources'] = $datasources;
            }

            if ($paths) {
                $props['FilePath'] = $paths | Select-Object -Index $i
                $props['FileName'] = [System.IO.Path]::GetFileName($props['FilePath'])
            }

            Write-Output (New-Object PSObject -Property $props)
            $i++
        }
    }
}
