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
                return (Get-Content $Path)
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
                    $props['Attributes'] = $_.Attributes.GetEnumerator() | ForEach-Object { $_ } # @{name=$_.name; value=$_.value}
                    # .get_attributes().GetEnumerator() - name,value
                }
                $props['XmlElement'] = $_
            }

            if ($paths) {
                $props['FileName'] = $paths | Select-Object -Index $i
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

            $worksheets = @()
            $xml | Select-Xml '/workbook/worksheets/worksheet' | Select-Object -ExpandProperty Node | ForEach-Object {
                $props = @{
                    Name = $_.Attributes['name'].Value;
                    DisplayName = if ($_.Attributes['caption']) { $_.Attributes['caption'].Value } else { $_.Attributes['name'].Value };
                    # TODO: Include list of referenced data sources.
                }
                $worksheets += New-Object PSObject -Property $props
            }

            $dashboards = @()
            $xml | Select-Xml '/workbook/dashboards/dashboard' | Select-Object -ExpandProperty Node | ForEach-Object {
                # TODO: This really slows down the whole cmdlet. Find a way to make the dashboards' Worksheets property lazy evaluated.
                $dashboardWorksheets = @()
                $_ | Select-Xml './zones//zone' | Select-Object -ExpandProperty Node | Where-Object { $null -eq $_.Attributes['type'] -and $null -ne $_.Attributes['name'] } | ForEach-Object {
                    # Assume any zone with a @name but not a @type is a worksheet zone.
                    $zone = $_
                    $dashboardWorksheets += ($worksheets | Where-Object { $_.Name -eq $zone.Attributes['name'].Value })
                }
                $props = @{
                    Name = $_.Attributes['name'].Value;
                    DisplayName = if ($_.Attributes['caption']) { $_.Attributes['caption'].Value } else { $_.Attributes['name'].Value };
                    Worksheets = $dashboardWorksheets;
                }
                $dashboards += New-Object PSObject -Property $props
            }

            $datasources = @()
            $parameters = @()
            $xml | Select-Xml '/workbook/datasources/datasource' | Select-Object -ExpandProperty Node | ForEach-Object {
                if ($_.name -eq 'Parameters' -and $_.hasconnection -eq 'false' -and $_.inline -eq 'true') {
                    # this special data source contains the definition of parameters
                    $_ | Select-Xml './column' | Select-Object -ExpandProperty Node | ForEach-Object {
                        $props = @{
                            Name = $_.Attributes['name'].Value; #-ireplace "^\[|\]$", "";
                            DisplayName = if ($_.Attributes['caption']) { $_.Attributes['caption'].Value } else { $_.Attributes['name'].Value };
                            DataType = $_.Attributes['datatype'].Value;
                            DomainType = $_.Attributes['param-domain-type'].Value;
                            Value = $_.Attributes['value'].Value;
                            ValueDisplayName = if ($_.Attributes['alias']) { $_.Attributes['alias'].Value } else { $_.Attributes['value'].Value };
                            # TODO: For "list" parameters, include a ValueList property with all of the values and aliases.
                        }
                        $parameters += New-Object PSObject -Property $props
                    }
                } else {
                    $columns = @()
                    $_ | Select-Xml './column' | Select-Object -ExpandProperty Node | ForEach-Object {
                        $column = @{
                            Name = $_.Attributes['name'].Value; #-ireplace "^\[|\]$", "";
                            DisplayName = if ($_.Attributes['caption']) { $_.Attributes['caption'].Value } else { $_.Attributes['name'].Value };
                            Role = $_.Attributes['role'].Value;
                            Type = $_.Attributes['type'].Value;
                            DataType = $_.Attributes['datatype'].Value;
                            Value = $_.Attributes['value'].Value;
                        }
                        if ($_.HasChildNodes) {
                            $column['Formula'] = ($_ | Select-Xml './calculation/@formula').Node.Value;
                        }
                        $columns += $column
                    }
                    $props = @{
                        Name = $_.Attributes['name'].Value;
                        DisplayName = if ($_.Attributes['caption']) { $_.Attributes['caption'].Value } else { $_.Attributes['name'].Value };
                        ConnectionType = ($_ | Select-Xml './connection/@class').Node.Value;
                        # TODO: Include a "Connection" PSObject property with properties specific to the type of connection (i.e. file path for CSVs and server for SQL Server, etc).
                        Columns = $columns;
                    }
                    $datasources += New-Object PSObject -Property $props
                }
            }

            $props = @{
                FileVersion = ($xml | Select-Xml '/workbook/@version').Node.Value;
                FileOriginalVersion = ($xml | Select-Xml '/workbook/@original-version').Node.Value;
                BuildVersion = ($xml | Select-Xml '//comment()[1]').Node.Value; #-ireplace "[^\d\.]", "";
                SourcePlatform = ($xml | Select-Xml '/workbook/@source-platform').Node.Value;
                SourceBuild = ($xml | Select-Xml '/workbook/@source-build').Node.Value;
                Worksheets = $worksheets;
                Dashboards = $dashboards;
                # Stories = $stories;
                Datasources = $datasources;
                Parameters = $parameters;
                # Actions = $actions;
                # Preferences = $preferences;
                # Styles = $styles;
                # Windows = $windows;
                DocumentXml = $xml;
            }

            if ($paths) {
                $props['FileName'] = $paths | Select-Object -Index $i
            }

            Write-Output (New-Object PSObject -Property $props)
            $i++
        }
    }
}
