# original script here https://joshua.poehls.me/2013/tableaukit-a-powershell-module-for-tableau/

function Get-TableauDocumentXml {
<#
.SYNOPSIS
Get Tableau Document Xml

.DESCRIPTION
Returns the workbook/datasource XML from a TWB(X)/TDS(X) file.

.PARAMETER Path
The filename including pathname to the Tableau document.

.NOTES
If the file is not compressed, the original contents are returned.
#>
Param(
    [Parameter(Mandatory,ValueFromPipeline)] [string]$Path
)
    begin {
        $originalCurrentDirectory = [System.Environment]::CurrentDirectory

        # System.IO.Compression.FileSystem requires at least .NET 4.5
        [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression") | Out-Null
    }

    process {
        [System.Environment]::CurrentDirectory = (Get-Location).Path
        $extension = [System.IO.Path]::GetExtension($Path)
        if ($extension -eq ".twb" -or $extension -eq ".tds") {
            return (Get-Content -LiteralPath $Path)
        }
        elseif ($extension -eq ".twbx" -or $extension -eq ".tdsx") {
            $archiveStream = $null
            $archive = $null
            $reader = $null

            if (Test-Path $Path) {
                try {
                    $archiveStream = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open)
                    $archive = New-Object System.IO.Compression.ZipArchive($archiveStream)
                    $xmlFiles = ($archive.Entries | Where-Object { $_.FullName -eq $_.Name -and ([System.IO.Path]::GetExtension($_.Name) -eq ".twb" -or [System.IO.Path]::GetExtension($_.Name) -eq ".tds") })
                    if ($null -eq $xmlFiles) {
                        throw "Main XML file not found."
                    }
                    $reader = New-Object System.IO.StreamReader $xmlFiles[0].Open()
                    $xml = $reader.ReadToEnd()
                    return $xml
                } finally {
                    if ($reader) {
                        $reader.Dispose()
                    }
                    if ($archive) {
                        $archive.Dispose()
                    }
                    if ($archiveStream) {
                        $archiveStream.Dispose()
                    }
                }
            } else {
                throw "File not found."
            }
        }
        else {
            throw "Unknown file type. Tableau document file types are expected."
        }
    }

    end {
        [System.Environment]::CurrentDirectory = $originalCurrentDirectory
    }
}

function Update-TableauDocumentFromXml {
<#
.SYNOPSIS
Update Tableau Document File XML

.DESCRIPTION
Inserts the workbook XML into a TWBX file.
or
Inserts the datasource XML into a TDSX file.

.PARAMETER Path
The literal file path to export to.

.PARAMETER DocumentXml
The workbook XML to export.

.PARAMETER Update
Whether to update the TWB inside the destination TWBX file
if the destination file exists.

.PARAMETER Force
Whether to overwrite the destination TWBX file if it exists.
By default, you will be prompted whether to overwrite any
existing file.
#>
[CmdletBinding(SupportsShouldProcess)]
Param(
    [Parameter(Mandatory,Position=0)] [string]$Path,
    [Parameter(Mandatory,Position=1,ValueFromPipeline)] [xml]$DocumentXml,
    [Parameter()] [switch]$Update,
    [Parameter()] [switch]$Force
)
    begin {
        $originalCurrentDirectory = [System.Environment]::CurrentDirectory
        # System.IO.Compression.FileSystem requires at least .NET 4.5
        [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression") | Out-Null
    }

    process {
        [System.Environment]::CurrentDirectory = (Get-Location).Path
        $entryName = [System.IO.Path]::GetFileNameWithoutExtension($Path) + '.twb'
        $createNewTwbx = $false

        if (Test-Path $Path) {
            if ($Update -or $Force -or $PSCmdlet.ShouldContinue('Overwrite existing file?', 'Confirm')) {
                if ($Update) {
                    if ($PSCmdlet.ShouldProcess($Path, 'Update TWB in packaged workbook')) {

                        [System.IO.FileStream]$fileStream = $null
                        [System.IO.Compression.ZipArchive]$zip = $null
                        try {
                            $fileStream = New-Object System.IO.FileStream -ArgumentList $Path, ([System.IO.FileMode]::Open), ([System.IO.FileAccess]::ReadWrite), ([System.IO.FileShare]::Read)
                            $zip = New-Object System.IO.Compression.ZipArchive -ArgumentList $fileStream, ([System.IO.Compression.ZipArchiveMode]::Update)

                            # Locate the existing TWB entry and remove it.
                            $entry = $zip.Entries |
                                Where-Object {
                                    # Look for a .twb file at the root level of the archive.
                                    $_.FullName -eq $_.Name -and ([System.IO.Path]::GetExtension($_.Name)) -eq '.twb'
                                } |
                                Select-Object -First 1
                            if ($entry) {
                                $entry.Delete()
                            }

                            $entry = $zip.CreateEntry($entryName, ([System.IO.Compression.CompressionLevel]::Optimal))
                            [System.IO.Stream]$entryStream = $null
                            try {
                                $entryStream = $entry.Open()
                                $DocumentXml.Save($entryStream)
                            }
                            finally {
                                if ($entryStream) {
                                    $entryStream.Dispose()
                                }
                            }
                        }
                        finally {
                            if ($zip) {
                                $zip.Dispose()
                            }
                            if ($fileStream) {
                                $fileStream.Dispose()
                            }
                        }
                    }
                }
                else {
                    if ($PSCmdlet.ShouldProcess($Path, 'Replace existing packaged workbook')) {
                        # delete existing TWBX
                        Remove-Item $Path -ErrorAction Stop #TODO: Figure out how to pass WhatIf and Confirm to this
                        $createNewTwbx = $true
                    }
                }
            }
        }
        else {
            if ($PSCmdlet.ShouldProcess($Path, 'Export packaged workbook')) {
                $createNewTwbx = $true
            }
        }

        if ($createNewTwbx) {
            [System.IO.FileStream]$fileStream = $null
            [System.IO.Compression.ZipArchive]$zip = $null
            try {
                $fileStream = New-Object System.IO.FileStream -ArgumentList $Path, ([System.IO.FileMode]::CreateNew), ([System.IO.FileAccess]::ReadWrite), ([System.IO.FileShare]::None)
                $zip = New-Object System.IO.Compression.ZipArchive -ArgumentList $fileStream, ([System.IO.Compression.ZipArchiveMode]::Update)

                $entry = $zip.CreateEntry($entryName, ([System.IO.Compression.CompressionLevel]::Optimal))
                [System.IO.Stream]$entryStream = $null
                try {
                    $entryStream = $entry.Open()
                    $DocumentXml.Save($entryStream)
                }
                finally {
                    if ($entryStream) {
                        $entryStream.Dispose()
                    }
                }
            }
            finally {
                if ($zip) {
                    $zip.Dispose()
                }
                if ($fileStream) {
                    $fileStream.Dispose()
                }
            }
        }
    }

    end {
        [System.Environment]::CurrentDirectory = $originalCurrentDirectory
    }
}

function Get-TableauDocumentObject {
<#
.SYNOPSIS
Get Tableau Document Object

.DESCRIPTION
Returns metadata information for local workbook(s).
#>
Param(
    [Parameter(
        Mandatory = $true,
        ParameterSetName = "Xml",
        Position = 0,
        ValueFromPipeline = $true)]
    [ValidateNotNullOrEmpty()]
    [xml[]]$DocumentXml,

    [Parameter(
        Mandatory = $true,
        ParameterSetName = "Path",
        Position = 0,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$Path,

    [Parameter(
        Mandatory = $true,
        ParameterSetName = "LiteralPath",
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true)]
    [Alias("FullName")]
    [ValidateNotNullOrEmpty()]
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
            $DocumentXml = $paths | ForEach-Object { Get-TableauDocumentXml $_ }
        }

        $i = 0
        foreach ($xml in $DocumentXml) {
            $worksheets = @()
            $xml | Select-Xml '/workbook/worksheets/worksheet' | Select-Object -ExpandProperty Node |
                ForEach-Object {
                    $props = @{
                        "Name" = $_.Attributes['name'].Value;
                        "DisplayName" = if ($_.Attributes['caption']) { $_.Attributes['caption'].Value } else { $_.Attributes['name'].Value };

                        # TODO: Include list of referenced data sources.
                    };
                    $worksheets += New-Object PSObject -Property $props
                }

            $dashboards = @()
            $xml | Select-Xml '/workbook/dashboards/dashboard' | Select-Object -ExpandProperty Node |
                ForEach-Object {
                    # TODO: This really slows down the whole cmdlet. Find a way to make the dashboards' Worksheets property lazy evaluated.
                    $dashboardWorksheets = @()
                    $_ | Select-Xml './zones//zone' | Selec -ExpandProperty Node |
                        # Assume any zone with a @name but not a @type is a worksheet zone.
                        Where-Object { $null -eq $_.Attributes['type'] -and $null -ne $_.Attributes['name'] } |
                        ForEach-Object {
                            $zone = $_
                            $dashboardWorksheets += ($worksheets | Where-Object { $_.Name -eq $zone.Attributes['name'].Value })
                        }

                    $props = @{
                        "Name" = $_.Attributes['name'].Value;
                        "DisplayName" = if ($_.Attributes['caption']) { $_.Attributes['caption'].Value } else { $_.Attributes['name'].Value };
                        "Worksheets" = $dashboardWorksheets;
                    };
                    $dashboards += New-Object PSObject -Property $props
                }

            $dataSources = @()
            $xml | Select-Xml '/workbook/datasources/datasource' | Select-Object -ExpandProperty Node |
                ForEach-Object {
                    $props = @{
                        "Name" = $_.Attributes['name'].Value;
                        "DisplayName" = if ($_.Attributes['caption']) { $_.Attributes['caption'].Value } else { $_.Attributes['name'].Value };
                        "ConnectionType" = ($_ | Select-Xml './connection/@class').Node.Value;

                        # TODO: Include a "Connection" PSObject property with properties specific to the type of connection (i.e. file path for CSVs and server for SQL Server, etc).
                    };
                    $dataSources += New-Object PSObject -Property $props
                }

            $parameters = @()
            $xml | Select-Xml '/workbook/datasources/datasource[@name="Parameters"]/column' | Select-Object -ExpandProperty Node |
                ForEach-Object {
                    $props = @{
                        "Name" = $_.Attributes['name'].Value -ireplace "^\[|\]$", "";
                        "DisplayName" = if ($_.Attributes['caption']) { $_.Attributes['caption'].Value } else { $_.Attributes['name'].Value -ireplace "^\[|\]$", "" };
                        "DataType" = $_.Attributes['datatype'].Value;
                        "DomainType" = $_.Attributes['param-domain-type'].Value;
                        "Value" = $_.Attributes['value'].Value;
                        "ValueDisplayName" = if ($_.Attributes['alias']) { $_.Attributes['alias'].Value } else { $_.Attributes['value'].Value };

                        # TODO: For "list" parameters, include a ValueList property with all of the values and aliases.
                    };
                    $parameters += New-Object PSObject -Property $props
                }

            $props = @{
                "FileVersion" = ($xml | Select-Xml '/workbook/@version').Node.Value;
                "BuildVersion" = ($xml | Select-Xml '//comment()[1]').Node.Value -ireplace "[^\d\.]", "";
                "Parameters" = $parameters;
                "DataSources" = $dataSources;
                "Worksheets" = $worksheets;
                "Dashboards" = $dashboards;
                "DocumentXml" = $xml;
            }

            if ($paths) {
                $props['FileName'] = $paths | Select-Object -Index $i
            }

            Write-Output (New-Object PSObject -Property $props)
            $i++
        }
    }
}

# Tests for the magic zip file header.
# Inspired by http://stackoverflow.com/a/1887113/31308
function Test-ZipFile($path) {
    try {
        $stream = New-Object System.IO.StreamReader -ArgumentList @($path)
        $reader = New-Object System.IO.BinaryReader -ArgumentList @($stream.BaseStream)
        $bytes = $reader.ReadBytes(4)
        if ($bytes.Length -eq 4) {
            if ($bytes[0] -eq 80 -and
                $bytes[1] -eq 75 -and
                $bytes[2] -eq 3 -and
                $bytes[3] -eq 4) {

                return $true;
            }
        }
    }
    finally {
        if ($reader) {
            $reader.Dispose();
        }
        if ($stream) {
            $stream.Dispose();
        }
    }
    return $false;
}
