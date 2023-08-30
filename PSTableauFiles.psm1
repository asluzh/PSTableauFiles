# original script here https://joshua.poehls.me/2013/tableaukit-a-powershell-module-for-tableau/

function Get-TableauWorkbookXml {
<#
.SYNOPSIS
    Gets the workbook XML from a TWB or TWBX file.

.NOTES
    tbd
#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Path
    )

    begin {
        $originalCurrentDirectory = [System.Environment]::CurrentDirectory

        # System.IO.Compression.FileSystem requires at least .NET 4.5
        [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression") | Out-Null
    }

    process {
        [System.Environment]::CurrentDirectory = (Get-Location).Path
        $extension = [System.IO.Path]::GetExtension($Path)
        if ($extension -eq ".twb") {
            return [xml](Get-Content -LiteralPath $Path)
        }
        elseif ($extension -eq ".twbx") {
            $archiveStream = $null
            $archive = $null
            $reader = $null

            try {
                $archiveStream = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open)
                $archive = New-Object System.IO.Compression.ZipArchive($archiveStream)
                $twbEntry = ($archive.Entries | Where-Object { $_.FullName -eq $_.Name -and [System.IO.Path]::GetExtension($_.Name) -eq ".twb" })[0]
                $reader = New-Object System.IO.StreamReader $twbEntry.Open()

                [xml]$xml = $reader.ReadToEnd()
                return $xml
            }
            finally {
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
        }
        else {
            throw "Unknown file type. Expected a TWB or TWBX file extension."
        }
    }

    end {
        [System.Environment]::CurrentDirectory = $originalCurrentDirectory
    }
}

function Get-TableauDatasourceLiveFile {
    <#
    .SYNOPSIS
        tbd

    .NOTES
        tbd
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [string]$Path
    )

    begin {
        $originalCurrentDirectory = [System.Environment]::CurrentDirectory

        # System.IO.Compression.FileSystem requires at least .NET 4.5
        [System.Reflection.Assembly]::LoadWithPartialName("System.IO.Compression") | Out-Null
    }

    process {
        [System.Environment]::CurrentDirectory = (Get-Location).Path
        $extension = [System.IO.Path]::GetExtension($Path)
        if ($extension -eq ".tds") {
            return [xml](Get-Content -LiteralPath $Path)
        }
        elseif ($extension -eq ".tdsx") {
            $archiveStream = $null
            $archive = $null
            $reader = $null

            try {
                $archiveStream = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open)
                $archive = New-Object System.IO.Compression.ZipArchive($archiveStream)
                $fileEntry = ($archive.Entries | Where-Object { $_.FullName -eq $_.Name -and [System.IO.Path]::GetExtension($_.Name) -eq ".json" })[0]
                $reader = New-Object System.IO.StreamReader $fileEntry.Open()

                $xml = $reader.ReadToEnd()
                return $xml
            }
            finally {
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
        }
        else {
            throw "Unknown file type. Expected a TDS or TDSX file extension."
        }
    }

    end {
        [System.Environment]::CurrentDirectory = $originalCurrentDirectory
    }
}

function Update-TableauWorkbookFromXml {
    <#
    .SYNOPSIS
        Exports the workbook XML to a TWB or TWBX file.

    .PARAMETER Path
        The literal file path to export to.

    .PARAMETER WorkbookXml
        The workbook XML to export.

    .PARAMETER Update
        Whether to update the TWB inside the destination TWBX file
        if the destination file exists.

    .PARAMETER Force
        Whether to overwrite the destination TWBX file if it exists.
        By default, you will be prompted whether to overwrite any
        existing file.

    .NOTES
        tbd
    #>
        [CmdletBinding(
            SupportsShouldProcess=$true
        )]
        param(
            [Parameter(
                Position=0,
                Mandatory=$true
            )]
            [string]$Path,

            [Parameter(
                Position=1,
                Mandatory=$true,
                ValueFromPipeline=$true
            )]
            [xml]$WorkbookXml,

            [switch]$Update,
            [switch]$Force
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
                                    $WorkbookXml.Save($entryStream)
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
                        $WorkbookXml.Save($entryStream)
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

function Update-TableauDatasourceFromLive {
        <#
        .SYNOPSIS
            Exports the workbook XML to a TWB or TWBX file.

        .PARAMETER Path
            The literal file path to export to.

        .PARAMETER WorkbookXml
            The workbook XML to export.

        .PARAMETER Update
            Whether to update the TWB inside the destination TWBX file
            if the destination file exists.

        .PARAMETER Force
            Whether to overwrite the destination TWBX file if it exists.
            By default, you will be prompted whether to overwrite any
            existing file.

        .NOTES
            tbd
        #>
            [CmdletBinding(
                SupportsShouldProcess=$true
            )]
            param(
                [Parameter(
                    Position=0,
                    Mandatory=$true
                )]
                [string]$Path,

                [Parameter(
                    Position=1,
                    Mandatory=$true,
                    ValueFromPipeline=$true
                )]
                [xml]$WorkbookXml,

                [switch]$Update,
                [switch]$Force
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
                                        $WorkbookXml.Save($entryStream)
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
                            $WorkbookXml.Save($entryStream)
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

function Get-TableauFilesProperties {
<#
.SYNOPSIS
    Gets metadata information for local workbook(s).

.NOTES
    tbd
#>

    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $true,
            ParameterSetName = "Xml",
            Position = 0,
            ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [xml[]]$WorkbookXml,

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
            $WorkbookXml = $paths | ForEach-Object { Get-TableauXml $_ }
        }

        $i = 0
        foreach ($xml in $WorkbookXml) {
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
                    $_ | Select-Xml './zones//zone' | Select-Object -ExpandProperty Node |
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
                "WorkbookXml" = $xml;
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
function Test-ZipFile([string]$path) {
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

Export-ModuleMember -Function Get-TableauWorkbookXml
Export-ModuleMember -Function Get-TableauDatasourceLiveFile
Export-ModuleMember -Function Update-TableauWorkbookFromXml
Export-ModuleMember -Function Update-TableauDatasourceFromLive
Export-ModuleMember -Function Get-TableauFilesProperties
