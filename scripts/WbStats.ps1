Import-Module ./PSTableauFiles -Force
$files = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.twbx -Exclude invalid.*
$result = Get-TableauFileObject -Path $files
# $result | ForEach-Object { [System.IO.Path]::GetFileName($_.FileName), ($_.Datasources | Measure-Object -Property Name | Select-Object -ExpandProperty Count) } #| Group-Object
# $result | ForEach-Object { $_.RepositoryLocation.Revision }
$stats = $result | Select-Object -Property `
    @{l='File';e={[System.IO.Path]::GetFileName($_.FileName)}},
    @{l='ParamCount';e={ if ($_.Parameters) { $_.Parameters | Measure-Object -Property Name | Select-Object -ExpandProperty Count } else {0} }},
    @{l='DatasourceCount';e={ if ($_.Datasources) { $_.Datasources | Measure-Object -Property Name | Select-Object -ExpandProperty Count } else {0} }},
    @{l='ColumnCount';e={ if ($_.Datasources.Columns) { $_.Datasources.Columns | Measure-Object -Property Name | Select-Object -ExpandProperty Count } else {0} }}
$stats | Group-Object -Property DatasourceCount | Select-Object -Property @{l='DatasourceCount';e={$_.Name}},Count,@{l='Files';e={$_.Group.File}} | Format-Table
$stats | Group-Object -Property ParamCount | Select-Object -Property @{l='ParamCount';e={$_.Name}},Count,@{l='Files';e={$_.Group.File}} | Format-Table
$stats | Group-Object -Property ColumnCount | Select-Object -Property @{l='ColumnCount';e={$_.Name}},Count,@{l='Files';e={$_.Group.File}} | Format-Table
$result | Where-Object { $_.FileName -like '*Hilbert*' } | Select-Object -ExpandProperty Datasources | Select-Object -ExpandProperty Columns | Select-Object -Property Type
