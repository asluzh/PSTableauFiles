Import-Module ./PSTableauFiles -Force
$files = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.twbx -Exclude invalid.*
$result = Get-TableauFileObject -Path $files
$result | Where-Object {$_.FileName -like "*hilbert*"} | ForEach-Object { $_.Datasources.Columns } | Select-Object -Property Name,DisplayName,Formula
# $result | ForEach-Object { $_.Datasources.Columns.Role } | Group-Object -NoElement | Format-Table
