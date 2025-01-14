Import-Module ./PSTableauFiles -Force
# $DebugPreference = 'SilentlyContinue' # display verbose output of the tests
$files = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.tds? -Exclude invalid.*
# $files = Get-ChildItem -Recurse -Path "./tests/assets" -Filter "Drunken Gauge Design.twbx"
$result = Get-TableauFileObject -Path $files
# $result | ForEach-Object { $_.FileName }
# $result | ForEach-Object { $_.Datasources.Columns.Role } | Group-Object -NoElement | Format-Table
# $result | ForEach-Object { $_.Datasources.Columns.Type } | Group-Object -NoElement | Format-Table
# $result | ForEach-Object { $_.Datasources.Columns.DataType } | Group-Object -NoElement | Format-Table
# $result | Where-Object { $_.FileName -eq "Drunken Gauge Design.twbx" } | ForEach-Object { $_.Datasources.Columns.DataType } | Group-Object -NoElement | Format-Table
# $result | Where-Object { $_.FileName -eq "Financial Markets Updated.twbx" } | ForEach-Object { $_.Datasources.Columns } | Select-Object -Property Name,DisplayName,Hidden
$result | ForEach-Object { $_.Datasources.Columns } | Select-Object -Property DisplayName,Hidden,Formula
