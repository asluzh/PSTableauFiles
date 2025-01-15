Import-Module ./PSTableauFiles -Force
$DebugPreference = 'Continue' # display verbose output of the tests
# $files = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.twb? -Exclude invalid.*
$files = Get-ChildItem -Recurse -Path "./tests/assets" -Filter "Financial Markets Updated.twbx"
$result = Get-TableauFileObject -Path $files
# $result | ForEach-Object { $_.FileName }
# $result | ForEach-Object { $_.Datasources.Columns.Role } | Group-Object -NoElement | Format-Table
# $result | ForEach-Object { $_.Datasources.Columns.Type } | Group-Object -NoElement | Format-Table
# $result | ForEach-Object { $_.Datasources.Columns.DataType } | Group-Object -NoElement | Format-Table
# $result | Where-Object { $_.FileName -eq "Drunken Gauge Design.twbx" } | ForEach-Object { $_.Datasources.Columns.DataType } | Group-Object -NoElement | Format-Table
# $result | Where-Object { $_.FileName -eq "Financial Markets Updated.twbx" } | ForEach-Object { $_.Datasources.Columns } | Select-Object -Property Name,DisplayName,Hidden
$result | Where-Object { $_.FileName -eq "Financial Markets Updated.twbx" } | ForEach-Object { $_.Datasources.Columns } | Select-Object -Property DisplayName,Formula,DisplayFormula | Where-Object Formula | Format-List
