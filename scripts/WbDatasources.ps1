Import-Module ./PSTableauFiles -Force
# $DebugPreference = 'SilentlyContinue' # display verbose output of the tests
$files = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.twbx -Exclude invalid.*
# $files = Get-ChildItem -Recurse -Path "./tests/assets" -Filter "Drunken Gauge Design.twbx"
$result = Get-TableauFileObject -Path $files
# $result | ForEach-Object { $_.FileName }
# $result | ForEach-Object { $_.Datasources.ConnectionType } | Format-Table
# $result | Where-Object { $_.FileName -eq "Drunken Gauge Design.twbx" } | ForEach-Object { $_.Datasources.Columns.DataType } | Group-Object -NoElement | Format-Table
# $result | Where-Object { $_.FileName -eq "Drunken Gauge Design.twbx" } | ForEach-Object { $_.Datasources.Columns } | Select-Object -Property Name,DisplayName,Hidden
# $result | Where-Object { $_.FileName -eq "Drunken Gauge Design.twbx" } | ForEach-Object { $_.Datasources.Folders } | Select-Object -Property Name,FolderItems | Format-Table
# $result | Where-Object { $_.FileName -eq "Drunken Gauge Design.twbx" } | ForEach-Object { $_.Datasources.Hierarchies } | Select-Object -Property Name,HierarchyItems | Format-Table
$result | Where-Object { $_.FileName -eq "Drunken Gauge Design.twbx" } | ForEach-Object { $_.Datasources.Encodings } | Format-Table
