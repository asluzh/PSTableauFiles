Import-Module ./PSTableauFiles -Force
$DebugPreference = 'SilentlyContinue' # display verbose output of the tests
$VerbosePreference = 'Continue' # display verbose output of the tests
$files = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.twbx -Exclude invalid.*
# $files = Get-ChildItem -Recurse -Path "./tests/assets" -Filter "Drunken Gauge Design.twbx"
$result = Get-TableauFileObject -Path $files
# $result | ForEach-Object { $_.FileName }
# $result | ForEach-Object { $_.Datasources.ConnectionType } | Format-Table
# $result | Where-Object { $_.FileName -eq "Financial Markets Updated.twbx" } | ForEach-Object { $_.Parameters } | Select-Object -Property DisplayName,Range,ValueList | Format-Table
$result | Where-Object { $_.FileName -eq "Financial Markets Updated.twbx" } | ForEach-Object { $_.Parameters } | Format-List
