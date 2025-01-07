Import-Module ./PSTableauFiles -Force
$files = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.twbx -Exclude invalid.*
$result = Get-TableauFileObject -Path $files
# $result | Where-Object { $_.FileName -eq "Drunken Gauge Design.twbx" } | ForEach-Object { $_.Preferences } | Format-Table
# $result | Where-Object { $_.FileName -eq "Drunken Gauge Design.twbx" } | ForEach-Object { $_.Styles } | Format-Table
# $result | Where-Object { $_.FileName -eq "Hilbert Curves.twbx" } | ForEach-Object { $_.ColorPalettes } | Format-Table
# $result | Where-Object { $_.FileName -eq "Hilbert Curves.twbx" } | ForEach-Object { $_.Preferences }
# $result | Where-Object { $_.FileName -eq "Hilbert Curves.twbx" } | ForEach-Object { $_.ColorPalettes }
$result | Where-Object { $_.FileName -eq "Drunken Gauge Design.twbx" } | ForEach-Object { $_.Styles }
