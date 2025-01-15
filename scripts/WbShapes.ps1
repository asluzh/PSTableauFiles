Add-Type -AssemblyName System.Drawing
Import-Module ./PSTableauFiles -Force
# $DebugPreference = 'SilentlyContinue' # display verbose output of the tests
# $VerbosePreference = 'Continue' # display verbose output of the tests
$files = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.twb? -Exclude invalid.*
# $files = Get-ChildItem -Recurse -Path "./tests/assets" -Filter "Drunken Gauge Design.twbx"
$result = Get-TableauFileObject -Path $files
# $result | ForEach-Object { $_.FileName }
# $result | ForEach-Object { $_.Datasources.ConnectionType } | Format-Table
# $result | Where-Object { $_.FileName -eq "Financial Markets Updated.twbx" } | ForEach-Object { $_.Parameters } | Select-Object -Property DisplayName,Range,ValueList | Format-Table
# $result | Where-Object { $_.FileName -eq "Superstore-2022-3.twb" } | ForEach-Object { $_.Thumbnails } | Format-List
$result | Where-Object { $_.FileName -eq "Another 10 hacks for BETTER user experience!.twbx" } | ForEach-Object {
    foreach ($shape in $_.Shapes) {
        $filename = [System.IO.Path]::GetFileName($shape.Name)
        [System.IO.File]::WriteAllBytes("./tests/temp/$filename", $shape.Image)
    }
}
