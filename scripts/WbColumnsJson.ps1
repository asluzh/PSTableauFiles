Import-Module ./PSTableauFiles -Force
# $DebugPreference = 'SilentlyContinue' # display verbose output of the tests
$files = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.twbx -Exclude invalid.*
# $files = Get-ChildItem -Recurse -Path "./tests/assets" -Filter "Drunken Gauge Design.twbx"
$result = Get-TableauFileObject -Path $files
$result | ForEach-Object {
    @{
        FilePath = $_.FilePath;
        Parameters = $_.Parameters;
        Datasources = $_.Datasources | ForEach-Object {
            @{
                Name = $_.Name;
                Columns = $_.Columns;
                Folders = $_.Folders;
                Hierarchies = $_.Hierarchies;
                Encodings = $_.Encodings;
            }
        };

    }
} | ConvertTo-Json -Depth 10 | Out-File "./tests/output/columns.json"