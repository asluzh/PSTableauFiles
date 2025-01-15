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
$result | Where-Object { $_.FileName -eq "Financial Markets Updated.twbx" } | ForEach-Object {
    foreach ($thumbnail in $_.Thumbnails) {
        # https://stackoverflow.com/questions/12484758/how-to-create-bitmap-via-powershell
        # https://stackoverflow.com/questions/78667054/how-to-save-bitmap-created-using-system-drawing-image-back-out-to-base64
        # $bitmap = New-Object System.Drawing.Bitmap($thumbnail.Width, $thumbnail.Height)
        # $bitmap = [System.Drawing.Bitmap]::new($thumbnail.Width, $thumbnail.Height)
        # FromStream([System.IO.MemoryStream]::new($thumbnail.Image))
        # $bitmap.Save("./tests/temp/$($thumbnail.Name).png", [System.Drawing.Imaging.ImageFormat]::PNG)
    }
}
