Import-Module ./PSTableauFiles -Force
$files = Get-ChildItem -Recurse -Path "./tests/assets" -Filter *.twbx -Exclude invalid.*
$result = Get-TableauFileStructure -Path $files -XmlAttributes
$stats = $result | Select-Object -Property FileVersion,BuildVersion,SourcePlatform,`
    @{l='File';e={[System.IO.Path]::GetFileName($_.FileName)}},
    @{l='TableauBuild';e={ if ($_.SourceBuild) { [regex]::Matches($_.SourceBuild, '\((.*)\)')[0].Groups[1].Value } else {'Unspecified'} }},
    @{l='TableauVersion';e={ if ($_.SourceBuild) { $_.SourceBuild -ireplace ' \(.*\)','' } else {'Unspecified'} }}
$stats | Group-Object -Property BuildVersion | Select-Object -Property @{l='BuildVersion';e={$_.Name}},Count,@{l='Files';e={$_.Group.File}} | Format-Table
$stats | Group-Object -Property TableauBuild | Select-Object -Property @{l='TableauBuild';e={$_.Name}},Count,@{l='Files';e={$_.Group.File}} | Format-Table
$stats | Group-Object -Property TableauVersion | Select-Object -Property @{l='TableauVersion';e={$_.Name}},Count,@{l='Files';e={$_.Group.File}} | Format-Table
$stats | Group-Object -Property SourcePlatform | Select-Object -Property @{l='SourcePlatform';e={$_.Name}},Count,@{l='Files';e={$_.Group.File}} | Format-Table
$stats | Group-Object -Property FileVersion | Select-Object -Property @{l='FileVersion';e={$_.Name}},Count,@{l='Files';e={$_.Group.File}} | Format-Table
