Import-Module ./PSTableauFiles -Force
$files = Get-ChildItem -Recurse -Path "./tests/assets" -Filter "Sales Dashboard.twbx" -Exclude invalid.*
$result = Get-TableauFileStructure -Path $files -XmlPath '/workbook/datasources/datasource' -XmlElements -XmlAttributes
$result | Select-Object -ExpandProperty Attributes | Format-Table
$result | Select-Object -ExpandProperty Elements | Format-Table
