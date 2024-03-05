# Get-TableauDocumentObject

## SYNOPSIS
Get Tableau Document Object

## SYNTAX

### Xml
```
Get-TableauDocumentObject [-DocumentXml] <XmlDocument[]> [-ProgressAction <ActionPreference>]
 [<CommonParameters>]
```

### Path
```
Get-TableauDocumentObject [-Path] <String[]> [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

### LiteralPath
```
Get-TableauDocumentObject -LiteralPath <String[]> [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
Returns metadata information for local workbook(s).

## EXAMPLES

### Example 1
```powershell
PS C:\> {{ Add example code here }}
```

{{ Add example description here }}

## PARAMETERS

### -DocumentXml
tbd

```yaml
Type: XmlDocument[]
Parameter Sets: Xml
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Path
tbd

```yaml
Type: String[]
Parameter Sets: Path
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -LiteralPath
tbd

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: FullName

Required: True
Position: Named
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### -ProgressAction
{{ Fill ProgressAction Description }}

```yaml
Type: ActionPreference
Parameter Sets: (All)
Aliases: proga

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES
Inspired by https://joshua.poehls.me/2013/tableaukit-a-powershell-module-for-tableau/

## RELATED LINKS
